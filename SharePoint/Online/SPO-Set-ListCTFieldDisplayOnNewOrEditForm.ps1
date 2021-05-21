# ======================================================================================================
# SPO-Set-ListCTFieldDisplayOnNewOrEditForm.ps1
# Use this script to update hide or show fields with SharePoint List/Library New and/or Edit forms.
# ------------------------------------------------------------------------------------------------------
# Created By:    Siôn Lewis
# Created Date:  21/05/2021
# Modified By:   Siôn Lewis
# Modified Date: 21/05/2021
# ------------------------------------------------------------------------------------------------------
# Instructions:  1. Update the variables at the top of the script before running in your environment.
#                2. Update the function calls under the MAIN comment.
# ------------------------------------------------------------------------------------------------------
# Note: Modern SharePoint forms have made hiding or showing fields much easier, via the UI. If, however, 
#       form has previously been modified using "old" techniques, then new option will not be available. 
# ======================================================================================================


# ------------------------------------------------------------------------------------------------------
# Edit these variables.
# ------------------------------------------------------------------------------------------------------
[string]$siteUrl  	= "https://tenant.sharepoint.com/sites/Example";
[string]$listTitle 	= "Site Request";
[string]$contentTypeId  = "0x0100SJL24BE2E8AA04F6B836B47A46956836A0079A016644482412E90F2D5860D9BD1E9";


# ======================================================================================================


#
# Hide or show a field within a List's or Library's NEW Form (using PnP and CSOM).
# ------------------------------------------------------------------------------------------------------
# Set-ListCTFieldDisplayOnNewForm -SiteUrl $siteUrl -ListTitle $listTitle -ContentTypeId $contentTypeId -FieldName "BCS Activity" -ShowInForm $false;
#
function Set-ListCTFieldDisplayOnNewForm {
    param (
        [string]$SiteUrl,
        [string]$ListTitle,
        [string]$ContentTypeId,
        [string]$FieldName,
        [bool]$ShowInForm
    )
    #Connect-PnPOnline -Url $SiteUrl -UseWebLogin;
    Connect-PnPOnline -Url $SiteUrl -Interactive;
	
    $cnn = Get-PnPConnection;
    $ctx = $cnn.Context;

    try {
        Write-Host ("Updating field '{0}' for content type: '{1}', associated to list: {2}..." -f $FieldName, $ContentTypeId, $ListTitle) -ForegroundColor White;

        if (!$ctx.ServerObjectIsNull.Value) { 
            #$rootWeb = $ctx.Site.RootWeb;
            $rootWeb = $ctx.Web;
            $ctx.Load($rootWeb);

            $list = $rootWeb.Lists.GetByTitle($ListTitle);
            $ctx.Load($list);

            $cTypes = $list.ContentTypes;
            $ctx.Load($cTypes);
	        
            $cType = $cTypes.GetById($ContentTypeId);
            $ctx.Load($cType);

            $fields = $cType.Fields;
            $ctx.Load($fields);

            #$field = $fields.GetByTitle($FieldName);
            $field = $fields.GetByInternalNameOrTitle($FieldName);
            $ctx.Load($field);

            $ctx.ExecuteQuery();


            $field.SetShowInNewForm($ShowInForm);
            $field.UpdateAndPushChanges($true);
            $field.Update();

            $ctx.ExecuteQuery();
            
            Write-Host ("Field: '{0}'; display in New Form: {1}.`n" -f $FieldName, $ShowInForm) -ForegroundColor Green;
        }
        else {
            Write-Host "Unable to get the SharePoint Online (CSOM) client context.`n" -ForegroundColor Red;
        }
    }
    catch {
        $errorMessage = $_.Exception.Message;
        $failedItem = $_.Exception.ItemName;
        Write-Host ("Error updating field '{0}' for content type: '{1}', associated to list: {2}... `n{3}, {4}`n" -f $FieldName, $ContentTypeId, $ListTitle, $errorMessage, $failedItem) -ForegroundColor Red;
    }
    finally {
        # Disconnect from SharePoint Online.
        $ctx.Dispose();
        Disconnect-PnPOnline;
    }
}

#
# Hide or show a field within a List's or Library's EDIT Form (using PnP and CSOM).
# ------------------------------------------------------------------------------------------------------
# Set-ListCTFieldDisplayOnEditForm -SiteUrl $siteUrl -ListTitle $listTitle -ContentTypeId $contentTypeId -FieldName "BCS Activity" -ShowInForm $false;
#
function Set-ListCTFieldDisplayOnEditForm {
    param (
        [string]$SiteUrl,
        [string]$ListTitle,
        [string]$ContentTypeId,
        [string]$FieldName,
        [bool]$ShowInForm
    )
    #Connect-PnPOnline -Url $SiteUrl -UseWebLogin;
    Connect-PnPOnline -Url $SiteUrl -Interactive;
	
    $cnn = Get-PnPConnection;
    $ctx = $cnn.Context;

    try {
        Write-Host ("Updating field '{0}' for content type: '{1}', associated to list: {2}..." -f $FieldName, $ContentTypeId, $ListTitle) -ForegroundColor White;

        if (!$ctx.ServerObjectIsNull.Value) { 
            #$rootWeb = $ctx.Site.RootWeb;
            $rootWeb = $ctx.Web;
            $ctx.Load($rootWeb);

            $list = $rootWeb.Lists.GetByTitle($ListTitle);
            $ctx.Load($list);

            $cTypes = $list.ContentTypes;
            $ctx.Load($cTypes);
	        
            $cType = $cTypes.GetById($ContentTypeId);
            $ctx.Load($cType);

            $fields = $cType.Fields;
            $ctx.Load($fields);

            #$field = $fields.GetByTitle($FieldName);
            $field = $fields.GetByInternalNameOrTitle($FieldName);
            $ctx.Load($field);

            $ctx.ExecuteQuery();

            
            $field.SetShowInEditForm($ShowInForm);
            $field.UpdateAndPushChanges($true);
            $field.Update();

            $ctx.ExecuteQuery();
            
            Write-Host ("Field: '{0}'; display in Edit Form: {1}.`n" -f $FieldName, $ShowInForm) -ForegroundColor Green;
        }
        else {
            Write-Host "Unable to get the SharePoint Online (CSOM) client context.`n" -ForegroundColor Red;
        }
    }
    catch {
        $errorMessage = $_.Exception.Message;
        $failedItem = $_.Exception.ItemName;
        Write-Host ("Error updating field '{0}' for content type: '{1}', associated to list: {2}... `n{3}, {4}`n" -f $FieldName, $ContentTypeId, $ListTitle, $errorMessage, $failedItem) -ForegroundColor Red;
    }
    finally {
        # Disconnect from SharePoint Online.
        $ctx.Dispose();
        Disconnect-PnPOnline;
    }
}


# ======================================================================================================


# ------------------------------------------------------------------------------------------------------
# MAIN: Add to or edit these function calls are pre your requirements.
# ------------------------------------------------------------------------------------------------------
Set-ListCTFieldDisplayOnNewForm -SiteUrl $siteUrl -ListTitle $listTitle -ContentTypeId $contentTypeId -FieldName "My Field" -ShowInForm $false;
Set-ListCTFieldDisplayOnEditForm -SiteUrl $siteUrl -ListTitle $listTitle -ContentTypeId $contentTypeId -FieldName "My Field" -ShowInForm $false;

