Connect-PnPOnline -Url "https://anthonyepoulin.sharepoint.com" -Interactive

$hubSiteUrl = "https://anthonyepoulin.sharepoint.com"

$hubSite = GetPnPHubSite -Identity $hubSiteUrl

if($hubSite -eq $null){
    Write-Host "The site $($hubSiteUrl) is not a hub site or could not be found. Please check the URL and try again." -ForegroundColor Red
    exit
}

$sitesInHub = Get-PnPTenantSite -Detailed | Where-Object { $_.HubSiteId -eq $hubSite.Id }

if($sitesInHub.Count -eq 0){
    Write-Host "There are no sites associated with the hub site $($hubSiteUrl). Please associate some sites and try again." -ForegroundColor Yellow
    exit
}

foreach($site in $sitesInHub){
    Write-Host "Connecting to site $($site.Url)..." -ForegroundColor Cyan
    Connect-PnPOnline -Url $site.Url -Interactive
    $componentId = "b99ce130-58f6-4147-84c6-e4d09180a357"
    $properties = "{`"siteUrl`":`"https://anthonyepoulin.sharepoint.com`",`"listGuid`":`"71DE9F3F-ADEA-4737-ABDD-5F4EB7EB73F6`",`"footerElementId`":`"aepCustomFooter`",`"cacheKey`":`"aepFooterLinks`",`"showCopyright`":true,`"companyName`":`"Anthony as a Service`"}"

    $extension = Get-PnPApplicationCustomizer -ClientSideComponentId $componentId
    if($null -eq $extension){
        Write-Host "Adding the Footer extension to site $($site.Url)..." -ForegroundColor Cyan
        Add-PnPApplicationCustomizer -ClientSideComponentId $componentId -ClientSideComponentProperties $properties
        Write-Host "The Footer extension has been added to site $($site.Url)." -ForegroundColor Green
    } else {
        Write-Host "The Footer extension is already installed on site $($site.Url)." -ForegroundColor Yellow
        Set-PnPApplicationCustomizer -ClientSideComponentId $componentId -ClientSideComponentProperties $properties
        Write-Host "The Footer extension properties have been updated on site $($site.Url)." -ForegroundColor Green
    }

    Set-PnPFooter -Enabled $false
    Write-Host "The out of the box footer has been disabled on site $($site.Url)." -ForegroundColor Green

}