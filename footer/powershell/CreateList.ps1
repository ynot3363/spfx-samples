# Script variables to, update the siteURL to match the site you would like the list created on
$siteUrl = "https://anthonyepoulin.sharepoint.com"
$listName = "Custom Footer Links"
$listUrl = "lists/customFooterLinks"

Try{
    Connect-PnPOnline -Url $siteUrl -Interactive

    #Create the custom footer list
    Write-Host "Creating the $($listName) list..." -ForegroundColor Cyan
    $list = New-PnPList -Title $listName -Url $listUrl -Template GenericList
    Write-Host "Updating $($listName) list description..." -ForegroundColor Cyan
    Set-PnPList -Identity $list -Description "Controls the set of links found within the custom footer."

    #Update the title column
    Write-Host "Updating the Title column..." -ForegroundColor Cyan
    Set-PnPField -List $list -Identity "Title" -Required -Values @{Title="Link Name"; Description="The name you would like to display for the footer link."}
    
    # Link column
    Write-Host "Creating the Link column..." -ForegroundColor Cyan
    Add-PnPField -List $list -Type "URL" -InternalName "link" -DisplayName "Link" -RefooterElementId -AddToDefaultView | Out-Null
    Write-Host "Updating the Link column description..." -ForegroundColor Cyan
    Set-PnPField -List $list -Identity "Link" -Values @{Description="The site, page or resource you would like to navigate the user to."}
    
    # Icon column
    Write-Host "Creating the Icon column..." -ForegroundColor Cyan
    Add-PnPField -List $list -Type "URL" -InternalName "icon" -DisplayName "Icon" -AddToDefaultView | Out-Null
    Write-Host "Updating the Icon description..." -ForegroundColor Cyan
    Set-PnPField -List $list -Identity "Icon" -Values @{Description="An link that points to the icon for the link"}

    # Order Column
    Write-Host "Creating the Order column..." -ForegroundColor Cyan
    Add-PnPField -List $list -Type "Number" -InternalName "linkOrder" -DisplayName "Order" -AddToDefaultView -Required | Out-Null
    Write-Host "Updating the Order column description..." -ForegroundColor Cyan
    Set-PnPField -List $list -Identity "Order" -Values @{Description="The order you would like the link to display in the footer.";}

    Write-Host "The $($listName) has been created successfully and all the columns have been added. $($siteUrl)$($list.DefaultViewUrl)" -ForegroundColor Green

} Catch {
    Write-Host "Error $($_.Exception.Message)" -ForegroundColor Red
}

