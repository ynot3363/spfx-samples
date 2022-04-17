# Script variables to, update the siteURL to match the site you would like the list created on
$siteUrl = "https://anthonyepoulin.sharepoint.com"
$listName = "Messages"
$listUrl = "lists/messages"

Try{
    Connect-PnPOnline -Url $siteUrl -Interactive

    #Create the custom footer list
    Write-Host "Creating the $($listName) list..." -ForegroundColor Cyan
    $list = New-PnPList -Title $listName -Url $listUrl -Template GenericList
    Write-Host "Updating $($listName) list description..." -ForegroundColor Cyan
    Set-PnPList -Identity $list -Description "Controls the messages to display on all modern SharePoint pages."

    #Update the title column
    Write-Host "Updating the Title column..." -ForegroundColor Cyan
    Set-PnPField -List $list -Identity "Title" -Values @{Title="Message"; Description="A brief message that you would like to convey to all the users."}
    
    # Details column
    Write-Host "Creating the Details column..." -ForegroundColor Cyan
    Add-PnPField -List $list -Type "Note" -InternalName "msg_details" -DisplayName "Details" -AddToDefaultView | Out-Null
    Write-Host "Updating the Details column description..." -ForegroundColor Cyan
    Set-PnPField -List $list -Identity "Details" -Values @{Description="Additional details you would like to convey to all users with your message."}
    
    # Link column
    Write-Host "Creating the Link column..." -ForegroundColor Cyan
    Add-PnPField -List $list -Type "URL" -InternalName "msg_link" -DisplayName "Link" -AddToDefaultView | Out-Null
    Write-Host "Updating the Link column description..." -ForegroundColor Cyan
    Set-PnPField -List $list -Identity "Link" -Values @{Description="The site, page or resource you would like to navigate the user to in regards to the message."}
    
    #Type column
    Write-Host "Creating the Type column..." -ForegroundColor Cyan
    Add-PnPField -List $list -Type "Choice" -InternalName "msg_type" -DisplayName "Type" -AddToDefaultView -Required -Choices "blocked","error","info","severeWarning","success","warning" | Out-Null
    Write-Host "Updating the Type column description..." -ForegroundColor Cyan
    Set-PnPField -List $list -Identity "Type" -Values @{Description="The type of message you would like to display to users."}

    # Publish Date column
    Write-Host "Creating the Publish Date column..." -ForegroundColor Cyan
    Add-PnPField -List $list -Type "DateTime" -InternalName "msg_publishDate" -DisplayName "Publish Date" -AddToDefaultView -Required | Out-Null
    Write-Host "Updating the Publish Date description..." -ForegroundColor Cyan
    Set-PnPField -List $list -Identity "Publish Date" -Values @{Description="The date you would like the message to appear for users"}

    # Expiration Date column
    Write-Host "Creating the Expiration Date column..." -ForegroundColor Cyan
    Add-PnPField -List $list -Type "DateTime" -InternalName "msg_expirationDate" -DisplayName "Expiration Date" -AddToDefaultView -Required | Out-Null
    Write-Host "Updating the Expiration Date description..." -ForegroundColor Cyan
    Set-PnPField -List $list -Identity "Expiration Date" -Values @{Description="The date you would like the message to disappear for users"}

    Write-Host "The $($listName) has been created successfully and all the columns have been added. The list id is: $($list.Id) and you can navigate to it by clicking: $($siteUrl)$($list.DefaultViewUrl)" -ForegroundColor Green

} Catch {
    Write-Host "Error $($_.Exception.Message)" -ForegroundColor Red
}

