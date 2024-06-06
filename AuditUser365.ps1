<#
    This script is for auditing users in Office/Microsoft 365 given a users UPN or Object ID. \
    The script will pull a users 
#>

#   Initialize Date variable with correct format
$date = Get-Date -Format "MM-dd-yy"
# Enter UPN  of SharePoint Admin
$sharePointAdmin = ""
# Enter domain here that you want to use. This is used to connect to SharePoint Online Service
$domain = ""

<#
    This section connects to Graph using the Powershell SDK, and SharePoint Online Service . You will need to do some digging on getting the correct scopes. 
    Please see this for more information: https://learn.microsoft.com/en-us/powershell/microsoftgraph/get-started?view=graph-powershell-1.0#determine-required-permission-scopes
#>
Connect-MgGraph
Connect-SPOService -Url "https://${domain}-admin.sharepoint.com"

#Add Admin as an admin on all sites
get-spoSite -Limit All | ForEach-Object { Set-SPOUser -Site $_.Url -LoginName $sharePointAdmin -IsSiteCollectionAdmin $true }

#   Enter ObjectID or User Principal Name of user you want to search for
$ObjectId = Read-Host "Enter Object ID or User Principal Name of user that you want to check"
$user = Get-MgUser -UserId $ObjectId -Property Id,UserPrincipalName,DisplayName,Mail,AccountEnabled
$name = $user.DisplayName
$UPN = $user.UserPrincipalName

#   This portion of the script is for reporting what groups a user is in. 
$groups = Get-MgUserMemberOf -UserId $user.Id
$groupMembership = @()
foreach ($groupId in $groups.Id)
{
    $group = Get-MgGroup -GroupId $groupId
    $groupProperties = @{'Display Name'=$group.DisplayName}
    $groupMembership += New-Object -TypeName PSObject -Property $groupProperties
}

$groupMembership | Format-Table


<#  This portion of the script is for reporting what SharePoint Sites a user has access to. 
    You first need to build the array of sites, and you only need the Url property. To prevent getting an error there is a 1 second delay using Start-Sleep to ensure that you do not get blocked from the SPO Service.
    It will then loop through each site, and each user to see if the user has access to the site. The output is the display name of the user, and the Url of the site. #>
$siteURLs = Get-SPOSite -Limit All | Select-Object -ExpandProperty Url
$siteMembership = @()
foreach ($url in $siteURLs)
{
    $siteAccess = Get-SPOUser -Site $url | Where-Object {$_.DisplayName -like $user.DisplayName} | Select-Object -Property DisplayName
    Start-Sleep -Seconds 1
    
    if ($null -ne $siteAccess)
    {
        $siteProperties = @{'URL'=$url}
        $siteMembership += New-Object -TypeName PSObject -Property $siteProperties
    }
}
$siteMembership | Format-Table


# This is the file path of exports
$exportGroupsFile = "${name}_Groups_${date}.csv"
$exportSitesFile = "${name}_Sites_${date}.csv"
$exportFilePath = ""

$exportGroups = Read-Host "Do you wish to export the groups to a CSV file? Enter Y for Yes and N for No"
if ($exportGroups -eq "Y") {
    $groupMembership | Export-Csv -NoTypeInformation -Path "$exportFilePath/$exportGroupsFile"
}

$exportSites = Read-Host "Do you wish to export the sites to a CSV file? Enter Y for Yes and N for No"
if ($exportSites -eq "Y") {
    $siteMembership | Export-Csv -NoTypeInformation -Path "$exportFilePath/$exportSitesFile"
}

#Everything below this comment makes changes. Do not run the below code if you are only auditing and not make changes to users. 
$Removal = Read-Host "Do you want to disable this user and remove them from groups and SharePoint Sites. Enter Y for Yes and N for No"
if ($Removal -eq "Y") {
        foreach ($groupId in $groups.Id) #No idea if this even works right
    {
        $groupName = Get-MgGroup -GroupId $groupId -Property DisplayName
        Remove-MgGroupMemberDirectoryObjectByRef -GroupId $groupId -DirectoryObjectId $user.Id
        Write-Output "$($user.DisplayName) has been removed from the group: $($groupName.DisplayName)"
    }

    foreach ($url in $siteMembership) #This is broken and needs to be fixed
    {
        Remove-SPOUser -Site $url.URL -LoginName $UPN
        Write-Output "$($user.DisplayName) has been removed from the group: $($url.URL)"
    }
}
else {
    Write-Host "Since you are not removing anything this script is now done"
    Exit
}
