<#
    This script is used to update the distribution list / email enabled security group. The email address is specified in the $groupEmailAddress variable.
    This script will use a CSV file as a source. The $source variable has the file path of the CSV, and the $file variable grabs the first csv file found with the .csv file extension. It will then import the CSV to the $data variable, and then move the file another folder. This script is assuming you have the Exchange Onine Powershell Module installed. The $owner variable needs to be the UPN of the User Principal Name of the user that the script is being ran under that is an owner of the group. This connect method will login with an interactive login prompt. 
#>

# Set variables for the owner, the folder / source, and the group email address
$owner = ""
$source = ""
$file = Get-ChildItem -Path $source -Filter "*.csv" | Select-Object -First 1
$groupEmailAddress = ""

# Check if a CSV file was found. If it found add the data of the CSV to a variable, and if it is not found recommedn to check.
if ($file) {
    # Read the contents of the CSV file
    $data = Import-Csv -Path $file.FullName
    # Moves the file to a location you specify
    $newPath = Join-Path -Path "" -ChildPath $file.Name
    Move-Item -Path $file.FullName -Destination $newPath
} else {
   Write-Host "No CSV file found. Please ensure that there is a CSV in $source. Please ensure there is a CSV file here. This script will not stop running. "
   exit
}

# Connects to Exhange Online using the UPN specified. This is an interactive login prompt
Connect-ExchangeOnline -UserPrincipalName $owner

$group = Get-DistributionGroup -Identity $groupEmailAddress

#Will prompt if you infact do want to update this group, and will update accordingly. This is under the impression that a column in the csv has the header of "email"
$group | Update-DistributionGroupMember -Members $data.email
