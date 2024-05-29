# Import the Microsoft.Graph module
Import-Module Microsoft.Graph

# Define the required scopes for Microsoft Graph API
$scopes = @("User.Read.All", "Group.Read.All", "Directory.Read.All")

# Connect to Microsoft Graph
Connect-MgGraph -Scopes $scopes

# Define the output directory and file paths
$outputDir = "C:\azure_ad_import"
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir
}
$usersCsvPath = "$outputDir\Users.csv"
$groupsCsvPath = "$outputDir\Groups.csv"
$computersCsvPath = "$outputDir\Computers.csv"

# Function to escape and quote fields containing special characters
function Escape-Field($field) {
    if ($field -match '[,"\r\n]') {
        $field = $field -replace '"', '""'
        return '"' + $field + '"'
    }
    return $field
}

# Validate UUID format
function Is-ValidUUID($uuid) {
    return $uuid -match '^[a-fA-F0-9]{8}\-[a-fA-F0-9]{4}\-[a-fA-F0-9]{4}\-[a-fA-F0-9]{4}\-[a-fA-F0-9]{12}$'
}

# Fetch and export users
$users = Get-MgUser -All
$userData = @()
foreach ($user in $users) {
    $manager = (Get-MgUserManager -UserId $user.Id -ErrorAction SilentlyContinue).Id
    $memberOf = (Get-MgUserMemberOf -UserId $user.Id).Id -join ";"
    if (-not (Is-ValidUUID $user.Id)) {
        Write-Host "Invalid UUID: $($user.Id)" -ForegroundColor Red
        continue
    }
    $userData += [PSCustomObject]@{
        UUID = $user.Id
        Username = $user.UserPrincipalName
        Email = $user.Mail
        Description = $user.JobTitle
        ManagerUUID = $manager
        memberOf = Escape-Field($memberOf)
    }
}
# Export without headers
$userData | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Out-File -FilePath $usersCsvPath -Encoding utf8

# Fetch and export groups
$groups = Get-MgGroup -All
$groupData = @()
foreach ($group in $groups) {
    $memberOf = (Get-MgGroupMemberOf -GroupId $group.Id).Id -join ";"
    if (-not (Is-ValidUUID $group.Id)) {
        Write-Host "Invalid UUID: $($group.Id)" -ForegroundColor Red
        continue
    }
    $groupData += [PSCustomObject]@{
        UUID = $group.Id
        GroupName = $group.DisplayName
        Description = $group.Description
        memberOf = Escape-Field($memberOf)
    }
}
# Export without headers
$groupData | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Out-File -FilePath $groupsCsvPath -Encoding utf8

# Fetch and export computers (devices)
$computers = Get-MgDevice -All
$computerData = @()
foreach ($computer in $computers) {
    if (-not (Is-ValidUUID $computer.Id)) {
        Write-Host "Invalid UUID: $($computer.Id)" -ForegroundColor Red
        continue
    }
    $computerData += [PSCustomObject]@{
        UUID = $computer.Id
        Name = $computer.DisplayName
        FQDN = $computer.DeviceOSType
        Description = $computer.DeviceTrustType
        memberOf = ""
    }
}
# Export without headers
$computerData | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Out-File -FilePath $computersCsvPath -Encoding utf8

# Disconnect from Microsoft Graph
Disconnect-MgGraph

Write-Host "Users, Groups, and Computers have been successfully exported to $outputDir"
