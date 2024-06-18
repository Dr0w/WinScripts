# Import the Microsoft.Graph module
Import-Module Microsoft.Graph

# Define the required scopes for Microsoft Graph API
$scopes = @("User.Read.All", "Group.Read.All", "Directory.Read.All")

# Function to escape and quote fields containing special characters
function Escape-Field($field) {
    if ($field -match '[,"\r\n]') {
        $field = $field -replace '"', '""'
        return '"' + $field + '"'
    }
    return $field
}

# Validate UUID format (case insensitive)
function Is-ValidUUID($uuid) {
    return $uuid -match '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'
}

# Check if connection to Microsoft Graph is valid
function Test-GraphConnection {
    try {
        Get-MgUser -Top 1 -ErrorAction Stop | Out-Null
        return $true
    } catch {
        return $false
    }
}

# Initialize success flag
$success = $true

Write-Host "Starting Azure AD export procedure" -ForegroundColor Cyan

try {
    Write-Host "Initiating connection to MS Graph API (this might take some time)..." -ForegroundColor Cyan
    Connect-MgGraph -Scopes $scopes
    $connected = Test-GraphConnection
    if ($connected) {
        Write-Host "Successfully connected to MS Graph API" -ForegroundColor Green
    } else {
        throw "Failed to validate the MS Graph API connection."
    }
} catch {
    Write-Host "Failed to authenticate to Microsoft Graph. Please check your credentials and network connection." -ForegroundColor Red
    $connected = $false
    $success = $false
}

if ($connected) {
    # Define the output directory and file paths
    $outputDir = "C:\azure_ad_import"
    if (-not (Test-Path -Path $outputDir)) {
        New-Item -ItemType Directory -Path $outputDir
    }
    $usersCsvPath = "$outputDir\Users.csv"
    $groupsCsvPath = "$outputDir\Groups.csv"
    $computersCsvPath = "$outputDir\Computers.csv"

    try {
    Write-Host "Fetching list of users..." -ForegroundColor Cyan
    $users = Get-MgUser -All

    Write-Host "Fetching detailed properties for users..." -ForegroundColor Cyan
    $userData = @()

    foreach ($user in $users) {
        try {
            $userWithDetails = Get-MgUser -UserId $user.Id -Property "displayName,givenName,surname,department,jobTitle,mail,mobilePhone,officeLocation,country,companyName,city"

            # Check if userWithDetails is not null
            if ($userWithDetails) {
                $manager = (Get-MgUserManager -UserId $user.Id -ErrorAction SilentlyContinue).Id
                $memberOf = (Get-MgUserMemberOf -UserId $user.Id).Id -join ";"

                $userObject = [PSCustomObject]@{
                    UUID = $user.Id
                    Username = $user.UserPrincipalName
                    Email = $user.Mail
                    Description = $user.JobTitle
                    ManagerUUID = $manager
                    memberOf = Escape-Field($memberOf)
                    wbsn_full_name = "attr:wbsn_full_name/=/$(Escape-Field($userWithDetails.DisplayName))"
                    wbsn_department = "attr:wbsn_department/=/$(Escape-Field($userWithDetails.Department))"
                    wbsn_title = "attr:wbsn_title/=/$(Escape-Field($userWithDetails.JobTitle))"
                    wbsn_telephone_number = "attr:wbsn_telephone_number/=/$(Escape-Field($userWithDetails.MobilePhone))"
                    first_name = "attr:First Name/=/$(Escape-Field($userWithDetails.GivenName))"
                    last_name = "attr:Last Name/=/$(Escape-Field($userWithDetails.Surname))"
                }

                $userData += $userObject
            } else {
                Write-Host "User details not retrieved for $($user.UserPrincipalName)" -ForegroundColor Yellow
            }
        } catch {
            Write-Host "Error processing user $($user.UserPrincipalName): $_" -ForegroundColor Red
        }
    }

    # Export user data to CSV without headers
    $userData | ConvertTo-Csv -NoTypeInformation | Select-Object -Skip 1 | Out-File -FilePath $usersCsvPath -Encoding utf8
    Write-Host "Users exported successfully to $usersCsvPath" -ForegroundColor Green

} catch {
    Write-Host "Failed to export users. Error: $_" -ForegroundColor Red
    $success = $false
}

    if ($success -and $connected) {
        try {
            Write-Host "Exporting groups..." -ForegroundColor Cyan
            # Fetch and export groups
            $groups = Get-MgGroup -All
            $groupData = @()
            foreach ($group in $groups) {
                $memberOf = (Get-MgGroupMemberOf -GroupId $group.Id).Id -join ";"
                if (-not (Is-ValidUUID $group.Id)) {
                    Write-Host "Invalid UUID found for group: $($group.Id)" -ForegroundColor Red
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
            Write-Host "Groups exported successfully to $groupsCsvPath" -ForegroundColor Green
        } catch {
            Write-Host "Failed to export groups. Please check for potential issues in the group data retrieval process." -ForegroundColor Red
            $success = $false
        }
    }

    if ($success -and $connected) {
        try {
            Write-Host "Exporting computers (devices)..." -ForegroundColor Cyan
            # Fetch and export computers (devices)
            $computers = Get-MgDevice -All
            $computerData = @()
            foreach ($computer in $computers) {
                if (-not (Is-ValidUUID $computer.Id)) {
                    Write-Host "Invalid UUID found for computer: $($computer.Id)" -ForegroundColor Red
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
            Write-Host "Computers exported successfully to $computersCsvPath" -ForegroundColor Green
        } catch {
            Write-Host "Failed to export computers. Please check for potential issues in the computer data retrieval process." -ForegroundColor Red
            $success = $false
        }
    }

    # Disconnect from Microsoft Graph
    try {
        Disconnect-MgGraph
        Write-Host "Disconnected from Microsoft Graph" -ForegroundColor Green
    } catch {
        Write-Host "No application to sign out from." -ForegroundColor Yellow
    }
}

# Display final status message
if ($success -and $connected) {
    Write-Host "Users, Groups, and Computers have been successfully exported to $outputDir" -ForegroundColor Green
} elseif (-not $connected) {
    Write-Host "Authentication failed. No data exported." -ForegroundColor Red
} else {
    Write-Host "Export process encountered errors. Please check the logs for details." -ForegroundColor Yellow
}
