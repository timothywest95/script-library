<#
.Synopsis
   Creates a status report for a Windows server.
.DESCRIPTION
   Creates a report of info about a Windows server, including system specs, uptime, unclean shutdowns, disk space, Windows Updates, and service statuses.
   Optionally emails and/or saves the report.
   Reccomended: use Task Scheduler to run the script daily, to report on errors. Schedule a separate task monthly or weekly with SendEmailReport set to email a report, even if there are no errors.
.EXAMPLE
   Create the report. Save a copy locally. Email only an error report (if any errors are found).
   .\Get-WindowsServerStatus.ps1
.EXAMPLE
    Do not save a local copy.
   .\Get-WindowsServerStatus.ps1 -SaveReport:$false
.EXAMPLE
   Email a full report.
   .\Get-WindowsServerStatus.ps1 -SendEmailReport
.PARAMETER SendEmailReport
 Email the report
.PARAMETER SaveReport
 Save a copy of the report locally (default)
.PARAMETER EmailErrors
 Email a report for errors, regardless of if SendEmailReport is set (default)
.NOTES
  Version:        1.0
  Author:         Timothy West
  Creation Date:  2024-08-10
  Emailing tested using SMTP2GO and adapted from here: https://github.com/timothywest95/script-library/blob/main/Windows/Send-SMTPEmail.ps1
  You should run the script interactively once to save SMTP credentials, then set it up with Task Scheduler.
#>

param (
    [switch]$SendEmailReport,
    [switch]$SaveReport = $true, # By default, save a report copy locally
    [switch]$EmailErrors = $true # By default, always send emails for errors
)

#region: User-Defined Variables - CHANGE THESE

# Check Threshholds
$uptimeThresholdDays = 35
$windowsUpdateHistoryDays = 10 #How many days in the past to get Windows update history for
$diskSpaceThresholdGB = 25
$customServices = @("BlueIris") # Replace with your service names. For multiple, do e.g: @("Service2", "Service2", "Service3") 

# Logging
$reportPath = "C:\!TECH\ServerStatusReports" #Saved reports path (if $SaveReport is set)
$maxReports = "100" # Maximum number of reports to retain in reports directory

# Email Variables
$smtpUsername = "sender@mydomain.com"
$fromAddress = "sender@mydomain.com" # Usually the same as $smtpUsername
$passwordFilePath = "C:\scripts\smtp_creds.txt" # Path to the file where the encrypted password will be stored
$SMTPServer = "mail.smtp2go.com"
$Port = "2525"
$recipientAddress = "jdoe@contoso.com"
#$CCRecipientAddress = "rroe@contoso.com"
#BCCRecipientAddress = "jsmith@fabrikam.com"

#endregion

#region: Import required modules
# DON'T CHANGE ANYTHING BELOW THIS LINE

# Check for and install the PSWindowsUpdate and Send-MailKitMessage modules
$moduleNames = @("PSWindowsUpdate", "Send-MailKitMessage")

foreach ($moduleName in $moduleNames) {
    # Check if the module is already installed
    $module = Get-Module -ListAvailable -Name $moduleName

    if (-not $module) {
        Write-Host "Module '$moduleName' is not installed. Attempting to install it..."
        # Try to install the module from the PSGallery
        try {
            Install-Module -Name $moduleName -Force -Scope CurrentUser -ErrorAction Stop
        }
        catch {
            Write-Host "Failed to install module '$moduleName'. Error: $_"
            exit 1
        }
    }

    # Import the module
    try {
        Import-Module -Name $moduleName -ErrorAction Stop
    }
    catch {
        Write-Host "Failed to import module '$moduleName'. Error: $_"
        exit 1
    }
}

#endregion


#region Initialize report content

$title = "$env:COMPUTERNAME Server Health Report"

$reportContent = @"
<head>
<!-- External stylesheets will not display in email clients like Gmail, which use their own CSS, but it's nice to have if saving the HTML report locally. -->
<link rel="stylesheet" href="https://cdn.simplecss.org/simple.min.css">
</head>

<title>$title</title>

<h1>$title - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')</h1>

"@

# Initialize Errors and  Error Report
$errors = @() #Initial blank array of errors.

$errorSubject = "WARNING! Errors found on $hostname"

$errorReportContent = @"
<title>$errorSubject</title>
"@

#endregion

#region Get System Info
$hostname = $env:COMPUTERNAME

# Make and Model
$system = Get-CimInstance -ClassName Win32_ComputerSystem
$make = $system.Manufacturer
$model = $system.Model

# Get Memory (Total Physical Memory in GB)
$memory = [math]::round((Get-CimInstance -ClassName Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum).Sum / 1GB, 2)

# Get CPU Info
$cpu = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1 -Property Name

# Get OS Info
$os = Get-CimInstance -ClassName Win32_OperatingSystem
$osVersion = $os.Caption + " " + $os.Version
$osArchitecture = $os.OSArchitecture

# Build HTML Table
$systemInfo = @"
<table border='1'>
    <tr><td>Hostname</td><td>$hostname</td></tr>
    <tr><td>Make</td><td>$make</td></tr>
    <tr><td>Model</td><td>$model</td></tr>
    <tr><td>Memory</td><td>$memory GB</td></tr>
    <tr><td>CPU</td><td>$($cpu.Name)</td></tr>
    <tr><td>Operating System</td><td>$osVersion ($osArchitecture)</td></tr>
</table>
"@

# Add System Info to report content
$reportContent += "`r`n<h2>System Info</h2><p>$systemInfo</p>"

#endregion

#region Storage Space
$reportContent += "<h2>Storage Report</h2>"
$diskReport = ""
$diskAlert = $false
$disks = Get-PSDrive -PSProvider FileSystem

# Filter out optical (CD/DVD) drives
$OpticalDrives = Get-CimInstance -Class Win32_CDROMDrive
$disksToTest = ($disks.Root).TrimEnd('\') | Where-Object {$OpticalDrives.Drive -notcontains $_}

# Start the HTML table
$diskReport += "<table border='1' cellpadding='5' cellspacing='0'>"
$diskReport += "<tr><th>Drive</th><th>Total Size (GB)</th><th>Used Space (GB)</th><th>Free Space (GB)</th><th>Status</th></tr>"

foreach ($diskToTest in $disksToTest) {
    $disk = Get-PSDrive -Name ($diskToTest).TrimEnd(':')

    # Calculate disk sizes
    $totalSizeGB = [math]::round($disk.Used/1GB + $disk.Free/1GB, 2)
    $usedSpaceGB = [math]::round($disk.Used/1GB, 2)
    $freeSpaceGB = [math]::round($disk.Free/1GB, 2)
    
    # Initialize the status column
    $status = "OK"
    $statusStyle = ""

    # Check for low disk space
    if ($freeSpaceGB -lt $diskSpaceThresholdGB) {
        $status = "Free space on $diskToTest is below $diskSpaceThresholdGB GB"
		$message = "Alert: $status."
		$reportContent += "<p style='color:red;'>$message</p>"
		Write-Host "$message"
        $statusStyle = "style='color:red;'"
        $diskAlert = $true
        $errors += "$message"
    }
    
    # Add a row to the HTML table
    $diskReport += "<tr>"
    $diskReport += "<td>$($disk.Name)</td>"
    $diskReport += "<td>$totalSizeGB</td>"
    $diskReport += "<td>$usedSpaceGB</td>"
    $diskReport += "<td>$freeSpaceGB</td>"
    $diskReport += "<td $statusStyle>$status</td>"
    $diskReport += "</tr>"
}

# End the HTML table
$diskReport += "</table>"

# Add disk report to the main report content
$reportContent += "$diskReport"

#endregion

#region Uptime Check
$lastBootUpTime = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
$uptime = New-TimeSpan -Start $lastBootUpTime -End (Get-Date)
$uptimeDays = (New-TimeSpan -Start $lastBootUpTime -End (Get-Date)).Days
$uptimeAlert = $uptimeDays -gt $uptimeThresholdDays

# Create Uptime table
$uptimeReport = ""
$uptimeReport += "<table border='1' cellpadding='5' cellspacing='0'>"
$uptimeReport += "<tr><td>Last Boot Time</td><td>$lastBootUpTime</td></tr><tr><td>Uptime</td><td>$($uptime.Days) days, $($uptime.Hours) hours, $($uptime.Minutes) minutes</td></tr>"
$uptimeReport += "</table>"

$reportContent += "`r`n<h2>Uptime</h2>"

if ($uptimeAlert) {
    $message = "Alert: Uptime exceeds threshold of $uptimeThresholdDays days"
    Write-Host "$message"
	$errors += $message
    $reportContent += "<p style='color:red;'>$message</p>"
}

$reportContent += "$uptimeReport"

#endregion

#region Recent Startups and Unclean Shutdowns
$recentBoots = Get-WinEvent -FilterHashtable @{LogName="System"; Id=6005,6006} -MaxEvents 10 | ConvertTo-Html "TimeCreated" -Fragment
$uncleanShutdowns = Get-WinEvent -FilterHashtable @{LogName="System"; Id=6008} -MaxEvents 5 | ConvertTo-Html "Message" -Fragment
$reportContent += "`r`n<h2>Recent Startups</h2>$recentBoots"
$reportContent += "<h2>Unclean Shutdowns</h2>$uncleanShutdowns"
#endregion

#region Windows Update History
$updateHistory = Get-WUHistory -MaxDate (Get-Date).AddDays(-$windowsUpdateHistoryDays)
$updateList = $updateHistory | Select Date,Result,KB,Title | ConvertTo-Html -Fragment

$reportContent += "`r`n<h2>Windows Update History</h2><h3>Recent Updates</h3>$updateList"

# List failed updates. The Sort-Object below sorts first by KB, then by date descending to show repeated failures for the same KB
$failedUpdateList = $updateHistory | Where-Object {$_.Result -eq 'Failed'} | Sort-Object -Property @{Expression={$_.KB}}, @{Expression={$_.Date} ;Descending=$true} | Select Date,Result,KB,Title | ConvertTo-Html -Fragment
$reportContent += "<h3>Recent Failed Updates</h3>$failedUpdateList"

#endregion

#region Service Checks
$reportContent += "<h2>Service Status</h2>"
$serviceReport = ""
$serviceReport += "<table border='1' cellpadding='5' cellspacing='0'>"
$serviceReport += "<tr><th>Service Name</th><th>Status</th></tr>"

foreach ($serviceName in $customServices) {
    $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
    if ($service) {
        $serviceStatus = $service.Status
        if ($serviceStatus -ne "Running") {
            $serviceReport += "<tr><td>$serviceName</td><td><span style='color:red;'>$serviceStatus</span></td></tr>"
			$message = "Alert: Service $serviceName is not running."
			$reportContent += "<p style='color:red;'>$message</p>"
			Write-Host "$message"
	        $errors += $message
        }
        else {
            $serviceReport += "<tr><td>$serviceName</td><td>$serviceStatus</td></tr>"
        }
    } else {
        $serviceReport += "<tr><td>$serviceName</td><td><span style='color:red;'>Not Found</span></td></tr>"
    }
}

$serviceReport += "</table>"
$reportContent += "$serviceReport"

#endregion

#region Error Report
if ($errors) {
	Write-Host "WARNING! Error conditions found."
    $errorReportContent += "<p style='color:red;'>Warning! Errors found.</p>`n<table border='1'>"
	# Loop through the errors array and add to the error report as a table
	foreach ($errorMessage in $errors) {
	    $errorReportContent += "<tr><td>$errorMessage</td></tr>"
	}
    $errorReportContent += "</table>"
    # Add the errors to the main report
	$reportContent += "`r`n<h2>Error Report</h2>$errorReportContent"
}
else {
	Write-Host "No error conditions found."
	$errorReportContent += "<p>No errors found.</p>"
	$reportContent += "`r`n<h2>Error Report</h2><p>No errors found.</p>"
}

#endregion


#region Save Report
if ($SaveReport) {
    if (-Not (Test-Path -Path $reportPath)) {
        Write-Host "Warning: report directory $reportPath not found. Not saving reports."
        break
    }
    # Get the timestamp for the report
    $datestamp = (Get-Date).ToString("yyyyMMdd_HHmmss")

    # Define the log file name
    $reportFileName = "$hostname" + "_status_" + "$datestamp" + ".html"
    $reportFilePath = Join-Path -Path $reportPath -ChildPath $reportFileName

    # Save the output to the log file
    $reportContent | Out-File -FilePath $reportFilePath
	Write-Host "Saved report to $reportFilePath"

    # Rotate logs - keep only $maxReports number of files
    $reportFiles = Get-ChildItem -Path $reportPath -Filter "$hostname*.html" | Sort-Object LastWriteTime -Descending

    if ($reportFiles.Count -gt $maxReports) {
        $filesToDelete = $reportFiles | Select-Object -Skip $maxReports
        foreach ($file in $filesToDelete) {
            Remove-Item -Path $file.FullName -Force
        }
    }
}

#endregion

#region Email Sending

# Will send up to two emails: a full report if SendEmailReport is set, and/or an error report if EmailErrors is set

$emailsToSend = @()

if ($SendEmailReport) {
    $emailsToSend += @{
        Subject      = $title
        HTMLBody     = $reportContent
    }
}

if ($errors -and $EmailErrors) {
    $emailsToSend += @{
        Subject      = $errorSubject
        HTMLBody     = $errorReportContent
    }
}

if ($emailsToSend.Count -gt 0) {
    # SMTP credentials. If credentials haven't been saved, will prompt user to save them.
    # Check if the password file exists
    if (-Not (Test-Path -Path $passwordFilePath)) {
        # Prompt the user to enter the password and save it securely to a file
        $securePassword = Read-Host -Prompt "SMTP password not found. Enter SMTP password. This will be stored for future use." -AsSecureString
        $securePassword | ConvertFrom-SecureString | Out-File -FilePath $passwordFilePath
        Write-Host "Password has been saved securely. The stored password will be used for future script runs."
    } else {
        # Read and decrypt the password from the file
        $securePassword = Get-Content -Path $passwordFilePath | ConvertTo-SecureString
    }


    # Authentication ([System.Management.Automation.PSCredential], optional)
    $Credential = [System.Management.Automation.PSCredential]::new("$smtpUsername", $securePassword)

    # Sender ([MimeKit.MailboxAddress] http://www.mimekit.net/docs/html/T_MimeKit_MailboxAddress.htm, required)
    $From = [MimeKit.MailboxAddress]"$fromAddress"

    # Recipient list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, required)
    $RecipientList = [MimeKit.InternetAddressList]::new()
    $RecipientList.Add([MimeKit.InternetAddress]"$recipientAddress")

    # CC list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
    #$CCList = [MimeKit.InternetAddressList]::new()
    #$CCList.Add([MimeKit.InternetAddress]"$CCRecipientAddress")

    # BCC list ([MimeKit.InternetAddressList] http://www.mimekit.net/docs/html/T_MimeKit_InternetAddressList.htm, optional)
    #$BCCList = [MimeKit.InternetAddressList]::new()
    #$BCCList.Add([MimeKit.InternetAddress]"$BCCRecipientAddress")

    # Loop through each email to send and prepare parameters
    foreach ($email in $emailsToSend) {
        $Parameters = @{
            "UseSecureConnectionIfAvailable" = $UseSecureConnectionIfAvailable
            "Credential"                     = $Credential
            "SMTPServer"                     = $SMTPServer
            "Port"                           = $Port
            "From"                           = $From
            "RecipientList"                  = $RecipientList
            "CCList"                         = $CCList
            "BCCList"                        = $BCCList
            "Subject"                        = $email.Subject
            "TextBody"                       = $null
            "HTMLBody"                       = $email.HTMLBody
        }

        # Send emails
        Send-MailKitMessage @Parameters
        [string[]]$emailSubjectString = $Parameters.Subject
		Write-Host "Sent email with subject: $emailSubjectString"
    }
}
#endregion
