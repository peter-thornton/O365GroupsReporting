# A script to check the activity of Office 365 Groups and Teams and report the groups and teams that might be deleted because they're not used.
# We check the group mailbox to see what the last time a conversation item was added to the Inbox folder. 
# Another check sees whether a low number of items exist in the mailbox, which would show that it's not being used.
# We also check the group document library in SharePoint Online to see whether it exists or has been used in the last 90 days.
# And we check Teams compliance items to figure out if any chatting is happening.

# Created 29-July-2016  Tony Redmond 
# V2.0 5-Jan-2018
# V3.0 17-Dec-2018


#Updated on March 18, 2019
# cmdlet to show group members
# Get-UnifiedGroupLinks -Identity $Group.Identity -LinkType Members ---Stores as type Array, default addributes are Name and RecipientType
# 
#
#


$O365username = 'pthornton@hu.onmicrosoft.com';
$O365password = Get-Content -Path 'D:\Peter\TeamsReportingScript\O365pass.txt' | ConvertTo-SecureString -Force 
$credO365 = New-Object -typename System.Management.Automation.PSCredential -argumentlist $O365username, $O365password;
$SPOURL = 'https://hu-admin.sharepoint.com'

Function Test-EXO {
  $Sessions = Get-PSSession | Where-Object {($_.ComputerName -eq 'outlook.office365.com') -and ($_.Availability -eq 'available')}
  If (!($Sessions)) {
    If ($LogFile) {
      Add-LogEntry -LogType 'INFO' -LogText 'Attempting to connect to Exchange Online'
    }
    Else {
      Write-Host 'Attempting to connect to Exchange Online'
    }
    $exoRetry = 0
    $exoRetryMax = 20
    $exoRetrySeconds = 30
    $EXOsession = New-PSSession `
    -ConfigurationName Microsoft.Exchange `
    -ConnectionUri https://outlook.office365.com/powershell-liveid `
    -Authentication Basic `
    -Credential $credO365 `
    -AllowRedirection
    While (($EXOSession.Availability -ne 'available') -and ($exoRetry -le $exoRetryMax)) {
      $exoRetry ++
      If ($LogFile) {
        Add-LogEntry -LogType 'WARNING' -LogText "Attempting to connect to EXO.  Retry $exoRetry of $exoRetryMax in $exoRetrySeconds seconds..."
      }
      Else {
        Write-Host "Attempting to connect to EXO.  Retry $exoRetry of $exoRetryMax in $exoRetrySeconds seconds..." -ForegroundColor Yellow
      }
      Start-Sleep -Seconds $exoRetrySeconds
      $EXOsession = New-PSSession `
      -ConfigurationName Microsoft.Exchange `
      -ConnectionUri https://outlook.office365.com/powershell-liveid `
      -Authentication Basic `
      -Credential $credO365 `
      -AllowRedirection
    }
    If ($EXOSession.Availability -eq 'available') {
      Import-PSSession `
      -Session $ExoSession `
      -AllowClobber
      If ($LogFile) {
        Add-LogEntry -LogType 'SUCCESS'-LogText 'Connected to Exchange Online'
      }
      Else {
        Write-Host 'Connected to Exchange Online' -ForegroundColor Green
      }
    }
    Else {
      If ($LogFile) {
        Add-LogEntry -LogType 'ERROR' -LogText 'Unable to connect to Exchange Online'
      }
      Else {
        Write-Host 'Unable to connect to Exchange Online' -ForegroundColor Red
      }
      Exit
    }
  }
  Else {
    If ($LogFile) {
      Add-LogEntry -LogType 'SUCCESS' -LogText 'Connected to Exchange Online'
    }
    Else {
      Write-Host 'Connected to Exchange Online' -ForegroundColor Green
    }
  }
}


Function Test-MSO {
  If (!(Get-Module -Name MSOnline)) {
    Write-Host ''
    Add-LogEntry -LogType 'INFO' -LogText 'Attempting to connect to Office 365 ...'
    Import-Module MSOnline
    If (!($credO365)) {
      $global:credO365 = Get-Credential -Message 'Enter your O365 UPN and password'
    }
    $Error.Clear()
    Connect-MsolService -Credential $credO365
    If ($Error) {
      Add-LogEntry -LogType 'ERROR' -LogText 'Could not connect to O365'
    }
    Else {
      Add-LogEntry -LogType 'SUCCESS' -LogText 'Connected to O365'
    }
  }
}

Function Test-SPO {
  If (!(Get-Module -Name Microsoft.Online.SharePoint.PowerShell)) {
    Write-Host ''
    Add-LogEntry -LogType 'INFO' -LogText 'Attempting to connect to Sharepoint Online ...'
    Import-Module Microsoft.Online.SharePoint.PowerShell
    If (!($credO365)) {
      $global:credO365 = Get-Credential -Message 'Enter your O365 UPN and password'
    }
    $Error.Clear()
    Connect-SPOService -Credential $credO365 -Url $SPOURL
    If ($Error) {
      Add-LogEntry -LogType 'ERROR' -LogText 'Could not connect to Sharepoint Online'
    }
    Else {
      Add-LogEntry -LogType 'SUCCESS' -LogText 'Connected to Sharepoint Online'
    }
  }
}

Function Test-Teams {
  If (!(Get-Module -Name MicrosoftTeams)) {
    Write-Host ''
    Add-LogEntry -LogType 'INFO' -LogText 'Attempting to connect to Microsoft Teams ...'
    Import-Module MicrosoftTeams
    If (!($credO365)) {
      $global:credO365 = Get-Credential -Message 'Enter your O365 UPN and password'
    }
    $Error.Clear()
    Connect-MicrosoftTeams -Credential $credO365
    If ($Error) {
      Add-LogEntry -LogType 'ERROR' -LogText 'Could not connect to Microsoft Teams'
    }
    Else {
      Add-LogEntry -LogType 'SUCCESS' -LogText 'Connected to Microsoft Teams'
    }
  }
}


#Connect all services
Test-EXO
Test-MSO
Test-SPO
Test-Teams

# Check that we are connected to Exchange Online
Write-Host "Checking that prerequisite PowerShell modules are loaded..."
Try { $OrgName = (Get-OrganizationConfig).Name }
   Catch  {
      Write-Host "Your PowerShell session is not connected to Exchange Online."
      Write-Host "Please connect to Exchange Online using an administrative account and retry."
      Break }

# And check that we're connected to SharePoint Online as well
Try { $SPOCheck = (Get-SPOTenant -ErrorAction SilentlyContinue ) }
   Catch {
      Write-Host "Your PowerShell session is not connected to SharePoint Online."
      Write-Host "Please connect to SharePoint Online using an administrative account and retry."
      Break }

# And finally the Teams module
Try { $TeamsCheck = (Get-Team) }
    Catch {
      Write-Host "Please connect to the Teams PowerShell module before proceeeding."
      Break }
       
# OK, we seem to be fully connected to both Exchange Online and SharePoint Online...
Write-Host "Checking for Obsolete Office 365 Groups in the tenant:" $OrgName

# Setup some stuff we use
$WarningDate = (Get-Date).AddDays(-90)
$WarningEmailDate = (Get-Date).AddDays(-365)
$Today = (Get-Date)
$Date = $Today.ToShortDateString()
$TeamsGroups = 0
$TeamsEnabled = $False
$ObsoleteSPOGroups = 0
$ObsoleteEmailGroups = 0
$Report = @()
$ReportFile = "c:\temp\GroupsActivityReport1.html"
$CSVFile = "c:\temp\GroupsActivityReport1.csv"
$htmlhead="<html>
	   <style>
	   BODY{font-family: Arial; font-size: 8pt;}
	   H1{font-size: 22px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H2{font-size: 18px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   H3{font-size: 16px; font-family: 'Segoe UI Light','Segoe UI','Lucida Grande',Verdana,Arial,Helvetica,sans-serif;}
	   TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
	   TH{border: 1px solid #969595; background: #dddddd; padding: 5px; color: #000000;}
	   TD{border: 1px solid #969595; padding: 5px; }
	   td.pass{background: #B7EB83;}
	   td.warn{background: #FFF275;}
	   td.fail{background: #FF2626; color: #ffffff;}
	   td.info{background: #85D4FF;}
	   </style>
	   <body>
           <div align=center>
           <p><h1>Office 365 Groups and Teams Activity Report</h1></p>
           <p><h3>Generated: " + $date + "</h3></p></div>"
		
# Get a list of all Office 365 Groups in the tenant
Write-Host "Extracting list of Office 365 Groups for checking..."
$Groups = Get-UnifiedGroup -ResultSize Unlimited | Sort-Object DisplayName
#$Groups = Get-UnifiedGroup -ResultSize 100 | Sort-Object DisplayName
# And create a hash table of Teams
$TeamsList = @{}
Get-Team | ForEach { $TeamsList.Add($_.GroupId, $_.DisplayName) }

Write-Host "Processing" $Groups.Count "groups"
# Progress bar
$ProgDelta = 100/($Groups.count)
$CheckCount = 0
$GroupNumber = 0

# Main loop
ForEach ($G in $Groups) {
   $GroupNumber++
   $GroupStatus = $G.DisplayName + " ["+ $GroupNumber +"/" + $Groups.Count + "]"
   Write-Progress -Activity "Checking group" -Status $GroupStatus -PercentComplete $CheckCount
   $CheckCount += $ProgDelta
   $ObsoleteReportLine = $G.DisplayName
   $SPOStatus = "Normal"
   $SPOActivity = "Document library in use"
   $NumberWarnings = 0
   $NumberofChats = 0
   $TeamChatData = $Null
   $TeamsEnabled = $False
   $LastItemAddedtoTeams = "No chats"
   $MailboxStatus = $Null
# Check who manages the group
  $ManagedBy = $G.ManagedBy
  Write-Host 'testing - value of managedby at line 241' $ManagedBy
If ([string]::IsNullOrWhiteSpace($ManagedBy) -and [string]::IsNullOrEmpty($ManagedBy)) {
     $ManagedBy = "No owners"
     Write-Host 'testing - value of managedby at line 244' $ManagedBy
     Write-Host $G.DisplayName "has no group owners!" -ForegroundColor Red}
  Else {
    $ManagedBy = (Get-Mailbox -Identity $G.ManagedBy[0]).PrimarySmtpAddress
    <#
     $ManagedBy = (Get-Mailbox -Identity $G.ManagedBy).p
     Write-Host 'testing - value of managedby at line 246' $ManagedBy[0]
     $TempManagedBy = $ManagedBy[0]
     $ManagedByDetails = Get-ADUser -Filter{mail -eq $TempManagedBy} -Properties * | Select-Object harvardEduADRoleAffiliateDesc0
     Write-host 'Managed By:' $ManangedBy ' Details:' $ManangedByDetails  
   #>


    }
   

#Check for members of group
$Members = Get-UnifiedGroupLinks -Identity $G.Alias -LinkType Members
  
# Fetch information about activity in the Inbox folder of the group mailbox  
   $Data = (Get-MailboxFolderStatistics -Identity $G.Alias -IncludeOldestAndNewestITems -FolderScope Inbox)
   $LastConversation = $Data.NewestItemReceivedDate
   $NumberConversations = $Data.ItemsInFolder
   $MailboxStatus = "Normal"
  
   If ($Data.NewestItemReceivedDate -le $WarningEmailDate) {
      Write-Host "Last conversation item created in" $G.DisplayName "was" $Data.NewestItemReceivedDate "-> Obsolete?"
      $ObsoleteReportLine = $ObsoleteReportLine + " Last conversation dated: " + $Data.NewestItemReceivedDate + "."
      $MailboxStatus = "Group Inbox Not Recently Used"
      $ObsoleteEmailGroups++
      $NumberWarnings++ }
   Else
      {# Some conversations exist - but if there are fewer than 20, we should flag this...
      If ($Data.ItemsInFolder -lt 20) {
           $ObsoleteReportLine = $ObsoleteReportLine + " Only " + $Data.ItemsInFolder + " conversation item(s) found."
           $MailboxStatus = "Low number of conversations"
           $NumberWarnings++}
      }

# Loop to check SharePoint document library
   If ($G.SharePointDocumentsUrl -ne $Null) {
      $SPOSite = (Get-SPOSite -Identity $G.SharePointDocumentsUrl.replace("/Shared Documents", ""))
      $AuditCheck = $G.SharePointDocumentsUrl + "/*"
      $AuditRecs = 0
      $AuditRecs = (Search-UnifiedAuditLog -RecordType SharePointFileOperation -StartDate $WarningDate -EndDate $Today -ObjectId $AuditCheck -SessionCommand ReturnNextPreviewPage)
      If ($AuditRecs -eq $null) {
         #Write-Host "No audit records found for" $SPOSite.Title "-> Potentially obsolete!"
         $ObsoleteSPOGroups++   
         $ObsoleteReportLine = $ObsoleteReportLine + " No SPO activity detected in the last 90 days."  
         }          
       
       }

   Else
       {
# The SharePoint document library URL is blank, so the document library was never created for this group
         #Write-Host "SharePoint has never been used for the group" $G.DisplayName 
        $ObsoleteSPOGroups++  
        $ObsoleteReportLine = $ObsoleteReportLine + " SPO document library never created." 
       }
# Report to the screen what we found - but only if something was found...   
  If ($ObsoleteReportLine -ne $G.DisplayName)
     {
     Write-Host $ObsoleteReportLine 
     }
# Generate the number of warnings to decide how obsolete the group might be...   
  If ($AuditRecs -eq $Null) {
       $SPOActivity = "No SPO activity detected in the last 90 days"
       $NumberWarnings++ }
   If ($G.SharePointDocumentsUrl -eq $Null) {
       $SPOStatus = "Document library never created"
       $NumberWarnings++ }
  
    $Status = "Pass"
    If ($NumberWarnings -eq 1)
       {
       $Status = "Warning"
    }
    If ($NumberWarnings -gt 1)
       {
       $Status = "Fail"
    } 

# If Team-Enabled, we can find the date of the last chat compliance record
If ($TeamsList.ContainsKey($G.ExternalDirectoryObjectId) -eq $True) {
      $TeamsEnabled = $True
      $TeamChatData = (Get-MailboxFolderStatistics -Identity $G.Alias -IncludeOldestAndNewestItems -FolderScope ConversationHistory)
      If ($TeamChatData.ItemsInFolder[1] -ne 0) {
          $LastItemAddedtoTeams = $TeamChatData.NewestItemReceivedDate[1]
          $NumberofChats = $TeamChatData.ItemsInFolder[1] 
          If ($TeamChatData.NewestItemReceivedDate -le $WarningEmailDate) {
            Write-Host "Team-enabled group" $G.DisplayName "has only" $TeamChatData.ItemsInFolder[1] "compliance record(s)" }
          }
      }

# Generate a line for this group for our report
    $ReportLine = [PSCustomObject][Ordered]@{
          GroupName           = $G.DisplayName
          ManagedBy           = $ManagedBy
          ManagedByDetails    = $ManagedByDetails
          Members             = $Members.Name -join ";"
          MemberCount         = $G.GroupMemberCount
          ExternalGuests      = $G.GroupExternalMemberCount
          Description         = $G.Notes
          MailboxStatus       = $MailboxStatus
          TeamEnabled         = $TeamsEnabled
          LastChat            = $LastItemAddedtoTeams
          NumberChats         = $NumberofChats
          LastConversation    = $LastConversation
          NumberConversations = $NumberConversations
          SPOActivity         = $SPOActivity
          SPOStatus           = $SPOStatus
          NumberWarnings      = $NumberWarnings
          Status              = $Status}
# And store the line in the report object
   $Report += $ReportLine     
#End of main loop
}
# Create the HTML report
$PercentTeams = ($TeamsList.Count/$Groups.Count)
$htmlbody = $Report | ConvertTo-Html -Fragment
$htmltail = "<p>Report created for: " + $OrgName + "
             </p>
             <p>Number of groups scanned: " + $Groups.Count + "</p>" +
             "<p>Number of potentially obsolete groups (based on document library activity): " + $ObsoleteSPOGroups + "</p>" +
             "<p>Number of potentially obsolete groups (based on conversation activity): " + $ObsoleteEmailGroups + "<p>"+
             "<p>Number of Teams-enabled groups    : " + $TeamsList.Count + "</p>" +
             "<p>Percentage of Teams-enabled groups: " + ($PercentTeams).tostring("P") + "</body></html>"	
$htmlreport = $htmlhead + $htmlbody + $htmltail
$htmlreport | Out-File $ReportFile  -Encoding UTF8

# Summary note 
Write-Host $ObsoleteSPOGroups "obsolete group document libraries and" $ObsoleteEmailGroups "obsolete email groups found out of" $Groups.Count "checked"
Write-Host "Summary report available in" $ReportFile "and CSV file saved in" $CSVFile
$Report | Export-CSV -NoTypeInformation $CSVFile