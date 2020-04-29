
<#
    Script to construct team link using Microsoft Teams PowerShell module.
    https://docs.microsoft.com/en-us/powershell/teams/intro
    https://www.powershellgallery.com/packages/MicrosoftTeams/
#>
[CmdletBinding()]
param (
    [string]$outputFile = "./list.json"
)

#Connect to Microsoft Teams
$connectTeams = Connect-MicrosoftTeams

$teamsList = @()
$teamLinkTemplate = "https://teams.microsoft.com/l/team/<ThreadId>/conversations?groupId=<GroupId>&tenantId=<TenantId>"

Write-Verbose "Get Team Names and Details"

$teams = Get-Team

Write-Verbose "Teams Count is $($teams.count)"

$teams | ForEach-Object { 

    Write-Verbose "Getting details for $($_.DisplayName)"
    #Retrieve team channel General (can be replaced by another Channel if needed)
    $channel = Get-TeamChannel -GroupId $_.GroupId | Where-Object {$_.DisplayName -eq "General"} | Select-Object -First 1

    #Construct the team link
    $teamLink = $teamLinkTemplate.Replace("<ThreadId>",$channel.Id).Replace("<GroupId>",$_.GroupId).Replace("<TenantId>",$connectTeams.TenantId)

    $team = @{  
        id = $_.GroupId
        name = $_.DisplayName
        description = $_.Description
        visibility = $_.Visibility
        url = $teamLink
        archived = $_.Archived
    }

    $teamsList += $team
}

$teamsList | ConvertTo-Json -depth 100 | Out-File $outputFile
