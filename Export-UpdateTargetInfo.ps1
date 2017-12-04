###############################################################################
#
# Export-UpdateTargetInfo.ps1
#
# Retrieve all computer's status, which approved specific update. 
#
# Parameters:
#   WSUSServer   : One or more WSUS server names(Host, FQDN, IPAddress).
#                  You can use comma separated names if you specify multiple servers.
#   KBNumber     : A knowledge base number.
#   Architecture : (optional) You can use "x86" or "x64".
#   Format       : (optional,default=CSV) You can use "CSV" or "Console".
#   OutputPath   : (optional,default=This script's path) Export CSV path.
#
# Outputs:
#   * If you specify CSV format:
#     -> Export two CSV files (ComputerDetail.csv and UpdateSummary.csv).
#   * If you specify Console format:
#     -> Only display to console.
#
###############################################################################
Param(
    [Parameter(Mandatory=$true)]
    $WSUSServers,

    [Parameter(Mandatory=$true)]
    [int]$KBNumber,

    [string]$Architecture,
    [string]$Format="CSV",
    [string]$OutputPath=$(Split-Path -Parent $MyInvocation.MyCommand.Path)
);

if (($Format -ne "CSV") -and ($Format -ne "Console")) {
    Write-Host "-Format option must specify CSV or Console." -foregroundcolor red
    break;
}

[void][reflection.assembly]::LoadWithPartialName("Microsoft.UpdateServices.Administration")

$systemLocale = Get-WinSystemLocale | % { $_.Name }
# i18n for enum Microsoft.UpdateServices.Administration.UpdateInstallationState
$Translation_UpdateInstallationState = @{
    "en-US" = @("Unknown", "NotApplicable", "NotInstalled", "Downloaded", "Installed", "Failed", "InstalledPendingReboot");
    "ja-JP" = @("不明", "適用なし", "インストールされていません", "ダウンロード済み", "インストール済み", "失敗", "再起動の保留中");
}

$dateFormatLocale = @{
    "en-US" = "MM-dd-yyyy_HH-mm-ss";
    "ja-JP" = "yyyy-MM-dd_HH-mm-ss";
}

$outputFileDate = $((Get-Date).ToString($dateFormatLocale[$systemLocale]))

function Export-UpdateInfoPerServer($WSUSServer) {
    
    try {
        
        try {
                $WSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WSUSServer,$false,8530)
        } catch {
                $WSUS = [Microsoft.UpdateServices.Administration.AdminProxy]::GetUpdateServer($WSUSServer,$false,80)
        }
    
        Write-Host "Connect to Windows Server Update Services" -foregroundcolor green

        $ComputerScope = New-Object Microsoft.UpdateServices.Administration.ComputerTargetScope
        $ComputerScope.IncludeSubgroups = $true
    
        $UpdateScope = New-Object Microsoft.UpdateServices.Administration.UpdateScope
        $UpdateScope.TextIncludes = $KBNumber
    
        $summaries = $WSUS.GetSummariesPerUpdate($UpdateScope, $ComputerScope)
    
        $summaryStatus = @()
        foreach ($object in $summaries) {
    
            # Search an update
            $RevisionId = New-Object Microsoft.UpdateServices.Administration.UpdateRevisionId
            $RevisionId.UpdateId = $object.UpdateId
            $update = $WSUS.GetUpdate($RevisionId)
    
            # Filter by architecture
            if (($Architecture -ne "") -and ($update.Title -notmatch $Architecture)) { continue; }
    
            $myObject = New-Object -TypeName PSObject
            $myObject | add-member -type Noteproperty -Name Title -Value $update.Title
            $myObject | add-member -type Noteproperty -Name KnowledgebaseArticles -Value $update.KnowledgebaseArticles[0]
            $myObject | add-member -type Noteproperty -Name InstalledCount -Value $object.InstalledCount
            $myObject | add-member -type Noteproperty -Name InstalledPendingRebootCount -Value $object.InstalledPendingRebootCount
            $myObject | add-member -type Noteproperty -Name DownloadedCount -Value $object.DownloadedCount
            $myObject | add-member -type Noteproperty -Name NotInstalledCount -Value $object.NotInstalledCount
            $myObject | add-member -type Noteproperty -Name FailedCount -Value $object.FailedCount
            
            $summaryStatus += $myObject   


            $computersStatus = @()
            $instInfoArray =  $update.GetUpdateInstallationInfoPerComputerTarget($ComputerScope)
            
            $p = 0
            foreach ($instInfo in $instInfoArray) {

                Write-Progress -Activity "Retrieving..." -PercentComplete $($p++ / $instInfoArray.Count * 100)

                if ($instInfo.UpdateApprovalAction -ne "Install") {
                    continue
                }
    
                $targetComputer = $WSUS.GetComputerTarget($instInfo.ComputerTargetId)

                $groupList = $targetComputer.ComputerTargetGroupIds | foreach-object { $WSUS.GetComputerTargetGroup($_).Name } | Sort-Object
                
                $myObject = New-Object -TypeName PSObject
                $myObject | Add-Member -Type Noteproperty -Name FullDomainName -Value $targetComputer.FullDomainName
                $myObject | Add-Member -Type Noteproperty -Name IPAddress -Value $targetComputer.IPAddress
                $myObject | Add-Member -Type Noteproperty -Name GroupOf -Value ($groupList -join ",")
                $myObject | Add-Member -Type Noteproperty -Name LastReportedStatusTime -Value $targetComputer.LastReportedStatusTime

                $myObject | Add-Member -Type Noteproperty -Name UpdateInstallationState -Value $Translation_UpdateInstallationState[$systemLocale][$instInfo.UpdateInstallationState]
                $myObject | Add-Member -Type Noteproperty -Name UpdateApprovalAction -Value $instInfo.UpdateApprovalAction

                $computersStatus += $myObject   
            }
    
            if ($Format -eq "CSV") {
                $computersStatus | Export-CSV -append -force -encoding Default -notype $OutputPath\ComputerDetail-$($WSUSServer)-KB$($update.KnowledgebaseArticles[0])-$outputFileDate.csv
            } else {
                $computersStatus | Format-Table
            }
        }
    
        if ($Format -eq "CSV") {
            $summaryStatus | select-object Title,KnowledgebaseArticles,InstalledCount,InstalledPendingRebootCount,NotInstalledCount,DownloadedCount,FailedCount | Export-CSV -notype -encoding Default $OutputPath\UpdateSummary-$($WSUSServer)-$outputFileDate.csv
        } else {
            $summaryStatus | Format-Table
        }
    
        Write-Host "  -> complete." -foregroundcolor green
    }
    catch [Exception] {
        Write-Host $_.Exception.GetType().FullName -foregroundcolor Red
        Write-Host $_.Exception.Message -foregroundcolor Red
        continue;
    }
}

function Export-UpdateInfo($WSUSServers) {
    foreach ($WSUSServer in $WSUSServers) {
        Export-UpdateInfoPerServer($WSUSServer)
    }
}

Export-UpdateInfo ($WSUSServers)