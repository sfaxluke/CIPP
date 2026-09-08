
function Invoke-CIPPMigrateOneDriveShortCuts {
    <#
    .SYNOPSIS
        Migrates OneDrive root shortcuts into the Shortcuts folder.
    .DESCRIPTION
        Lists drive root children with Prefer: Include-Feature=AddToOneDrive, then PATCH-moves
        each remoteItem shortcut that is not already under Shortcuts into special/shortcuts
        (or a root Shortcuts folder if the special folder is not provisioned yet).
    .FUNCTIONALITY
        Internal
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Username,

        [Parameter(Mandatory = $true)]
        $TenantFilter,

        $Headers,

        [string]$APIName = 'Migrate OneDrive shortcuts',

        [string]$ItemId,

        [switch]$ListOnly
    )

    $PreferHeaders = @{ Prefer = 'Include-Feature=AddToOneDrive' }
    $EscapedUser = [System.Uri]::EscapeDataString($Username)
    $ListUri = "https://graph.microsoft.com/beta/users/$EscapedUser/drive/root/children?`$select=id,name,remoteItem,parentReference"

    try {
        $RootChildren = @(New-GraphGetRequest -uri $ListUri -tenantid $TenantFilter -asapp $true -extraHeaders $PreferHeaders)
    } catch {
        $ErrorMessage = Get-CippException -Exception $_
        $Normalized = $ErrorMessage.NormalizedError
        if ($Normalized -match 'itemNotFound|ResourceNotFound|404|does not have a drive|no drive') {
            $Result = "No OneDrive found for $Username"
            if (-not $ListOnly) {
                Write-LogMessage -API $APIName -headers $Headers -message $Result -Sev 'Info'
            }
            if ($ListOnly) { return @() }
            throw $Result
        }
        $Result = "Could not list OneDrive shortcuts for $Username : $Normalized"
        Write-LogMessage -API $APIName -headers $Headers -message $Result -Sev 'Error' -LogData $ErrorMessage
        throw $Result
    }

    $ToMigrate = @($RootChildren | Where-Object {
            $_.remoteItem -and
            ($_.parentReference.path -notmatch '/Shortcuts(/|$)')
        })

    if (-not [string]::IsNullOrWhiteSpace($ItemId)) {
        $ToMigrate = @($ToMigrate | Where-Object { $_.id -eq $ItemId })
        if ($ToMigrate.Count -eq 0 -and -not $ListOnly) {
            $Result = "No root OneDrive shortcut with id $ItemId found for $Username"
            Write-LogMessage -API $APIName -headers $Headers -message $Result -Sev 'Info'
            throw $Result
        }
    }

    if ($ListOnly) {
        return $ToMigrate
    }

    if ($ToMigrate.Count -eq 0) {
        $Result = "No root OneDrive shortcuts to migrate for $Username"
        Write-LogMessage -API $APIName -headers $Headers -message $Result -Sev 'Info'
        return $Result
    }

    # Resolve (or create) the Shortcuts destination folder once per run.
    $ShortcutsFolderId = $null
    try {
        $ShortcutsFolder = New-GraphGetRequest -uri "https://graph.microsoft.com/beta/users/$EscapedUser/drive/special/shortcuts?`$select=id,name" -tenantid $TenantFilter -asapp $true
        $ShortcutsFolderId = $ShortcutsFolder.id
    } catch {
        try {
            $Created = New-GraphPOSTRequest -uri "https://graph.microsoft.com/beta/users/$EscapedUser/drive/root/children" -tenantid $TenantFilter -asapp $true -body (@{
                    name                                = 'Shortcuts'
                    folder                              = @{}
                    '@microsoft.graph.conflictBehavior' = 'fail'
                } | ConvertTo-Json -Depth 5)
            $ShortcutsFolderId = $Created.id
        } catch {
            try {
                $Existing = New-GraphGetRequest -uri "https://graph.microsoft.com/beta/users/$EscapedUser/drive/root:/Shortcuts?`$select=id,name" -tenantid $TenantFilter -asapp $true
                $ShortcutsFolderId = $Existing.id
            } catch {
                $ErrorMessage = Get-CippException -Exception $_
                $Result = "Could not resolve or create the Shortcuts folder for $Username : $($ErrorMessage.NormalizedError)"
                Write-LogMessage -API $APIName -headers $Headers -message $Result -Sev 'Error' -LogData $ErrorMessage
                throw $Result
            }
        }
    }

    $Migrated = [System.Collections.Generic.List[string]]::new()
    $Failures = [System.Collections.Generic.List[string]]::new()

    foreach ($Item in $ToMigrate) {
        $ShortcutName = [string]$Item.name
        $SiteUrl = $Item.remoteItem.sharepointIds.siteUrl
        $SiteSuffix = if ($SiteUrl) { " (site $SiteUrl)" } else { '' }

        try {
            $MoveBody = @{
                parentReference = @{ id = $ShortcutsFolderId }
            } | ConvertTo-Json -Depth 5
            $null = New-GraphPOSTRequest -uri "https://graph.microsoft.com/beta/users/$EscapedUser/drive/items/$($Item.id)" -tenantid $TenantFilter -type 'PATCH' -body $MoveBody -asapp $true
            $Migrated.Add($ShortcutName)
        } catch {
            $ErrorMessage = Get-CippException -Exception $_
            $FailMsg = "Could not migrate OneDrive shortcut '$ShortcutName' for ${Username}${SiteSuffix}: $($ErrorMessage.NormalizedError)"
            Write-LogMessage -API $APIName -headers $Headers -message $FailMsg -Sev 'Error' -LogData $ErrorMessage
            $Failures.Add($FailMsg)
        }
    }

    $Summary = "Migrated $($Migrated.Count) OneDrive shortcut(s) to the Shortcuts folder for $Username"
    if ($Migrated.Count -gt 0) {
        $Summary += ": $($Migrated -join ', ')"
    }
    if ($Failures.Count -gt 0) {
        $Summary += ". Failures ($($Failures.Count)): $($Failures -join ' | ')"
        Write-LogMessage -API $APIName -headers $Headers -message $Summary -Sev 'Error'
        throw $Summary
    }

    Write-LogMessage -API $APIName -headers $Headers -message $Summary -Sev 'Info'
    return $Summary
}
