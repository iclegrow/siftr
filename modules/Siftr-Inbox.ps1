<#
.SYNOPSIS
    Siftr inbox-action helpers — categorize classified emails and move some to
    Outlook folders.  Configuration is loaded from config.json when available.

.DESCRIPTION
    Provides functions to categorize messages and move them between Outlook
    folders based on Siftr triage classifications.

    On load, the module discovers the siftr_personal directory (checking
    $env:SIFTR_PERSONAL, ~/.siftr, and OneDrive) and reads config.json from
    it when present.  Folder rules and category names honour the config;
    legacy hard-coded defaults are used when no config.json exists.

    Current back-end: Outlook COM automation.
    Designed so the back-end can be swapped (e.g. to Graph API) without
    changing the public interface.

    Public interface
    ────────────────
    Get-SiftrInboxRootMessages   Return messages from the Inbox root.
    Get-SiftrConversationRootMessages
                                Return Inbox-root messages for one conversation.
    Move-SiftrMessage           Move one message by InternetMessageId.
    Set-SiftrMessageCategories  Apply one or more Outlook categories to a message.
    Set-SiftrMessageReadState   Mark a message (+ conversation) as read/unread.
    Invoke-SiftrInboxActions    Apply category and folder rules to a batch of
                                classified messages.
    Get-SiftrFolderMap          Return the current tier → folder mapping.
    Get-SiftrCategoryMap        Return the current tier → category mapping.
    Get-SiftrPersonalPath       Return the discovered personal-data directory path.
    Get-SiftrConfig             Return the loaded config object (or $null).
    Get-SiftrRepoRoot           Return the root directory of the Siftr repository.

.NOTES
    Source this file:  . "<siftr-repo-root>\modules\Siftr-Inbox.ps1"
#>

# ═══════════════════════════════════════════════════════════════════════════════
#  CONFIGURATION — loaded from config.json, with legacy defaults
# ═══════════════════════════════════════════════════════════════════════════════

# Discover the personal-data directory (first match wins)
$script:SiftrPersonalPath = $null
$_candidatePaths = @(
    $env:SIFTR_PERSONAL,
    (Join-Path $env:USERPROFILE '.siftr'),
    (Join-Path $env:USERPROFILE '<personal-data-path>')
) | Where-Object { $_ -and (Test-Path $_) }

if ($_candidatePaths) {
    $script:SiftrPersonalPath = if ($_candidatePaths -is [array]) { $_candidatePaths[0] } else { $_candidatePaths }
}

# Load config.json if found
$script:SiftrConfig = $null
if ($script:SiftrPersonalPath) {
    $configFile = Join-Path $script:SiftrPersonalPath 'config.json'
    if (Test-Path $configFile) {
        $script:SiftrConfig = Get-Content $configFile -Raw | ConvertFrom-Json
    }
}

# Build folder rules from config (with legacy defaults)
$script:SiftrFolderRules = @{}
if ($script:SiftrConfig -and $script:SiftrConfig.actions) {
    $lp = $script:SiftrConfig.actions.lowPriority
    if ($lp -and $lp.behavior -eq 'move' -and $lp.folder) {
        $script:SiftrFolderRules['LOW PRIORITY'] = $lp.folder
    }
    $cal = $script:SiftrConfig.actions.calendar
    if ($cal -and $cal.behavior -eq 'move' -and $cal.folder) {
        $script:SiftrFolderRules['CALENDAR'] = $cal.folder
    }
} else {
    # Legacy defaults when no config.json exists
    $script:SiftrFolderRules = @{
        'LOW PRIORITY' = 'LowPri'
        'CALENDAR'     = 'Meetings'
    }
}

# Build category rules from config (with legacy defaults)
if ($script:SiftrConfig -and $script:SiftrConfig.categories) {
    $c = $script:SiftrConfig.categories
    $u = if ($c.urgent) { $c.urgent } else { 'Urgent' }
    $a = if ($c.action) { $c.action } else { 'Action' }
    $i = if ($c.inform) { $c.inform } else { 'Inform' }
    $script:SiftrCategoryRules = @{
        'URGENT ACTION'     = @($a, $u)
        'ACTION NEEDED'     = @($a)
        'PRIORITY INFORMED' = @($i, $u)
        'INFORMED'          = @($i)
    }
} else {
    # Legacy defaults
    $script:SiftrCategoryRules = @{
        'URGENT ACTION'     = @('Action', 'Urgent')
        'ACTION NEEDED'     = @('Action')
        'PRIORITY INFORMED' = @('Inform', 'Urgent')
        'INFORMED'          = @('Inform')
    }
}

function Get-SiftrFolderMap {
    <#
    .SYNOPSIS  Return the current tier-to-folder mapping.
    #>
    [PSCustomObject]$script:SiftrFolderRules
}

function Get-SiftrCategoryMap {
    <#
    .SYNOPSIS  Return the current tier-to-category mapping.
    #>
    [PSCustomObject]$script:SiftrCategoryRules
}

function Get-SiftrPersonalPath {
    <#
    .SYNOPSIS  Return the discovered personal-data directory path.
    #>
    $script:SiftrPersonalPath
}

function Get-SiftrConfig {
    <#
    .SYNOPSIS  Return the loaded config object (or $null if no config.json).
    #>
    $script:SiftrConfig
}

function Get-SiftrRepoRoot {
    <#
    .SYNOPSIS  Return the root directory of the Siftr repository.
    .DESCRIPTION
        Resolves the repo root by walking up from the directory containing
        this module file (modules\Siftr-Inbox.ps1 → parent = repo root).
    #>
    (Split-Path -Parent (Split-Path -Parent $PSCommandPath))
}

# ═══════════════════════════════════════════════════════════════════════════════
#  BACK-END: OUTLOOK COM
# ═══════════════════════════════════════════════════════════════════════════════

function _Get-OutlookInbox {
    <# Internal: return the Inbox MAPIFolder object. #>
    $ol = New-Object -ComObject Outlook.Application
    $ns = $ol.GetNamespace("MAPI")
    $ns.GetDefaultFolder(6)  # olFolderInbox
}

function _Get-BodyPreview {
    param(
        [Parameter(Mandatory)]$Item,
        [int]$MaxLength = 280
    )

    $body = ''
    if ($Item.Body) {
        $body = ($Item.Body -replace '\s+', ' ').Trim()
    }

    if ($body.Length -le $MaxLength) {
        return $body
    }

    $body.Substring(0, $MaxLength) + '...'
}

function _Is-SiftrEligibleInboxItem {
    param([Parameter(Mandatory)]$Item)

    if ($null -eq $Item) { return $false }
    if ($Item.Class -eq 43) { return $true } # olMail

    $messageClass = [string]$Item.MessageClass
    if ($messageClass -like 'IPM.Schedule.Meeting*') { return $true }

    return $false
}

function _ConvertTo-SiftrInboxRecord {
    param([Parameter(Mandatory)]$Item)

    [PSCustomObject]@{
        Subject           = $item.Subject
        ReceivedTime      = [datetime]$item.ReceivedTime
        ReceivedDateTime  = ([datetime]$item.ReceivedTime).ToString('o')
        InternetMessageId = $item.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001F')
        ConversationId    = $item.ConversationID
        BodyPreview       = _Get-BodyPreview -Item $item
        Importance        = switch ([int]$item.Importance) {
            2 { 'high' }
            0 { 'low' }
            default { 'normal' }
        }
        From = [PSCustomObject]@{
            Name    = $item.SenderName
            Address = $item.SenderEmailAddress
        }
        To = $item.To
        CC = $item.CC
        Categories = [string]$item.Categories
        EntryId = $item.EntryID
        MessageClass = [string]$item.MessageClass
        IsRead = -not $item.UnRead
    }
}

function _Get-SiftrFolderItemsSnapshot {
    param([Parameter(Mandatory)]$Folder)

    $results = [System.Collections.Generic.List[object]]::new()
    $items = $Folder.Items
    $items.Sort('[ReceivedTime]', $true)

    for ($item = $items.GetFirst(); $null -ne $item; $item = $items.GetNext()) {
        $results.Add($item)
    }

    $results
}

function _Get-SiftrConversationInboxItems {
    param(
        [Parameter(Mandatory)]$Inbox,
        [Parameter(Mandatory)][string]$ConversationId,
        [switch]$IncludeCategorized,
        [switch]$IncludeRead
    )

    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($item in (_Get-SiftrFolderItemsSnapshot -Folder $Inbox)) {
        if (-not (_Is-SiftrEligibleInboxItem -Item $item)) { continue }
        if ([string]$item.ConversationID -ne $ConversationId) { continue }
        if (-not $IncludeRead -and $item.UnRead -ne $true) { continue }
        if (-not $IncludeCategorized -and -not [string]::IsNullOrWhiteSpace($item.Categories)) { continue }
        $results.Add($item)
    }

    $results
}

function _Find-MessageByInternetId {
    <#
    .SYNOPSIS  Internal: locate a MailItem in the Inbox by InternetMessageId.
    .OUTPUTS   Outlook.MailItem or $null.
    #>
    param(
        [Parameter(Mandatory)][string]$InternetMessageId,
        [Parameter(Mandatory)]$Inbox
    )
    $escaped = $InternetMessageId.Replace("'", "''")
    $filter  = "@SQL=""http://schemas.microsoft.com/mapi/proptag/0x1035001F"" = '$escaped'"
    $Inbox.Items.Find($filter)
}

function _Normalize-SiftrTier {
    param([Parameter(Mandatory)][string]$Tier)
    ($Tier -replace '^[^\w]+', '').Trim().ToUpper()
}

function _Resolve-SiftrCategories {
    param(
        [Parameter(Mandatory)][string]$Tier,
        [AllowNull()][object[]]$RequestedCategories,
        [bool]$AllowOverride = $false
    )

    $tierClean = _Normalize-SiftrTier -Tier $Tier
    $defaultCategories = if ($script:SiftrCategoryRules.ContainsKey($tierClean)) {
        @($script:SiftrCategoryRules[$tierClean] | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    }
    else {
        @()
    }

    if (-not $AllowOverride -or $null -eq $RequestedCategories -or $RequestedCategories.Count -eq 0) {
        return $defaultCategories
    }

    $requested = @($RequestedCategories |
        ForEach-Object { [string]$_ } |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique)

    if ($requested.Count -ne $defaultCategories.Count) {
        return $defaultCategories
    }

    foreach ($category in $requested) {
        if ($category -notin $defaultCategories) {
            return $defaultCategories
        }
    }

    return $requested
}

function _Get-InboxSubfolder {
    <#
    .SYNOPSIS  Internal: get a subfolder under the Inbox, optionally creating it.
    .PARAMETER AutoCreate  When set, create the folder if it doesn't exist.
    .OUTPUTS   MAPIFolder or throws (unless AutoCreate is set).
    #>
    param(
        [Parameter(Mandatory)][string]$FolderName,
        [Parameter(Mandatory)]$Inbox,
        [switch]$AutoCreate
    )
    $folder = try { $Inbox.Folders.Item($FolderName) } catch { $null }
    if (-not $folder -and $AutoCreate) {
        $folder = $Inbox.Folders.Add($FolderName)
    }
    if (-not $folder) {
        throw "Inbox subfolder '$FolderName' not found. Run 'siftr setup' to create it, or create it manually in Outlook."
    }
    $folder
}

# ═══════════════════════════════════════════════════════════════════════════════
#  PUBLIC API
# ═══════════════════════════════════════════════════════════════════════════════

function Get-SiftrInboxRootMessages {
    <#
    .SYNOPSIS  Return Inbox-root messages for Siftr triage.
    .DESCRIPTION
        Reads directly from Outlook COM so Siftr triage scope matches the
        follow-up action scope. Only messages currently in the Inbox root are
        returned; subfolders are intentionally excluded.
    .PARAMETER Since
        Lower bound for ReceivedTime. Defaults to 24 hours ago.
    .PARAMETER Limit
        Maximum number of messages to return. Defaults to 100.
    .PARAMETER IncludeRead
        When set, include read mail too. By default only unread mail is returned.
    .PARAMETER SkipCategorized
        When set, skip messages that already have one or more Outlook categories.
    .PARAMETER IncludeSubfolders
        When set, also scan immediate child folders of the Inbox (e.g.
        LowPri, Meetings) so mails moved by prior triage are still visible.
        If -Subfolders is also provided, only those named folders are added.
    .PARAMETER Subfolders
        Optional string array of subfolder names to include (e.g. 'LowPri').
        Only used when -IncludeSubfolders is set. If omitted, all child
        folders are included.
    .OUTPUTS
        PSCustomObjects with the fields Siftr uses for classification.
    #>
    param(
        [datetime]$Since = (Get-Date).AddHours(-24),
        [int]$Limit = 100,
        [switch]$IncludeRead,
        [switch]$SkipCategorized,
        [switch]$IncludeSubfolders,
        [string[]]$Subfolders
    )

    # Outlook COM ReceivedTime is always local — ensure $Since is local too
    if ($Since.Kind -eq [System.DateTimeKind]::Utc) {
        $Since = $Since.ToLocalTime()
    }

    $inbox = _Get-OutlookInbox

    $foldersToScan = @($inbox)
    if ($IncludeSubfolders) {
        foreach ($subfolder in $inbox.Folders) {
            if ($Subfolders -and $Subfolders.Count -gt 0) {
                if ($subfolder.Name -in $Subfolders) {
                    $foldersToScan += $subfolder
                }
            } else {
                $foldersToScan += $subfolder
            }
        }
    }

    $results = [System.Collections.Generic.List[object]]::new()
    foreach ($folder in $foldersToScan) {
        foreach ($item in (_Get-SiftrFolderItemsSnapshot -Folder $folder)) {
            if ($results.Count -ge $Limit) { break }
            if (-not (_Is-SiftrEligibleInboxItem -Item $item)) { continue }
            if (-not $IncludeRead -and $item.UnRead -ne $true) { continue }
            if ($item.ReceivedTime -lt $Since) { continue }
            if ($SkipCategorized -and -not [string]::IsNullOrWhiteSpace($item.Categories)) { continue }

            $results.Add((_ConvertTo-SiftrInboxRecord -Item $item))
        }
        if ($results.Count -ge $Limit) { break }
    }

    $results
}

function Get-SiftrConversationRootMessages {
    <#
    .SYNOPSIS  Return Inbox-root messages for a single conversation.
    .DESCRIPTION
        Fetches all eligible items in the Inbox root for the supplied
        ConversationId. This is useful when Siftr wants to classify the latest
        message using sibling thread context without touching subfolders.
    .PARAMETER ConversationId
        Outlook ConversationID for the thread to inspect.
    .PARAMETER IncludeCategorized
        When set, include items that already have Outlook categories.
    .PARAMETER IncludeRead
        When set, include read mail too. By default only unread mail is returned.
    .OUTPUTS
        PSCustomObjects ordered newest-first with the same shape as
        Get-SiftrInboxRootMessages.
    #>
    param(
        [Parameter(Mandatory)][string]$ConversationId,
        [switch]$IncludeCategorized,
        [switch]$IncludeRead
    )

    $inbox = _Get-OutlookInbox
    $items = _Get-SiftrConversationInboxItems `
        -Inbox $inbox `
        -ConversationId $ConversationId `
        -IncludeCategorized:$IncludeCategorized `
        -IncludeRead:$IncludeRead

    foreach ($item in $items) {
        _ConvertTo-SiftrInboxRecord -Item $item
    }
}

function Move-SiftrMessage {
    <#
    .SYNOPSIS  Move a single email to an Inbox subfolder.
    .PARAMETER InternetMessageId  The RFC-2822 Message-ID header value.
    .PARAMETER TargetFolder       Name of the Inbox subfolder (e.g. "LowPri").
    .OUTPUTS   PSCustomObject with Moved, Subject, TargetFolder, Error fields.
    #>
    param(
        [Parameter(Mandatory)][string]$InternetMessageId,
        [Parameter(Mandatory)][string]$TargetFolder
    )

    $result = [PSCustomObject]@{
        Moved       = $false
        Subject     = $null
        TargetFolder = $TargetFolder
        Error       = $null
    }

    try {
        $inbox  = _Get-OutlookInbox
        $item   = _Find-MessageByInternetId -InternetMessageId $InternetMessageId -Inbox $inbox

        if (-not $item) {
            $result.Error = "Message not found in Inbox root"
            return $result
        }

        $result.Subject = $item.Subject
        $folder = _Get-InboxSubfolder -FolderName $TargetFolder -Inbox $inbox
        $item.Move($folder) | Out-Null
        $result.Moved = $true
    }
    catch {
        $result.Error = $_.Exception.Message
    }

    $result
}

function Set-SiftrMessageCategories {
    <#
    .SYNOPSIS  Apply one or more Outlook categories to a single email.
    .PARAMETER InternetMessageId  The RFC-2822 Message-ID header value.
    .PARAMETER Categories         Category names to apply.
    .OUTPUTS   PSCustomObject with Updated, Subject, Categories, Error fields.
    #>
    param(
        [Parameter(Mandatory)][string]$InternetMessageId,
        [Parameter(Mandatory)][string[]]$Categories
    )

    $result = [PSCustomObject]@{
        Updated    = $false
        Subject    = $null
        Categories = @()
        Error      = $null
    }

    try {
        $inbox = _Get-OutlookInbox
        $item  = _Find-MessageByInternetId -InternetMessageId $InternetMessageId -Inbox $inbox

        if (-not $item) {
            $result.Error = "Message not found in Inbox root"
            return $result
        }

        $result.Subject = $item.Subject

        $existing = @()
        if ($item.Categories) {
            $existing = $item.Categories -split '\s*,\s*' | Where-Object { $_ }
        }

        $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
        foreach ($name in @($existing) + @($Categories)) {
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            [void]$set.Add($name.Trim())
        }

        $finalCategories = @($set | Sort-Object)
        $item.Categories = $finalCategories -join ', '
        $item.Save()

        $result.Updated = $true
        $result.Categories = $finalCategories
    }
    catch {
        $result.Error = $_.Exception.Message
    }

    $result
}

function Set-SiftrMessageReadState {
    <#
    .SYNOPSIS  Mark a message (and optionally its conversation) as read or unread.
    .PARAMETER InternetMessageId  The RFC-2822 Message-ID header value.
    .PARAMETER Read               $true to mark read, $false to mark unread.
    .PARAMETER WholeConversation   When set, mark all messages in the same
                                   Inbox conversation as read/unread.
    .OUTPUTS   PSCustomObject with Updated count, ConversationId, Error fields.
    #>
    param(
        [Parameter(Mandatory)][string]$InternetMessageId,
        [bool]$Read = $true,
        [switch]$WholeConversation
    )

    $result = [PSCustomObject]@{
        Updated        = 0
        ConversationId = $null
        Error          = $null
    }

    try {
        $inbox = _Get-OutlookInbox
        $item  = _Find-MessageByInternetId -InternetMessageId $InternetMessageId -Inbox $inbox

        if (-not $item) {
            $result.Error = "Message not found in Inbox root"
            return $result
        }

        $convId = $item.ConversationID
        $result.ConversationId = $convId

        if ($WholeConversation -and $convId) {
            # Walk all Inbox items in this conversation
            foreach ($msg in $inbox.Items) {
                if ($msg.Class -ne 43) { continue }  # olMail only
                if ($msg.ConversationID -eq $convId -and $msg.UnRead -eq $Read) {
                    $msg.UnRead = -not $Read
                    $msg.Save()
                    $result.Updated++
                }
            }
        }
        else {
            if ($item.UnRead -eq $Read) {
                $item.UnRead = -not $Read
                $item.Save()
                $result.Updated++
            }
        }
    }
    catch {
        $result.Error = $_.Exception.Message
    }

    $result
}

function Invoke-SiftrInboxActions {
    <#
    .SYNOPSIS  Apply category and folder rules to a batch of classified messages.

    .DESCRIPTION
        Accepts an array of classification objects (as produced by Siftr's
        triage step) and applies Outlook categories plus optional folder moves
        based on configured rules.

        Each input object must have at least:
          - InternetMessageId  (string)
          - Tier               (string, e.g. "LOW PRIORITY", "CALENDAR")

        Optional fields:
          - Categories         (string[] override; supports multi-category items)
          - Subject, From      (for reporting)
          - ConversationId     (fan out the latest thread classification to
                                uncategorized Inbox-root siblings)
          - ReceivedDateTime   (used to choose the latest classification when
                                multiple items from one conversation are present)

    .PARAMETER Classifications
        Array of objects with InternetMessageId and Tier properties.

    .PARAMETER WhatIf
        When set, reports what would be moved without actually moving.

    .OUTPUTS  Summary object with Categorized, Moved, Skipped, Errors counts
              and Details.
    #>
    param(
        [Parameter(Mandatory)][array]$Classifications,
        [switch]$WhatIf
    )

    $summary = [PSCustomObject]@{
        Categorized = 0
        Moved   = 0
        Skipped = 0
        Errors  = 0
        Details = [System.Collections.Generic.List[PSCustomObject]]::new()
    }

    $actionQueue = [System.Collections.Generic.List[object]]::new()
    $queuedIds = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    $inbox = $null

    foreach ($group in ($Classifications | Group-Object {
        if ($null -ne $_.PSObject.Properties['ConversationId'] -and -not [string]::IsNullOrWhiteSpace([string]$_.ConversationId)) {
            "conv::$([string]$_.ConversationId)"
        }
        else {
            "msg::$([string]$_.InternetMessageId)"
        }
    })) {
        $seed = $group.Group |
            Sort-Object @{
                Expression = {
                    if ($null -eq $_.PSObject.Properties['ReceivedDateTime']) { return [datetime]::MinValue }
                    try { return [datetime]$_.ReceivedDateTime } catch { return [datetime]::MinValue }
                }
                Descending = $true
            } |
            Select-Object -First 1

        if ([string]::IsNullOrWhiteSpace([string]$seed.InternetMessageId)) {
            continue
        }

        $targets = @()
        if ($group.Name -like 'conv::*') {
            if ($null -eq $inbox) {
                $inbox = _Get-OutlookInbox
            }

            $targets = @(_Get-SiftrConversationInboxItems `
                -Inbox $inbox `
                -ConversationId ([string]$seed.ConversationId) `
                -IncludeRead)

            if ($targets.Count -eq 0) {
                $targets = @($seed)
            }
        }
        else {
            $targets = @($seed)
        }

        foreach ($target in $targets) {
            $targetId = ''
            try { $targetId = [string]$target.InternetMessageId } catch {}
            if ([string]::IsNullOrWhiteSpace($targetId)) {
                try { $targetId = [string]$target.PropertyAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001F') } catch {}
            }

            if ([string]::IsNullOrWhiteSpace($targetId)) { continue }
            if (-not $queuedIds.Add($targetId)) { continue }

            $targetSubject = $seed.Subject
            try {
                if (-not [string]::IsNullOrWhiteSpace([string]$target.Subject)) {
                    $targetSubject = [string]$target.Subject
                }
            } catch {}

            $targetReceivedDateTime = $seed.ReceivedDateTime
            try {
                if ($null -ne $target.ReceivedDateTime) {
                    $targetReceivedDateTime = $target.ReceivedDateTime
                }
            } catch {}
            try {
                if ($targetReceivedDateTime -eq $seed.ReceivedDateTime -and $null -ne $target.ReceivedTime) {
                    $targetReceivedDateTime = ([datetime]$target.ReceivedTime).ToString('o')
                }
            } catch {}

            $targetTier = $seed.Tier
            try {
                if ([string]$target.MessageClass -like 'IPM.Schedule.Meeting*') {
                    $targetTier = 'CALENDAR'
                }
            } catch {}

            $allowCategoryOverride = $false
            if ($null -ne $seed.PSObject.Properties['AllowCategoryOverride']) {
                try { $allowCategoryOverride = [bool]$seed.AllowCategoryOverride } catch { $allowCategoryOverride = $false }
            }

            $actionQueue.Add([PSCustomObject]@{
                InternetMessageId = $targetId
                Tier = $targetTier
                Categories = if ($targetTier -eq 'CALENDAR') { $null } elseif ($null -ne $seed.PSObject.Properties['Categories']) { $seed.Categories } else { $null }
                AllowCategoryOverride = $allowCategoryOverride
                Subject = $targetSubject
                ConversationId = if ($null -ne $seed.PSObject.Properties['ConversationId']) { $seed.ConversationId } else { $null }
                ReceivedDateTime = $targetReceivedDateTime
            })
        }
    }

    foreach ($msg in $actionQueue) {
        $tierClean = _Normalize-SiftrTier -Tier $msg.Tier
        $categories = @()

        $allowCategoryOverride = $false
        if ($null -ne $msg.PSObject.Properties['AllowCategoryOverride']) {
            try { $allowCategoryOverride = [bool]$msg.AllowCategoryOverride } catch { $allowCategoryOverride = $false }
        }

        $requestedCategories = @()
        if ($null -ne $msg.PSObject.Properties['Categories'] -and $msg.Categories) {
            $requestedCategories = @($msg.Categories)
        }

        $categories = @(_Resolve-SiftrCategories -Tier $msg.Tier -RequestedCategories $requestedCategories -AllowOverride:$allowCategoryOverride)

        $targetFolder = $script:SiftrFolderRules[$tierClean]
        if ($categories.Count -eq 0 -and -not $targetFolder) {
            $summary.Skipped++
            continue
        }

        if ($categories.Count -gt 0) {
            if ($WhatIf) {
                $summary.Details.Add([PSCustomObject]@{
                    Action     = 'WouldCategorize'
                    Subject    = $msg.Subject
                    Tier       = $msg.Tier
                    Categories = $categories
                    Folder     = $null
                    Error      = $null
                })
                $summary.Categorized++
            }
            else {
                $categoryResult = Set-SiftrMessageCategories `
                    -InternetMessageId $msg.InternetMessageId `
                    -Categories $categories

                $categoryAction = if ($categoryResult.Updated) {
                    'Categorized'
                }
                elseif ($categoryResult.Error -eq 'Message not found in Inbox root') {
                    'Skipped'
                }
                else {
                    'Failed'
                }

                $summary.Details.Add([PSCustomObject]@{
                    Action     = $categoryAction
                    Subject    = if ($categoryResult.Subject) { $categoryResult.Subject } else { $msg.Subject }
                    Tier       = $msg.Tier
                    Categories = $categories
                    Folder     = $null
                    Error      = $categoryResult.Error
                })

                if ($categoryResult.Updated) { $summary.Categorized++ }
                elseif ($categoryAction -eq 'Skipped') { $summary.Skipped++ }
                else                                  { $summary.Errors++ }
            }
        }

        if (-not $targetFolder) {
            continue
        }

        if ($WhatIf) {
            $summary.Details.Add([PSCustomObject]@{
                Action     = 'WouldMove'
                Subject    = $msg.Subject
                Tier       = $msg.Tier
                Categories = $categories
                Folder     = $targetFolder
                Error      = $null
            })
            $summary.Moved++
            continue
        }

        $result = Move-SiftrMessage `
            -InternetMessageId $msg.InternetMessageId `
            -TargetFolder $targetFolder

        $moveAction = if ($result.Moved) {
            'Moved'
        }
        elseif ($result.Error -eq 'Message not found in Inbox root') {
            'Skipped'
        }
        else {
            'Failed'
        }

        $summary.Details.Add([PSCustomObject]@{
            Action     = $moveAction
            Subject    = if ($result.Subject) { $result.Subject } else { $msg.Subject }
            Tier       = $msg.Tier
            Categories = $categories
            Folder     = $targetFolder
            Error      = $result.Error
        })

        if ($result.Moved) { $summary.Moved++ }
        elseif ($moveAction -eq 'Skipped') { $summary.Skipped++ }
        else                               { $summary.Errors++ }
    }

    # Print summary line
    $parts = @()
    foreach ($category in (($summary.Details | Where-Object {
        $_.Action -in 'Categorized', 'WouldCategorize'
    } | ForEach-Object { $_.Categories }) | Sort-Object -Unique)) {
        $n = ($summary.Details | Where-Object {
            $_.Action -in 'Categorized', 'WouldCategorize' -and $_.Categories -contains $category
        }).Count
        if ($n -gt 0) { $parts += "$n → $category" }
    }
    foreach ($folder in ($script:SiftrFolderRules.Values | Sort-Object -Unique)) {
        $n = ($summary.Details | Where-Object { $_.Folder -eq $folder -and $_.Action -in 'Moved','WouldMove' }).Count
        if ($n -gt 0) { $parts += "$n → $folder" }
    }
    if ($summary.Skipped -gt 0) { $parts += "$($summary.Skipped) skipped" }
    if ($summary.Errors -gt 0) { $parts += "$($summary.Errors) errors" }

    $label = if ($WhatIf) { "🏷️📦 Dry run" } else { "🏷️📦 Siftr actions" }
    $line = if ($parts.Count -gt 0) {
        "$label`: $($parts -join ', ')"
    }
    else {
        "$label`: no category or folder actions"
    }
    Write-Host $line -ForegroundColor $(if ($summary.Errors) { 'Yellow' } else { 'Green' })

    $summary
}
