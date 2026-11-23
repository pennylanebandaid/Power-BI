<#
.SYNOPSIS
    Documents every item in one or more Power BI workspaces: folder, workspace
    permissions, refresh schedule, data sources, and the tables and queries
    each semantic model or paginated report is built from. One CSV, one row
    per item.

.DESCRIPTION
    Full guide, with a sample of the output: README.md next to this script
    (power-bi-api/workspace-inventory/ in the work-notes repository).

    Give it one or more workspace GUIDs (-WorkspaceId) or a file of them
    (-WorkspaceIdFile). Run it signed in with either:
    - an account with a role in the workspace (Admin, Member or Contributor
      gets everything), or
    - a Fabric administrator account.

    As a Fabric administrator, it reads the workspace with an admin metadata
    scan (the scanner API), which works without a role in the workspace.
    Otherwise it reads the workspace with the normal workspace APIs.

    A semantic model's tables and Power Query (M) code come from the scan
    when the tenant settings below are on; otherwise from the model's
    definition (Fabric "Get Semantic Model Definition"), which gives the same
    information.

    Output: one file in -OutputFolder, PowerBI-inventory-<date>-<time>.csv,
    with one row per item and everything about the item on that row:

        WorkspaceName, WorkspaceId     the workspace
        ItemName, ItemType, ItemId     semantic model, interactive report,
                                       paginated report, dashboard, App (one row
                                       for a workspace's published app), or
                                       Gateway (one row per gateway the items use:
                                       type, version, members, settings in
                                       QueryDetail; its roles in ItemAccess)
        Folder                         workspace folder path, e.g. Finance/Monthly
        Owner, Created, ModifiedBy,    who owns the item (model: configured by;
        Modified                       report: created by; app: published by) and
                                       when it was created and last changed
        Endorsement                    Promoted or Certified (and by whom)
        SensitivityLabel               the sensitivity label's ID (names are in
                                       Microsoft Purview)
        Description                    the item's description
                                       With a standard account these come only
                                       for semantic models (reports: description
                                       only; dashboards: none)
        InApp                          reports and dashboards: the name of the
                                       workspace app they're published in, or
                                       No. Fabric admin account only
        WorkspacePermissions           who has which workspace role
        ItemAccess                     App rows: everyone the app is shared with
                                       and their permissions (the audience each
                                       person is in isn't available from the API).
                                       Other rows: who has access to the item other than
                                       through a workspace role (shared directly,
                                       or a Viewer given more than Read), and with
                                       which permissions. All item types with a
                                       Fabric admin account; semantic models only
                                       with a standard account (needs Admin or
                                       Member on the workspace)
        RefreshSchedule                semantic models: days, times, time zone.
                                       Blank for reports (see their model's row);
                                       paginated reports have no refresh
        Security                       semantic models: each row-level security
                                       role with its filters, any object-level
                                       security (hidden tables and columns), and
                                       its members
        DataSources                    servers, databases, URLs, dataflows, and for
                                       each source its gateway and connection
                                       (name, authentication, privacy level). On a
                                       Gateway row: each connection with its users
                                       and the items that use it
                                       (blank for an interactive report that uses a
                                       semantic model: see the model's own row)
        TableCount, QueryCount,        counts (see below) and their breakdown
        QueryDetail
        Tables                         each table, shared query or parameter (or
                                       paginated report dataset) with its kind
                                       and connectors
        Queries                        the full M, DAX or SQL of each of them

    Signed in as a Fabric administrator, it also saves a tenant settings
    snapshot, PowerBI-tenant-settings-<date>-<time>.csv: every Admin portal
    tenant setting (enabled or not, which security groups it applies to or
    excludes, who it's delegated to, its extra options) and every capacity,
    domain and workspace override, one row each.

    A workspace that can't be read gets one "Workspace error" row with the
    reason, and the script carries on with the next workspace.

    Email subscriptions are read by a separate script,
    Get-ReportSubscriptions.ps1, because the admin call for them is limited
    to 200 an hour. Cells longer
    than Excel can show (32,767 characters) are cut, with a note.

    Counts:

    - Semantic model
        TableCount = tables in the finished model (not counting the hidden
                     auto date/time tables Power BI adds).
        QueryCount = Power Query queries: tables loaded by Power Query, plus
                     shared queries (staging/helper queries and functions that
                     aren't loaded as tables). Parameters and DAX calculated
                     tables are listed in QueryDetail but not counted.
    - Paginated report
        TableCount = QueryCount = datasets in the .rdl.
    - Interactive reports and dashboards have no queries of their own;
      QueryDetail names the semantic model they use.

    Calls made: a POST to start the admin scan (admin/workspaces/getInfo)
    when you are a Fabric administrator, a POST per model to read its
    definition when needed (getDefinition), and GETs for everything else.
    Both POSTs only read: nothing in Power BI is created, changed, refreshed
    or run.

    Requirements:
    - A role in the workspace, or the Fabric administrator role.
    - Model tables and queries come from one of:
        a) the scan, if these tenant settings (Admin API settings) are on:
           "Enhance admin APIs responses with detailed metadata"
           "Enhance admin APIs responses with DAX and mashup expressions"
        b) the model definition, which needs Contributor or higher on the
           workspace. Used automatically when the scan has no tables.
    - For folder names, refresh schedules and paginated report datasets:
      a role in the workspace (Contributor or higher for the .rdl download).
      Without one, folders show as IDs, refresh schedules fall back to the
      capacity's refreshables (Premium or Fabric capacity only) and paginated
      report counts show "Unknown".

.PARAMETER WorkspaceId
    One or more workspace GUIDs (the part after /groups/ in the workspace
    URL), separated by commas.

.PARAMETER WorkspaceIdFile
    A text file with workspace GUIDs, one per line. Any GUID on a line is
    used, so a line can also hold the workspace name or a note. Lines
    starting with # are skipped. Can be used together with -WorkspaceId.

.PARAMETER OutputFolder
    Where to save the CSV file. Defaults to the current folder. Put the path
    in quotes if it contains a space: -OutputFolder "C:\pbi\New Folder"

.PARAMETER FabricApiUrl
    Base URL of the Fabric REST API, used for workspace folders and model definitions.
    The default is the commercial cloud. Folder names need a role in the
    workspace; without one, the admin item list gives each item's folder ID.

.PARAMETER SkipModelDefinition
    Don't read model definitions when the scan returns no tables. Their
    counts then show "Unknown".

.PARAMETER SkipRdlDownload
    Don't download paginated reports. Their dataset counts show "Unknown".

.PARAMETER NoAdminScan
    Don't try the admin scan; read the workspace with your workspace role.
    Use this when signed in with an account that isn't a Fabric administrator.
    (Without it, the script tries the scan and switches over on its own.)

.PARAMETER SkipTenantSettings
    Don't save the tenant settings snapshot. (It's only saved when signed in
    as a Fabric administrator, and never with -NoAdminScan.)

.PARAMETER PassThru
    Also return the rows as objects, for example to pipe into Out-GridView.

.EXAMPLE
    Connect-PowerBIServiceAccount
    .\Get-WorkspaceInventory.ps1 -WorkspaceId 11111111-2222-3333-4444-555555555555

.EXAMPLE
    # Several workspaces
    .\Get-WorkspaceInventory.ps1 -WorkspaceId 11111111-2222-3333-4444-555555555555, 66666666-7777-8888-9999-000000000000 -NoAdminScan

.EXAMPLE
    # A list of workspaces in a text file, one GUID per line
    .\Get-WorkspaceInventory.ps1 -WorkspaceIdFile C:\pbi\workspaces.txt -OutputFolder C:\pbi\inventory -NoAdminScan

.EXAMPLE
    .\Get-WorkspaceInventory.ps1 -WorkspaceId 11111111-2222-3333-4444-555555555555 -OutputFolder C:\pbi\inventory

.EXAMPLE
    # Signed in with your standard account (a role in the workspace, not a Fabric administrator)
    .\Get-WorkspaceInventory.ps1 -WorkspaceId 11111111-2222-3333-4444-555555555555 -NoAdminScan
#>
# PositionalBinding off: a stray word (such as an unquoted folder name with a
# space) is reported as an error instead of being taken as another parameter.
[CmdletBinding(PositionalBinding = $false)]
param(
    [Parameter(Position = 0)] [guid[]] $WorkspaceId,
    [string] $WorkspaceIdFile,
    [string] $OutputFolder = (Get-Location).Path,
    [ValidatePattern('^https://')] [string] $FabricApiUrl = "https://api.fabric.microsoft.com/v1",
    [switch] $SkipRdlDownload,
    [switch] $SkipModelDefinition,
    [switch] $NoAdminScan,
    [switch] $SkipTenantSettings,
    [switch] $PassThru
)

$ErrorActionPreference = "Stop"
$ScriptVersion = "2026-11-23.14"

# On any error, show where it happened, so the problem can be found from the output alone.
trap {
    $inv = $_.InvocationInfo
    Write-Host ""
    Write-Host "ERROR (script version $ScriptVersion): $($_.Exception.Message)" -ForegroundColor Red
    if ($inv -and $inv.ScriptLineNumber) { Write-Host "  at line $($inv.ScriptLineNumber): $("$($inv.Line)".Trim())" -ForegroundColor Red }
    Write-Host "  call stack:`n$($_.ScriptStackTrace)" -ForegroundColor Red
    break
}
Write-Host "Get-WorkspaceInventory version $ScriptVersion (PowerShell $($PSVersionTable.PSVersion))"

# ---------------------------------------------------------------------------
# Sign-in check
# ---------------------------------------------------------------------------

if (-not (Get-Command Invoke-PowerBIRestMethod -ErrorAction SilentlyContinue)) {
    throw "The MicrosoftPowerBIMgmt module isn't installed. See power-bi-api/power-bi-api-powershell-vscode.md in the work-notes repository."
}
try { $null = Get-PowerBIAccessToken } catch {
    throw "You aren't signed in. Run Connect-PowerBIServiceAccount and try again."
}

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

function Invoke-PbiGet([string] $Url) {
    Invoke-PowerBIRestMethod -Url $Url -Method Get -ErrorAction Stop | ConvertFrom-Json
}

# Runs a GET and returns $null instead of failing. The reason goes in $script:LastError.
function Try-PbiGet([string] $Url) {
    try { $script:LastError = $null; Invoke-PbiGet $Url }
    catch { $script:LastError = Get-ShortError $_; $null }
}

function Get-ShortError($ErrorRecord) {
    $msg = "$($ErrorRecord.Exception.Message)"
    # Power BI and Fabric send the reason as JSON in the response body
    $detail = $null
    $body = "$(if ($ErrorRecord.ErrorDetails) { $ErrorRecord.ErrorDetails.Message })"
    if ($body -match '^\s*\{') {
        try {
            $j = $body | ConvertFrom-Json
            $e = $(if (Get-Prop $j "error") { $j.error } else { $j })
            $code = $(if (Get-Prop $e "errorCode") { $e.errorCode } else { Get-Prop $e "code" })
            $text = Get-Prop $e "message"
            $detail = Join-NonEmpty ": " @($code, $text)
        } catch { }
    }
    $short = $(if ($msg -match '401|Unauthorized') { "not signed in or not allowed (401)" }
               elseif ($msg -match '403|Forbidden') { "no permission (403)" }
               elseif ($msg -match '404|NotFound')  { "not found (404)" }
               else { ($msg -split "`n")[0].Trim() })
    $(if ($detail) { "$short - $detail" } else { $short })
}

function Get-Prop($Object, [string] $Name) {
    if ($null -ne $Object -and $Object.PSObject.Properties[$Name]) { $Object.$Name } else { $null }
}

function Join-NonEmpty([string] $Separator, $Values) {
    (@($Values) | Where-Object { $_ -and "$_".Trim() } | Select-Object -Unique) -join $Separator
}

# ---- Fabric API (folders, model definitions, gateways, connections) ----

# GET with paging. On "too many requests" (429) it waits and tries again.
function Get-FabricPaged([string] $Url, [string] $ListProperty = "value") {
    $next = $Url
    while ($next) {
        $page = $null
        for ($try = 1; -not $page; $try++) {
            try {
                $headers = @{ Authorization = (Get-PowerBIAccessToken -AsString) }
                $page = Invoke-RestMethod -Uri $next -Method Get -Headers $headers -ErrorAction Stop
            } catch {
                if ($try -lt 4 -and "$($_.Exception.Message)" -match '429|Too ?Many') {
                    Write-Host "  Fabric API busy (429); waiting 60 seconds ..." -ForegroundColor Yellow
                    Start-Sleep -Seconds 60
                } else { throw }
            }
        }
        $list = Get-Prop $page $ListProperty
        if ($null -eq $list -and $ListProperty -eq "value") { $list = Get-Prop $page "Overrides" }
        foreach ($v in @($list)) { if ($v) { $v } }
        $next = Get-Prop $page "continuationUri"
    }
}

# ---- Gateways and connections ----
# Read once per run. Each account only sees the gateways and connections it
# has a role on (gateway Admin or Connection creator, connection User or
# Owner); the rest are shown by ID.

$script:gatewayById    = $null
$script:connectionById = $null
$script:gatewayUse     = @{}   # gateway ID -> connection ID -> list of "workspace / item"
$script:gatewaySource  = @{}   # connection ID -> its data source as text
function Initialize-Gateways {
    if ($null -ne $script:gatewayById) { return }
    $script:gatewayById = @{}; $script:connectionById = @{}
    try { foreach ($g in Get-FabricPaged "$FabricApiUrl/gateways") { $script:gatewayById["$($g.id)".ToLowerInvariant()] = $g } }
    catch { Write-Warning "Gateways couldn't be listed: $(Get-ShortError $_). Gateways are shown by ID only." }
    try { foreach ($c in Get-FabricPaged "$FabricApiUrl/connections") { $script:connectionById["$($c.id)".ToLowerInvariant()] = $c } }
    catch { Write-Warning "Connections couldn't be listed: $(Get-ShortError $_). Connections are shown by ID only." }
    Write-Host "Gateways visible to this account: $($script:gatewayById.Count); connections: $($script:connectionById.Count)"
}

function Get-GatewayLabel([string] $GatewayId) {
    $g = $script:gatewayById["$GatewayId".ToLowerInvariant()]
    if (-not $g) { return "gateway $GatewayId" }
    $name = $(if (Get-Prop $g "displayName") { $g.displayName } else { "personal gateway $GatewayId" })
    "gateway: $name ($(Get-Prop $g 'type'))"
}

function Get-ConnectionLabel([string] $ConnectionId) {
    $c = $script:connectionById["$ConnectionId".ToLowerInvariant()]
    if (-not $c) { return $(if ($ConnectionId) { "connection $ConnectionId" }) }
    $cred = Get-Prop $c "credentialDetails"
    $sso  = Get-Prop $cred "singleSignOnType"
    $info = Join-NonEmpty ", " @((Get-Prop $cred "credentialType"), $(if ($sso -and $sso -ne "None") { "SSO $sso" }),
                                 $(if (Get-Prop $c "privacyLevel") { "privacy $($c.privacyLevel)" }))
    "connection: $($c.displayName)$(if ($info) { " ($info)" })"
}

# One data source as text. $UsedBy ("workspace / item") records which items
# use each gateway connection, for the Gateway rows.
function Format-DataSource($S, [string] $UsedBy) {
    $c = Get-Prop $S "connectionDetails"
    $where = Join-NonEmpty " / " @((Get-Prop $c "server"), (Get-Prop $c "database"), (Get-Prop $c "url"),
                                   (Get-Prop $c "path"), (Get-Prop $c "account"), (Get-Prop $c "kind"))
    $gatewayId    = "$(Get-Prop $S 'gatewayId')"
    # Workspace-API sources carry the real ID in connectionId (their datasourceId is the script's own key)
    $connectionId = "$(if ($S.PSObject.Properties['connectionId']) { $S.connectionId } else { Get-Prop $S 'datasourceId' })"
    $gw = ""
    if ($gatewayId) {
        Initialize-Gateways
        $conn = $script:connectionById[$connectionId.ToLowerInvariant()]
        if ($conn -and "$(Get-Prop $conn 'connectivityType')" -match 'Cloud') {
            $gw = " [cloud $(Get-ConnectionLabel $connectionId)]"
        } else {
            $gw = " [$(Join-NonEmpty '; ' @((Get-GatewayLabel $gatewayId), (Get-ConnectionLabel $connectionId)))]"
            $gKey = $gatewayId.ToLowerInvariant()
            if (-not $script:gatewayUse.ContainsKey($gKey)) { $script:gatewayUse[$gKey] = @{} }
            $cKey = $(if ($connectionId) { $connectionId.ToLowerInvariant() } else { "(unknown)" })
            if (-not $script:gatewayUse[$gKey].ContainsKey($cKey)) { $script:gatewayUse[$gKey][$cKey] = New-Object System.Collections.Generic.List[string] }
            if ($UsedBy -and -not $script:gatewayUse[$gKey][$cKey].Contains($UsedBy)) { $script:gatewayUse[$gKey][$cKey].Add($UsedBy) }
            if (-not $script:gatewaySource.ContainsKey($cKey)) { $script:gatewaySource[$cKey] = $(if ($where) { "$($S.datasourceType): $where" } else { "$($S.datasourceType)" }) }
        }
    }
    if ($where) { "$($S.datasourceType): $where$gw" } else { "$($S.datasourceType)$gw" }
}

# A principal from a Fabric role assignment, as text.
function Format-FabricPrincipal($P) {
    $type = Get-Prop $P "type"
    if ($type -eq "EntireTenant") { return "Entire organization" }
    $name = $(if (Get-Prop $P "displayName") { $P.displayName } else { Get-Prop $P "id" })
    $upn  = Get-Prop (Get-Prop $P "userDetails") "userPrincipalName"
    $text = $(if ($upn -and $upn -ne $name) { "$name ($upn)" } else { "$name" })
    switch ($type) { "Group" { "$text (group)" } "ServicePrincipal" { "$text (app)" } default { $text } }
}

# The connector functions an M expression calls, e.g. "Sql.Database, Web.Contents".
$mSourcePattern = '\b(Sql\.Databases?|Oracle\.Database|Odbc\.\w+|OleDb\.\w+|OData\.Feed|Web\.Contents|SharePoint\.\w+|Excel\.Workbook|Csv\.Document|File\.Contents|Folder\.Files|AnalysisServices\.Databases?|PowerPlatform\.Dataflows|PowerBI\.Dataflows|Lakehouse\.Contents|Snowflake\.Databases|Databricks\.Catalogs|AzureStorage\.\w+|GoogleBigQuery\.Database|Teradata\.Database|DB2\.Database|MySQL\.Database|PostgreSQL\.Database|SapHana\.Database|Json\.Document|Xml\.Tables|CommonDataService\.Database|Dataverse\.\w+)\s*\('
function Get-MSources([string] $Expression) {
    if (-not $Expression) { return "" }
    Join-NonEmpty ", " ([regex]::Matches($Expression, $mSourcePattern) | ForEach-Object { $_.Groups[1].Value })
}

# ---------------------------------------------------------------------------
# Item-level access (beyond workspace roles)
# ---------------------------------------------------------------------------

# "ReadReshareExplore" -> "Read, Reshare, Build" (Explore is what the service calls Build).
function Format-AccessRight([string] $Right) {
    $parts = [regex]::Matches($Right, '[A-Z][a-z]*') | ForEach-Object { if ($_.Value -eq "Explore") { "Build" } else { $_.Value } }
    $parts -join ", "
}

function Get-PrincipalKeys($P) {
    @((Get-Prop $P "graphId"), (Get-Prop $P "identifier"), (Get-Prop $P "emailAddress")) |
        Where-Object { $_ } | ForEach-Object { "$_".ToLowerInvariant() }
}

function Format-Principal($P) {
    $type  = Get-Prop $P "principalType"
    if ($type -eq "None") { return "Entire organization" }
    $name  = $(if (Get-Prop $P "displayName") { $P.displayName } else { Get-Prop $P "identifier" })
    $email = Get-Prop $P "emailAddress"
    $text  = $(if ($email -and $email -ne $name) { "$name ($email)" } else { "$name" })
    switch ($type) { "Group" { "$text (group)" } "App" { "$text (app)" } default { $text } }
}

# Who has access to an item other than through a workspace role, from a list
# of the item's users. $RightProperty is the access-right field of that list
# (reportUserAccessRight, datasetUserAccessRight, dashboardUserAccessRight).
# Workspace Viewers are listed too if the item gives them more than Read.
function Format-ItemAccess($Users, [string] $RightProperty) {
    $extra = foreach ($u in @($Users) | Where-Object { $_ }) {
        $right = "$(Get-Prop $u $RightProperty)"
        if (-not $right -or $right -eq "None") { continue }
        $role = $null
        foreach ($k in Get-PrincipalKeys $u) { if ($roleByKey.ContainsKey($k)) { $role = $roleByKey[$k]; break } }
        if (-not $role) { "$(Format-Principal $u): $(Format-AccessRight $right)" }
        elseif ($role -eq "Viewer" -and $right -ne "Read") { "$(Format-Principal $u): Viewer role; on this item: $(Format-AccessRight $right)" }
    }
    $text = Join-NonEmpty "; " @($extra | Sort-Object)
    $(if ($text) { $text } else { "None beyond workspace roles" })
}

$noAdminAccess = "Not available without a Fabric admin account"

# ---- Item details: owner, dates, endorsement, sensitivity label, description ----
# The admin scan has them for every item type; the workspace APIs have them for
# semantic models, the description for reports, and none for dashboards.

# Times are shown in UTC. Power BI returns them in UTC, but PowerShell 7 may
# turn a time ending in "Z" into local time; that is turned back here.
function Format-When($Value) {
    if (-not $Value) { return $null }
    $when = $null
    if ($Value -is [datetime]) { $when = $Value }
    else {
        $parsed = [datetime]::MinValue
        if ([datetime]::TryParse("$Value", [Globalization.CultureInfo]::InvariantCulture, [Globalization.DateTimeStyles]::RoundtripKind, [ref] $parsed)) { $when = $parsed }
        else { return "$Value" }
    }
    if ($when.Kind -eq [DateTimeKind]::Local) { $when = $when.ToUniversalTime() }
    $when.ToString("yyyy-MM-dd HH:mm") + " UTC"
}

function Get-ItemMeta($Item) {
    $e = Get-Prop $Item "endorsementDetails"
    $endorsement = Get-Prop $e "endorsement"
    if ($endorsement -eq "None") { $endorsement = $null }
    if ($endorsement -and (Get-Prop $e "certifiedBy")) { $endorsement = "$endorsement by $($e.certifiedBy)" }
    [pscustomobject]@{
        Owner            = @((Get-Prop $Item "configuredBy"), (Get-Prop $Item "createdBy"), (Get-Prop $Item "publishedBy")) | Where-Object { $_ } | Select-Object -First 1
        Created          = Format-When (@((Get-Prop $Item "createdDate"), (Get-Prop $Item "createdDateTime")) | Where-Object { $_ } | Select-Object -First 1)
        ModifiedBy       = Get-Prop $Item "modifiedBy"
        Modified         = Format-When (@((Get-Prop $Item "modifiedDateTime"), (Get-Prop $Item "lastUpdate")) | Where-Object { $_ } | Select-Object -First 1)
        Endorsement      = $endorsement
        SensitivityLabel = Get-Prop (Get-Prop $Item "sensitivityLabel") "labelId"
        Description      = Get-Prop $Item "description"
    }
}

# ---------------------------------------------------------------------------
# Workspaces to read
# ---------------------------------------------------------------------------

$workspaceIds = New-Object System.Collections.Generic.List[string]
foreach ($id in @($WorkspaceId)) { if ($id) { $workspaceIds.Add("$id") } }
if ($WorkspaceIdFile) {
    if (-not (Test-Path -LiteralPath $WorkspaceIdFile)) { throw "WorkspaceIdFile '$WorkspaceIdFile' doesn't exist." }
    # Any GUID on any line counts, so a CSV export or a list with names and notes works too.
    foreach ($line in Get-Content -LiteralPath $WorkspaceIdFile) {
        if ($line.TrimStart().StartsWith("#")) { continue }
        foreach ($m in [regex]::Matches($line, '[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}')) {
            $workspaceIds.Add($m.Value.ToLowerInvariant())
        }
    }
}
$workspaceIds = @($workspaceIds | Select-Object -Unique)
if (-not $workspaceIds.Count) { throw "No workspaces given. Use -WorkspaceId (one or more GUIDs) or -WorkspaceIdFile (a file with one GUID per line)." }
Write-Host "Workspaces to read: $($workspaceIds.Count)"

$rows      = New-Object System.Collections.Generic.List[object]   # one per item, all workspaces
$tableRows = New-Object System.Collections.Generic.List[object]   # tables/queries, joined into the item rows at the end
$failed    = New-Object System.Collections.Generic.List[object]
$script:adminScanRefused = $false
$script:appsByWorkspace = $null   # tenant app list, read once on the first admin-scanned workspace
$n = 0

foreach ($wsId in $workspaceIds) {
    $n++
    Write-Host ""
    Write-Host "=== Workspace $n of $($workspaceIds.Count): $wsId" -ForegroundColor Cyan
    try {
        $wsName = $null; $result = $null; $readBy = $null
        # ---------------------------------------------------------------------------
        # 1. Read the workspace: admin scan if you are a Fabric administrator,
        #    otherwise the normal workspace APIs (needs a role in the workspace)
        # ---------------------------------------------------------------------------

        function Get-WorkspaceByAdminScan {
            $scanUrl = "admin/workspaces/getInfo?lineage=True&datasourceDetails=True&datasetSchema=True&datasetExpressions=True&getArtifactUsers=True"
            $body = @{ workspaces = @("$wsId") } | ConvertTo-Json
            $scan = Invoke-PowerBIRestMethod -Url $scanUrl -Method Post -Body $body -ErrorAction Stop | ConvertFrom-Json
            $deadline = (Get-Date).AddMinutes(10)
            do {
                Start-Sleep -Seconds 5
                $status = Invoke-PbiGet "admin/workspaces/scanStatus/$($scan.id)"
                Write-Host "  scan status: $($status.status)"
                if ((Get-Date) -gt $deadline) { throw "the scan didn't finish in 10 minutes" }
            } while ($status.status -ne "Succeeded" -and $status.status -ne "Failed")
            if ($status.status -eq "Failed") { throw "the scan failed: $(Get-Prop $status 'error' | ConvertTo-Json -Compress)" }
            $r = Invoke-PbiGet "admin/workspaces/scanResult/$($scan.id)"
            if (-not (@($r.workspaces) | Select-Object -First 1)) { throw "the scan didn't return the workspace" }
            $r
        }

        # Builds the same shape as the scan result from the workspace APIs.
        function Get-WorkspaceByApis {
            $f = [uri]::EscapeDataString("id eq '$wsId'")
            $g = @((Invoke-PbiGet "groups?`$filter=$f").value) | Select-Object -First 1
            if (-not $g) { throw "Workspace $wsId not found, or this account has no role in it. Sign in with an account that has a role in the workspace (or with a Fabric administrator account)." }
            $wsPath = "groups/$wsId"

            $users = Try-PbiGet "$wsPath/users"
            if (-not $users) { Write-Warning "Workspace users couldn't be read: $script:LastError" }

            $instances = New-Object System.Collections.Generic.List[object]
            $flows = Try-PbiGet "$wsPath/datasets/upstreamDataflows"
            $models = foreach ($d in @((Invoke-PbiGet "$wsPath/datasets").value) | Where-Object { $_ }) {
                $usages = @()
                $src = Try-PbiGet "$wsPath/datasets/$($d.id)/datasources"
                if ($src) {
                    foreach ($x in @($src.value) | Where-Object { $_ }) {
                        $id = [guid]::NewGuid().ToString()
                        $instances.Add([pscustomobject]@{
                            datasourceId = $id; connectionId = Get-Prop $x "datasourceId"; datasourceType = Get-Prop $x "datasourceType"
                            connectionDetails = Get-Prop $x "connectionDetails"; gatewayId = Get-Prop $x "gatewayId"
                        })
                        $usages += [pscustomobject]@{ datasourceInstanceId = $id }
                    }
                }
                $up = @()
                if ($flows) {
                    $up = @(@($flows.value) | Where-Object { $_ -and $_.datasetObjectId -eq $d.id } | ForEach-Object {
                        [pscustomobject]@{ targetDataflowId = $_.dataflowObjectId; groupId = $_.workspaceObjectId }
                    })
                }
                [pscustomobject]@{
                    id = $d.id; name = $d.name; ContentProviderType = Get-Prop $d "ContentProviderType"
                    configuredBy = Get-Prop $d "configuredBy"; createdDate = Get-Prop $d "createdDate"; description = Get-Prop $d "description"
                    endorsementDetails = Get-Prop $d "endorsementDetails"; sensitivityLabel = Get-Prop $d "sensitivityLabel"
                    datasourceUsages = $usages; upstreamDataflows = $up
                }
            }

            $dbList = Try-PbiGet "$wsPath/dashboards"
            $dbs = foreach ($db in @($(if ($dbList) { $dbList.value })) | Where-Object { $_ }) {
                $tiles = Try-PbiGet "$wsPath/dashboards/$($db.id)/tiles"
                [pscustomobject]@{ id = $db.id; displayName = $db.displayName; tiles = @($(if ($tiles) { $tiles.value })) }
            }

            [pscustomobject]@{
                workspaces = @([pscustomobject]@{
                    id = $g.id; name = $g.name
                    isOnDedicatedCapacity = Get-Prop $g "isOnDedicatedCapacity"; capacityId = Get-Prop $g "capacityId"
                    users = @($(if ($users) { $users.value }))
                    datasets = @($models)
                    reports = @((Invoke-PbiGet "$wsPath/reports").value)
                    dashboards = @($dbs)
                })
                datasourceInstances = $instances.ToArray()
            }
        }

        $result = $null
        if (-not $NoAdminScan -and -not $script:adminScanRefused) {
            Write-Host "Scanning workspace $wsId (admin scan) ..." -ForegroundColor Cyan
            try { $result = Get-WorkspaceByAdminScan; $readBy = "admin scan" }
            catch {
                $why = Get-ShortError $_
                if ($why -match '401|403') { $script:adminScanRefused = $true }   # don't retry it for the other workspaces
                Write-Host "  Admin scan not available ($why); reading the workspace with your workspace role instead." -ForegroundColor Yellow
            }
        }
        if (-not $result) {
            Write-Host "Reading workspace $wsId ..." -ForegroundColor Cyan
            $result = Get-WorkspaceByApis; $readBy = "workspace APIs"
        }

        $wsInfo = @($result.workspaces) | Select-Object -First 1
        $wsName = $wsInfo.name
        $ws     = "groups/$wsId"
        Write-Host "Workspace: $wsName (read with the $readBy)" -ForegroundColor Cyan
        $isFabricAdmin = $readBy -eq "admin scan"   # admin-only calls below are skipped otherwise

        [hashtable] $sourceById = @{}
        foreach ($s in @(Get-Prop $result "datasourceInstances") + @(Get-Prop $result "misconfiguredDatasourceInstances")) {
            if ($s -and (Get-Prop $s "datasourceId")) { $sourceById[$s.datasourceId] = $s }
        }

        $datasets   = @(Get-Prop $wsInfo "datasets")   | Where-Object { $_ }
        $reports    = @(Get-Prop $wsInfo "reports")    | Where-Object { $_ -and -not (Get-Prop $_ "appId") }   # skip app copies
        $dashboards = @(Get-Prop $wsInfo "dashboards") | Where-Object { $_ -and -not (Get-Prop $_ "appId") }

        # Workspace app (admin only). App names and publishers come from the
        # tenant's app list (read once per run), audiences from the app's users.
        # Which items are in the app comes from two places, combined:
        #   - the scan's app copies: each report and dashboard in the app is listed
        #     again with an appId (and, for reports, usually originalReportId);
        #   - the app's own lists (apps/{id}/reports and /dashboards), which the
        #     signed-in account can read when it has access to the app.
        # A copy is matched to its original by ID, else by name ("[App] " removed).
        $wsApps = @()
        [hashtable] $inAppById = @{}
        [hashtable] $appContentsUnknown = @{}   # app name -> why its contents couldn't be read
        [hashtable] $appUnmatched = @{}         # app name -> app items not matched to a workspace item
        if ($isFabricAdmin) {
            if ($null -eq $script:appsByWorkspace) {
                $script:appsByWorkspace = @{}
                $skip = 0
                do {
                    $page = Try-PbiGet "admin/apps?`$top=5000$(if ($skip) { "&`$skip=$skip" })"
                    if (-not $page) { Write-Warning "Workspace apps couldn't be read: $script:LastError"; break }
                    foreach ($a in @($page.value) | Where-Object { $_ }) {
                        $key = "$(Get-Prop $a 'workspaceId')".ToLowerInvariant()
                        if (-not $script:appsByWorkspace.ContainsKey($key)) { $script:appsByWorkspace[$key] = @() }
                        $script:appsByWorkspace[$key] += $a
                    }
                    $skip += 5000
                } while (@($page.value).Count -eq 5000)
            }
            $wsApps = @($script:appsByWorkspace["$wsId".ToLowerInvariant()]) | Where-Object { $_ }
            $appName = @{}
            foreach ($a in $wsApps) { $appName[$a.id] = $a.name }
            $byName = @{ report = @{}; dashboard = @{} }
            foreach ($r in $reports)    { $byName.report["$($r.name)".Trim().ToLowerInvariant()] = $r.id }
            foreach ($db in $dashboards) { $byName.dashboard["$($db.displayName)".Trim().ToLowerInvariant()] = $db.id }
            $originalIds = @{}
            foreach ($x in @($reports) + @($dashboards)) { if ($x) { $originalIds["$($x.id)"] = $true } }

            # Match one app item to the workspace item it was published from
            function Find-Original($Copy, [string] $Kind) {
                $orig = "$(Get-Prop $Copy 'originalReportId')"
                if ($orig -and $originalIds.ContainsKey($orig)) { return $orig }
                if ($originalIds.ContainsKey("$($Copy.id)")) { return "$($Copy.id)" }   # some lists give the original's own ID
                $name = $(if (Get-Prop $Copy "name") { $Copy.name } else { Get-Prop $Copy "displayName" })
                $plain = ("$name" -replace '^\s*\[App\]\s*', '').Trim().ToLowerInvariant()
                $byName[$Kind][$plain]
            }

            foreach ($a in $wsApps) {
                $copies = New-Object System.Collections.Generic.List[object]
                foreach ($c in @(Get-Prop $wsInfo "reports") | Where-Object { $_ -and (Get-Prop $_ "appId") -eq $a.id }) { $copies.Add([pscustomobject]@{ Item = $c; Kind = "report" }) }
                foreach ($c in @(Get-Prop $wsInfo "dashboards") | Where-Object { $_ -and (Get-Prop $_ "appId") -eq $a.id }) { $copies.Add([pscustomobject]@{ Item = $c; Kind = "dashboard" }) }
                $fromApp = 0; $appErrors = @()
                foreach ($kind in "report", "dashboard") {
                    $list = Try-PbiGet "apps/$($a.id)/$($kind)s"
                    if ($list) { foreach ($c in @($list.value) | Where-Object { $_ }) { $copies.Add([pscustomobject]@{ Item = $c; Kind = $kind }); $fromApp++ } }
                    else { $appErrors += "$($kind)s: $script:LastError" }
                }
                $matched = 0; $unmatched = @()
                foreach ($c in $copies) {
                    $orig = Find-Original $c.Item $c.Kind
                    if ($orig) { if (-not $inAppById.ContainsKey($orig)) { $matched++ }; $inAppById[$orig] = $a.name }
                    else { $unmatched += $(if (Get-Prop $c.Item "name") { $c.Item.name } else { Get-Prop $c.Item "displayName" }) }
                }
                $unmatched = @($unmatched | Select-Object -Unique)
                if ($unmatched.Count) { $appUnmatched[$a.name] = $unmatched }
                if (-not $copies.Count) {
                    $appContentsUnknown[$a.name] = "the scan listed none of the app's items and the app's own lists couldn't be read ($($appErrors -join '; '))"
                }
                Write-Host ("  App '{0}': {1} items matched to the workspace{2}{3}" -f $a.name, $matched,
                    $(if ($appErrors.Count -and $fromApp -eq 0) { "; app's own lists not readable ($($appErrors -join '; '))" } else { "" }),
                    $(if ($unmatched.Count) { "; not matched: $($unmatched -join ', ')" } else { "" }))
            }
        }
        [hashtable] $datasetById = @{}
        foreach ($d in $datasets) { $datasetById[$d.id] = $d }

        # ---------------------------------------------------------------------------
        # 2. Workspace permissions
        # ---------------------------------------------------------------------------

        $permissionRows = @(foreach ($u in @(Get-Prop $wsInfo "users") | Where-Object { $_ }) {
            [pscustomobject]@{
                WorkspaceName = $wsName
                Principal     = $(if (Get-Prop $u "displayName") { $u.displayName } else { Get-Prop $u "identifier" })
                Email         = Get-Prop $u "emailAddress"
                PrincipalType = Get-Prop $u "principalType"
                Role          = Get-Prop $u "groupUserAccessRight"
                Keys          = @(Get-PrincipalKeys $u)
            }
        })
        $permissionSummary = Join-NonEmpty " | " (("Admin", "Member", "Contributor", "Viewer") | ForEach-Object {
            $role = $_
            $names = @($permissionRows | Where-Object Role -eq $role | Sort-Object Principal | ForEach-Object {
                if ($_.Email -and $_.Email -ne $_.Principal) { "$($_.Principal) ($($_.Email))" } else { $_.Principal }
            })
            if ($names.Count) { "${role}: $($names -join ', ')" }
        })
        if (-not $permissionSummary) { $permissionSummary = "(no users returned)" }

        # Who has which workspace role, by any of their IDs, for the item-access check
        [hashtable] $roleByKey = @{}
        foreach ($p in $permissionRows) { foreach ($k in $p.Keys) { $roleByKey[$k] = $p.Role } }

        # Item access for one item. The admin scan includes each item's users;
        # without it only semantic models can be read (needs Admin or Member).
        function Get-ItemAccess($Item, [string] $Kind) {
            $right = @{ dataset = "datasetUserAccessRight"; report = "reportUserAccessRight"; dashboard = "dashboardUserAccessRight" }[$Kind]
            if ($isFabricAdmin) {
                if ($null -eq $Item.PSObject.Properties["users"]) { return "Not returned by the admin scan" }
                return Format-ItemAccess (Get-Prop $Item "users") $right
            }
            if ($Kind -ne "dataset") { return $noAdminAccess }
            $u = Try-PbiGet "$ws/datasets/$($Item.id)/users"
            if (-not $u) { return "Could not read: $script:LastError (needs Admin or Member on the workspace)" }
            Format-ItemAccess $u.value $right
        }

        # ---------------------------------------------------------------------------
        # 3. Folders (Fabric API)
        # ---------------------------------------------------------------------------

        [hashtable] $folderOfItem = @{}
        $folderRows   = @()
        $folderSource = "Fabric API"

        # Folder names: needs a role in the workspace (Viewer or higher).
        [hashtable] $byId = @{}
        try {
            foreach ($f in Get-FabricPaged "$FabricApiUrl/workspaces/$wsId/folders") { $byId[$f.id] = $f }
        } catch {
            $folderSource = "folder names unavailable ($(Get-ShortError $_)); add yourself to the workspace to see them"
            Write-Warning "Folder names couldn't be read: $(Get-ShortError $_). They need a role in the workspace; folder IDs are shown instead."
        }
        $pathOf = {
            param($id)
            if (-not $byId.ContainsKey($id)) { return "(folder $id)" }
            $parts = @(); $seen = 0
            while ($id -and $byId.ContainsKey($id) -and $seen -lt 50) {
                $parts = @($byId[$id].displayName) + $parts
                $id = Get-Prop $byId[$id] "parentFolderId"
                $seen++
            }
            $parts -join "/"
        }
        $folderRows = foreach ($f in $byId.Values) {
            [pscustomobject]@{ WorkspaceName = $wsName; FolderPath = & $pathOf $f.id; FolderId = $f.id }
        }

        # Which folder each item is in: workspace call first, admin call if you have no role in the workspace.
        $items = $null
        try { $items = @(Get-FabricPaged "$FabricApiUrl/workspaces/$wsId/items") }
        catch {
            try {
                if (-not $isFabricAdmin) { throw }
                $items = @(Get-FabricPaged "$FabricApiUrl/admin/items?workspaceId=$wsId" "itemEntities")
            } catch {
                $folderSource = "unavailable ($(Get-ShortError $_))"
                Write-Warning "Item folders couldn't be read from the Fabric API: $(Get-ShortError $_)"
            }
        }
        foreach ($i in @($items) | Where-Object { $_ }) {
            $fid = Get-Prop $i "folderId"
            $folderOfItem[$i.id] = $(if ($fid) { & $pathOf $fid } else { "(workspace root)" })
        }

        function Get-ItemFolder([string] $Id) {
            if ($folderOfItem.ContainsKey($Id)) { return $folderOfItem[$Id] }
            if ($folderOfItem.Count) { return "(not found)" }
            "Unknown"
        }

        # ---------------------------------------------------------------------------
        # 4. Refresh schedules
        # ---------------------------------------------------------------------------

        function Format-Schedule($S, [string] $Label) {
            if (-not $S) { return $null }
            $enabled = Get-Prop $S "enabled"
            if ($enabled -eq $false) { return "$Label off" }
            $parts = @("$Label on")
            if (Get-Prop $S "frequency")             { $parts += "every $($S.frequency) min" }
            if (@(Get-Prop $S "days")  | Where-Object { $_ }) { $parts += (@($S.days)  -join ",") }
            if (@(Get-Prop $S "times") | Where-Object { $_ }) { $parts += (@($S.times) -join ",") }
            if (Get-Prop $S "localTimeZoneId")       { $parts += $S.localTimeZoneId }
            $parts -join "; "
        }

        # Fallback for when you aren't a member of the workspace: the capacity's refreshables (admin, Premium/Fabric only).
        [hashtable] $refreshables = @{}
        $capacityId = Get-Prop $wsInfo "capacityId"
        if ($isFabricAdmin -and (Get-Prop $wsInfo "isOnDedicatedCapacity") -and $capacityId) {
            $r = Try-PbiGet "admin/capacities/$capacityId/refreshables?`$top=5000"
            if ($r) { foreach ($x in @($r.value)) { $refreshables[$x.id] = $x } }
        }

        [hashtable] $scheduleCache = @{}
        function Get-ModelSchedule([string] $ModelWsId, [string] $ModelId) {
            $key = "$ModelWsId|$ModelId"
            if ($scheduleCache.ContainsKey($key)) { return $scheduleCache[$key] }
            $s1 = Format-Schedule (Try-PbiGet "groups/$ModelWsId/datasets/$ModelId/refreshSchedule") "Scheduled refresh"
            $err = $script:LastError
            $s2 = Format-Schedule (Try-PbiGet "groups/$ModelWsId/datasets/$ModelId/directQueryRefreshSchedule") "DirectQuery cache refresh"
            $text = Join-NonEmpty " | " @($s1, $s2)
            if (-not $text -and $refreshables.ContainsKey($ModelId)) {
                $text = Format-Schedule (Get-Prop $refreshables[$ModelId] "refreshSchedule") "Scheduled refresh"
                if ($text) { $text = "$text (from capacity refreshables)" }
            }
            if (-not $text) {
                $d = $datasetById[$ModelId]
                $mode = Get-Prop $d "ContentProviderType"
                if ($mode -match 'DirectQuery')   { $text = "None (DirectQuery: queries the source live)" }
                elseif ($err -match '403|401')    { $text = "Could not read: add yourself to the workspace to see schedules" }
                elseif ($err)                     { $text = "None or not readable ($err)" }
                else                              { $text = "None" }
            }
            $scheduleCache[$key] = $text
            $text
        }

        # ---------------------------------------------------------------------------
        # 5. Semantic model tables and queries
        #    From the scan when the tenant settings allow it; otherwise from the
        #    model's definition (Fabric "Get Semantic Model Definition").
        # ---------------------------------------------------------------------------

        [hashtable] $mConnectorsByModel = @{}

        # Takes a model's tables and shared expressions in one shape, whichever way
        # they were read, adds them to the tables CSV and returns the counts.
        #   Table:      Name, Kind, Hidden, Columns, Measures, Query
        #   Expression: Name, Query
        function Add-ModelRows([string] $ModelId, [string] $ModelName, $Tables, $Expressions, [string] $From) {
            $counts = @{ pq = 0; calc = 0; dl = 0; cg = 0; other = 0; auto = 0; params = 0; shared = 0 }
            $connectors = @()
            foreach ($t in @($Tables) | Where-Object { $_ }) {
                if ($t.Name -match '^(LocalDateTable_|DateTableTemplate_)') { $counts.auto++; continue }
                switch -Wildcard ($t.Kind) {
                    "Power Query*"         { $counts.pq++ }
                    "DAX calculated table" { $counts.calc++ }
                    "Direct Lake"          { $counts.dl++ }
                    "Calculation group"    { $counts.cg++ }
                    default                { $counts.other++ }
                }
                $src = Get-MSources $t.Query
                if ($src) { $connectors += $src }
                $tableRows.Add([pscustomobject]@{
                    ItemId = $ModelId; WorkspaceId = $wsId; WorkspaceName = $wsName; ItemName = $ModelName; ItemType = "Semantic model"
                    Name = $t.Name; Kind = $t.Kind; Hidden = $t.Hidden
                    Columns = $t.Columns; Measures = $t.Measures
                    Sources = $src; Query = $t.Query
                })
            }
            foreach ($e in @($Expressions) | Where-Object { $_ }) {
                $isParam = "$($e.Query)" -match 'IsParameterQuery\s*=\s*true'
                if ($isParam) { $counts.params++ } else { $counts.shared++ }
                $src = Get-MSources $e.Query
                if ($src) { $connectors += $src }
                $tableRows.Add([pscustomobject]@{
                    ItemId = $ModelId; WorkspaceId = $wsId; WorkspaceName = $wsName; ItemName = $ModelName; ItemType = "Semantic model"
                    Name = $e.Name; Kind = $(if ($isParam) { "Parameter" } else { "Shared query (not loaded as a table)" })
                    Hidden = $null; Columns = $null; Measures = $null
                    Sources = $src; Query = $e.Query
                })
            }
            $mConnectorsByModel[$ModelId] = Join-NonEmpty ", " (($connectors -join ", ") -split ",\s*")

            $parts = @("$($counts.pq) Power Query tables")
            if ($counts.dl)    { $parts += "$($counts.dl) Direct Lake tables" }
            $parts += "$($counts.calc) DAX calculated tables"
            if ($counts.cg)    { $parts += "$($counts.cg) calculation groups" }
            if ($counts.other) { $parts += "$($counts.other) other tables" }
            $parts += "$($counts.shared) shared queries", "$($counts.params) parameters", "$($counts.auto) auto date tables (not counted)"
            [pscustomobject]@{
                Tables  = $counts.pq + $counts.calc + $counts.dl + $counts.cg + $counts.other
                Queries = $counts.pq + $counts.shared
                Detail  = "$($parts -join ', ') (from $From)"
            }
        }

        # --- Read from the scan -----------------------------------------------------

        function Get-ScanModelParts($D) {
            $tables = foreach ($t in @(Get-Prop $D "tables") | Where-Object { $_ }) {
                $expr = (@(Get-Prop $t "source") | Where-Object { $_ } | ForEach-Object { Get-Prop $_ "expression" }) -join "`n"
                $kind = $(if (-not $expr) { "Unknown (no source returned)" }
                          elseif ($expr -match '(?s)^\s*let\b.*\bin\b' -or $expr -match $mSourcePattern -or $expr -match '^\s*#?"?[\w ]+"?\s*$') { "Power Query" }
                          else { "DAX calculated table" })
                [pscustomobject]@{
                    Name = $t.name; Kind = $kind; Hidden = [bool](Get-Prop $t "isHidden")
                    Columns  = @(Get-Prop $t "columns" | Where-Object { $_ -and $_.name -notlike 'RowNumber-*' }).Count
                    Measures = @(Get-Prop $t "measures" | Where-Object { $_ }).Count
                    Query = $expr
                }
            }
            $exprs = foreach ($e in @(Get-Prop $D "expressions") | Where-Object { $_ }) {
                [pscustomobject]@{ Name = $e.name; Query = Get-Prop $e "expression" }
            }
            [pscustomobject]@{
                Tables = @($tables); Expressions = @($exprs)
                Roles = @(ConvertFrom-RoleObjects (Get-Prop $D "roles")); RolesKnown = ($null -ne $D.PSObject.Properties["roles"])
            }
        }

        # --- Read from the model definition (TMDL) ----------------------------------

        function Get-FirstHeader($Response, [string] $Name) {
            $v = $Response.Headers[$Name]
            if ($v -is [array]) { $v[0] } else { $v }
        }

        # POST .../getDefinition only returns a copy of the definition; it changes nothing.
        function Get-ModelDefinitionParts([string] $ModelId) {
            $h = @{ Authorization = (Get-PowerBIAccessToken -AsString) }
            $url = "$FabricApiUrl/workspaces/$wsId/semanticModels/$ModelId/getDefinition?format=TMDL"
            $resp = Invoke-WebRequest -Uri $url -Method Post -Headers $h -UseBasicParsing -ErrorAction Stop
            if ($resp.StatusCode -eq 202) {
                $location = Get-FirstHeader $resp "Location"
                $wait     = 3
                $retry    = Get-FirstHeader $resp "Retry-After"
                if ($retry -as [int]) { $wait = [int]$retry }
                $deadline = (Get-Date).AddMinutes(5)
                do {
                    Start-Sleep -Seconds ([Math]::Max(2, [Math]::Min($wait, 20)))
                    $op = Invoke-RestMethod -Uri $location -Method Get -Headers $h -ErrorAction Stop
                    if ((Get-Date) -gt $deadline) { throw "the definition wasn't ready after 5 minutes" }
                } while ($op.status -in "NotStarted", "Running", "Undefined")
                if ($op.status -ne "Succeeded") { throw "the definition request $($op.status): $(Get-Prop (Get-Prop $op 'error') 'message')" }
                $def = Invoke-RestMethod -Uri "$location/result" -Method Get -Headers $h -ErrorAction Stop
            } else {
                $def = $resp.Content | ConvertFrom-Json
            }
            [hashtable] $parts = @{}
            foreach ($p in @($def.definition.parts)) {
                $parts[$p.path] = [Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($p.payload))
            }
            $parts
        }

        function Get-Indent([string] $Line) {
            $ws = [regex]::Match($Line, '^[\t ]*').Value
            ($ws.ToCharArray() | Where-Object { $_ -eq "`t" }).Count + [Math]::Floor(($ws.ToCharArray() | Where-Object { $_ -eq ' ' }).Count / 4)
        }

        function ConvertFrom-TmdlName([string] $Name) {
            $n = $Name.Trim()
            if ($n.Length -ge 2 -and $n.StartsWith("'") -and $n.EndsWith("'")) { $n = $n.Substring(1, $n.Length - 2) -replace "''", "'" }
            $n
        }

        # The value of a TMDL property: what follows "=" on its line, or the lines
        # below it that are indented deeper than $PropIndent. Returns text and the next line index.
        function Get-TmdlValue([string[]] $Lines, [int] $Index, [int] $PropIndent, [string] $Rest) {
            $body = @(); $k = $Index + 1
            while ($k -lt $Lines.Count -and ($Lines[$k].Trim() -eq "" -or (Get-Indent $Lines[$k]) -gt $PropIndent)) {
                $body += $Lines[$k]; $k++
            }
            while ($body.Count -and $body[-1].Trim() -eq "") { $body = @($body | Select-Object -SkipLast 1) }
            if ($body.Count) {
                $min = ($body | Where-Object { $_.Trim() } | ForEach-Object { [regex]::Match($_, '^[\t ]*').Value.Length } | Measure-Object -Minimum).Minimum
                $body = $body | ForEach-Object { if ($_.Length -ge $min) { $_.Substring($min) } else { $_.TrimStart() } }
            }
            $text = (@($(if ($Rest.Trim()) { $Rest.Trim() })) + @($body)) -join "`n"
            if ($text.StartsWith('```')) { $text = ($text -replace '^```\s*', '' -replace '\s*```$', '') }
            [pscustomobject]@{ Text = $text; Next = $k }
        }

        $tmdlName = "('(?:[^']|'')+'|[^\s=]+)"

        function ConvertFrom-TmdlTable([string] $Text) {
            $lines = $Text -split "`r?`n"
            $name = $null; $hidden = $false; $columns = 0; $measures = 0
            $types = @(); $query = $null; $policyQuery = $null; $section = $null; $isCalcGroup = $false
            $i = 0
            while ($i -lt $lines.Count) {
                $line = $lines[$i]; $indent = Get-Indent $line; $t = $line.Trim()
                if ($indent -eq 0 -and $t -match "^table\s+$tmdlName") { $name = ConvertFrom-TmdlName $Matches[1] }
                elseif ($indent -eq 1) {
                    $section = ($t -split '\s+')[0]
                    if ($t -match '^column\s')               { $columns++ }
                    elseif ($t -match '^measure\s')          { $measures++ }
                    elseif ($t -match '^isHidden\s*$')       { $hidden = $true }
                    elseif ($t -match '^calculationGroup\b') { $isCalcGroup = $true }
                    elseif ($t -match "^partition\s+$tmdlName\s*=\s*(\w+)") { $types += $Matches[2] }
                }
                elseif ($indent -eq 2 -and $t -match '^(?:source|sourceExpression)\s*=\s*(.*)$' -and $section -in "partition", "refreshPolicy") {
                    $v = Get-TmdlValue $lines $i $indent $Matches[1]
                    if ($section -eq "partition") { if (-not $query) { $query = $v.Text } } else { $policyQuery = $v.Text }
                    $i = $v.Next; continue
                }
                $i++
            }
            $kind = $(if ($isCalcGroup -or $types -contains "calculationGroup") { "Calculation group" }
                      elseif ($types -contains "policyRange")                    { "Power Query (incremental refresh)" }
                      elseif ($types -contains "m")                              { "Power Query" }
                      elseif ($types -contains "calculated")                     { "DAX calculated table" }
                      elseif ($types -contains "entity")                         { "Direct Lake" }
                      elseif ($types.Count)                                      { "Other ($($types -join ', '))" }
                      else                                                       { "Unknown (no partition)" })
            if (-not $query -and $policyQuery) { $query = $policyQuery }
            [pscustomobject]@{ Name = $name; Kind = $kind; Hidden = $hidden; Columns = $columns; Measures = $measures; Query = $query }
        }

        function ConvertFrom-TmdlExpressions([string] $Text) {
            $lines = $Text -split "`r?`n"
            $i = 0
            while ($i -lt $lines.Count) {
                $line = $lines[$i]
                if ((Get-Indent $line) -eq 0 -and $line.Trim() -match "^expression\s+$tmdlName\s*=\s*(.*)$") {
                    $name = ConvertFrom-TmdlName $Matches[1]
                    $v = Get-TmdlValue $lines $i 1 $Matches[2]
                    [pscustomobject]@{ Name = $name; Query = $v.Text }
                    $i = $v.Next; continue
                }
                $i++
            }
        }

        # Older models can come back as model.bim (TMSL) instead of TMDL.
        function ConvertFrom-ModelBim([string] $Text) {
            $model = ($Text | ConvertFrom-Json).model
            $tables = foreach ($t in @(Get-Prop $model "tables") | Where-Object { $_ }) {
                $types = @(Get-Prop $t "partitions" | ForEach-Object { Get-Prop (Get-Prop $_ "source") "type" })
                $expr  = @(Get-Prop $t "partitions" | ForEach-Object { @(Get-Prop (Get-Prop $_ "source") "expression") -join "`n" }) | Select-Object -First 1
                $kind = $(if (Get-Prop $t "calculationGroup") { "Calculation group" }
                          elseif ($types -contains "policyRange") { "Power Query (incremental refresh)" }
                          elseif ($types -contains "m")          { "Power Query" }
                          elseif ($types -contains "calculated") { "DAX calculated table" }
                          elseif ($types -contains "entity")     { "Direct Lake" }
                          else                                   { "Other ($($types -join ', '))" })
                [pscustomobject]@{
                    Name = $t.name; Kind = $kind; Hidden = [bool](Get-Prop $t "isHidden")
                    Columns = @(Get-Prop $t "columns").Count; Measures = @(Get-Prop $t "measures").Count; Query = $expr
                }
            }
            $exprs = foreach ($e in @(Get-Prop $model "expressions") | Where-Object { $_ }) {
                [pscustomobject]@{ Name = $e.name; Query = @($e.expression) -join "`n" }
            }
            [pscustomobject]@{ Tables = @($tables); Expressions = @($exprs); Roles = @(ConvertFrom-RoleObjects (Get-Prop $model "roles")); RolesKnown = $true }
        }

        # --- Row-level security roles ----------------------------------------------
        # One shape whichever way they were read:
        #   Name, Permission, Filters ("Table: DAX"), Hidden (object-level security), Members

        function ConvertFrom-TmdlRole([string] $Text) {
            $lines = $Text -split "`r?`n"
            $role = [pscustomobject]@{ Name = $null; Permission = $null; Filters = @(); Hidden = @(); Members = @() }
            $section = $null; $table = $null; $column = $null
            $i = 0
            while ($i -lt $lines.Count) {
                $line = $lines[$i]; $indent = Get-Indent $line; $t = $line.Trim()
                if ($indent -eq 0 -and $t -match "^role\s+$tmdlName") { $role.Name = ConvertFrom-TmdlName $Matches[1] }
                elseif ($indent -eq 1) {
                    $section = ($t -split '\s+')[0]; $column = $null
                    if ($t -match '^modelPermission:\s*(\w+)') { $role.Permission = $Matches[1] }
                    elseif ($t -match "^tablePermission\s+$tmdlName\s*=\s*(.*)$") {
                        $table = ConvertFrom-TmdlName $Matches[1]
                        $v = Get-TmdlValue $lines $i 1 $Matches[2]
                        $role.Filters += "${table}: $(($v.Text -split '\r?\n' | ForEach-Object { $_.Trim() }) -join ' ')"
                        $i = $v.Next; continue
                    }
                    elseif ($t -match "^tablePermission\s+$tmdlName\s*$") { $table = ConvertFrom-TmdlName $Matches[1] }
                    elseif ($t -match "^member\s+$tmdlName") { $role.Members += ConvertFrom-TmdlName $Matches[1] }
                }
                elseif ($indent -eq 2 -and $section -eq "tablePermission") {
                    if ($t -match '^metadataPermission:\s*none') { $role.Hidden += $table }
                    elseif ($t -match "^columnPermission\s+$tmdlName") { $column = ConvertFrom-TmdlName $Matches[1] }
                }
                elseif ($indent -eq 3 -and $section -eq "tablePermission" -and $column -and $t -match '^metadataPermission:\s*none') { $role.Hidden += "$table[$column]" }
                elseif ($indent -eq 2 -and $section -eq "member" -and $t -match '^memberType:\s*group') {
                    $role.Members[-1] = "$($role.Members[-1]) (group)"
                }
                $i++
            }
            $role
        }

        function ConvertFrom-RoleObjects($Roles) {
            foreach ($r in @($Roles) | Where-Object { $_ }) {
                $filters = @(); $hidden = @()
                foreach ($tp in @(Get-Prop $r "tablePermissions") | Where-Object { $_ }) {
                    $f = @(Get-Prop $tp "filterExpression") -join " "
                    if ($f) { $filters += "$($tp.name): $(($f -split '\r?\n' | ForEach-Object { $_.Trim() }) -join ' ')" }
                    if ("$(Get-Prop $tp 'metadataPermission')" -eq "none") { $hidden += $tp.name }
                    foreach ($cp in @(Get-Prop $tp "columnPermissions") | Where-Object { $_ }) {
                        if ("$(Get-Prop $cp 'metadataPermission')" -eq "none") { $hidden += "$($tp.name)[$($cp.name)]" }
                    }
                }
                $members = @(Get-Prop $r "members") | Where-Object { $_ } | ForEach-Object {
                    $n = $(if (Get-Prop $_ "memberName") { $_.memberName } else { Get-Prop $_ "memberId" })
                    $(if ("$(Get-Prop $_ 'memberType')" -match 'group') { "$n (group)" } else { "$n" })
                }
                [pscustomobject]@{ Name = $r.name; Permission = Get-Prop $r "modelPermission"; Filters = $filters; Hidden = $hidden; Members = @($members) }
            }
        }

        function Format-Roles($Roles) {
            $list = @($Roles) | Where-Object { $_ }
            if (-not $list.Count) { return "No row-level security roles" }
            ($list | Sort-Object Name | ForEach-Object {
                $parts = @()
                $parts += $(if ($_.Filters.Count) { "filters: $($_.Filters -join '; ')" } else { "no filters" })
                if ($_.Hidden.Count) { $parts += "hides (object-level security): $($_.Hidden -join ', ')" }
                $parts += $(if ($_.Members.Count) { "members: $($_.Members -join ', ')" } else { "members: none assigned" })
                "Role '$($_.Name)'$(if ($_.Permission) { " ($($_.Permission))" }): $($parts -join '; ')"
            }) -join " | "
        }

        function Get-DefinitionModelParts([string] $ModelId) {
            $parts = Get-ModelDefinitionParts $ModelId
            $tableKeys = @($parts.Keys | Where-Object { $_ -like "definition/tables/*.tmdl" })
            if ($tableKeys.Count) {
                $tables = foreach ($k in $tableKeys) { ConvertFrom-TmdlTable $parts[$k] }
                $exprs  = @()
                if ($parts.ContainsKey("definition/expressions.tmdl")) { $exprs = @(ConvertFrom-TmdlExpressions $parts["definition/expressions.tmdl"]) }
                $roles = @($parts.Keys | Where-Object { $_ -like "definition/roles/*.tmdl" } | ForEach-Object { ConvertFrom-TmdlRole $parts[$_] })
                return [pscustomobject]@{ Tables = @($tables); Expressions = $exprs; Roles = $roles; RolesKnown = $true }
            }
            if ($parts.ContainsKey("model.bim")) { return ConvertFrom-ModelBim $parts["model.bim"] }
            throw "the definition had no tables (parts: $(@($parts.Keys) -join ', '))"
        }

        # --- Pick the source for each model ------------------------------------------

        function Get-ModelQueryInfo($D) {
            if (@(Get-Prop $D "tables" | Where-Object { $_ }).Count) {
                $m = Get-ScanModelParts $D
                $info = Add-ModelRows $D.id $D.name $m.Tables $m.Expressions "admin scan"
                if (Get-Prop $D "schemaMayNotBeUpToDate") { $info.Detail += " Schema may not be up to date." }
                $security = $(if ($m.RolesKnown) { Format-Roles $m.Roles } else { "Unknown (the admin scan didn't return roles)" })
                $info | Add-Member -NotePropertyName Security -NotePropertyValue $security
                return $info
            }
            if ($SkipModelDefinition) {
                return [pscustomobject]@{ Tables = $null; Queries = $null; Detail = "Unknown (the scan returned no tables and -SkipModelDefinition was used)"; Security = "Unknown (-SkipModelDefinition was used)" }
            }
            try {
                $m = Get-DefinitionModelParts $D.id
                $info = Add-ModelRows $D.id $D.name $m.Tables $m.Expressions "model definition"
                $info | Add-Member -NotePropertyName Security -NotePropertyValue (Format-Roles $m.Roles)
                return $info
            } catch {
                $why = Get-ShortError $_
                if ($why -match '403') {
                    $label = Get-Prop (Get-Prop $D "sensitivityLabel") "labelId"
                    $why = "$why; " + $(
                        if ($why -match 'Scope') { "the sign-in doesn't include the SemanticModel.ReadWrite.All permission this call needs" }
                        elseif ($label) { "the model has a sensitivity label ($label): if it's encrypted, reading the definition needs usage rights to remove the label in Microsoft Purview" }
                        else { "reading a model's definition needs Contributor or higher, and is refused for models with an encrypted sensitivity label" })
                }
                return [pscustomobject]@{ Tables = $null; Queries = $null; Detail = "Unknown (scan returned no tables; definition: $why)"; Security = "Unknown (definition: $why)" }
            }
        }

        function Get-ModelSources($D) {
            $list = foreach ($u in @(Get-Prop $D "datasourceUsages") | Where-Object { $_ }) {
                if ($sourceById.ContainsKey($u.datasourceInstanceId)) { Format-DataSource $sourceById[$u.datasourceInstanceId] "$wsName / $($D.name)" }
            }
            $flows = @(Get-Prop $D "upstreamDataflows" | Where-Object { $_ }).Count
            if ($flows) { $list = @($list) + "Dataflows: $flows" }
            $models = @(Get-Prop $D "upstreamDatasets" | Where-Object { $_ }).Count
            if ($models) { $list = @($list) + "Other semantic models: $models" }
            $text = Join-NonEmpty "; " $list
            if (-not $text) {
                # Fall back to the connector functions found in the model's M code
                $m = $(if ($mConnectorsByModel.ContainsKey($D.id)) { $mConnectorsByModel[$D.id] })
                $text = $(if ($m) { "From M code: $m" } else { "(none returned)" })
            }
            $text
        }

        function Get-ModelLabel([string] $ModelId, [string] $ModelWsId) {
            if ($datasetById.ContainsKey($ModelId)) { return "'$($datasetById[$ModelId].name)'" }
            $(if ($ModelWsId) { "$ModelId (in workspace $ModelWsId)" } else { "$ModelId (in another workspace)" })
        }

        # ---------------------------------------------------------------------------
        # 6. Paginated report datasets (.rdl)
        # ---------------------------------------------------------------------------

        function Get-RdlInfo([string] $ReportId, [string] $ReportName) {
            $rdl = $null; $from = $null; $why = $null
            if (-not $SkipRdlDownload) {
                $tmp = Join-Path ([IO.Path]::GetTempPath()) "pbi-inventory-$ReportId.rdl"
                try {
                    Invoke-PowerBIRestMethod -Url "$ws/reports/$ReportId/Export" -Method Get -OutFile $tmp -TimeoutSec 900 -ErrorAction Stop
                    $rdl = [xml](Get-Content -LiteralPath $tmp -Raw); $from = "downloaded .rdl"
                } catch { $why = Get-ShortError $_ }
                finally { if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Force } }
            }
            if (-not $rdl) {
                $reason = $(if ($SkipRdlDownload) { "download skipped" } elseif ($why) { "download failed: $why; downloads need Contributor on the workspace" } else { "no .rdl" })
                return [pscustomobject]@{ Count = $null; Detail = "Unknown ($reason)."; Models = @() }
            }

            $ns = New-Object Xml.XmlNamespaceManager $rdl.NameTable
            $ns.AddNamespace("r", $rdl.DocumentElement.NamespaceURI)
            $sets = @($rdl.SelectNodes("//r:DataSets/r:DataSet", $ns))
            foreach ($s in $sets) {
                $cmd = $s.SelectSingleNode("r:Query/r:CommandText", $ns)
                $dsn = $s.SelectSingleNode("r:Query/r:DataSourceName", $ns)
                $tableRows.Add([pscustomobject]@{
                    ItemId = $ReportId; WorkspaceId = $wsId; WorkspaceName = $wsName; ItemName = $ReportName; ItemType = "Paginated report"
                    Name = $s.GetAttribute("Name"); Kind = "Dataset (data source: $(if ($dsn) { $dsn.InnerText }))"
                    Hidden = $null; Columns = @($s.SelectNodes("r:Fields/r:Field", $ns)).Count; Measures = $null
                    Sources = $null; Query = $(if ($cmd) { $cmd.InnerText })
                })
            }
            $models = @($rdl.SelectNodes("//r:DataSource//r:ConnectString", $ns) | ForEach-Object {
                if ($_.InnerText -match 'sobe_wowvirtualserver-([0-9a-fA-F-]{36})') { $Matches[1] }
            })
            [pscustomobject]@{
                Count  = $sets.Count
                Detail = "$($sets.Count) datasets: $((@($sets | ForEach-Object { $_.GetAttribute('Name') })) -join ', ') ($from)"
                Models = $models
            }
        }

        # ---------------------------------------------------------------------------
        # 7. Build the rows
        # ---------------------------------------------------------------------------

        [hashtable] $itemObjById = @{}
        foreach ($x in @($datasets) + @($reports) + @($dashboards) + @($wsApps)) { if ($x -and (Get-Prop $x "id")) { $itemObjById["$($x.id)"] = $x } }
        function Add-Row($Id, $Name, $Type, $Folder, $Schedule, $Sources, $TableCount, $QueryCount, $QueryDetail, $Access, $Security) {
            $meta = Get-ItemMeta $(if ($Id -and $itemObjById.ContainsKey("$Id")) { $itemObjById["$Id"] })
            $inApp = $null
            if ($Type -in "Interactive report", "Paginated report", "Dashboard") {
                $inApp = $(if (-not $isFabricAdmin) { $noAdminAccess }
                           elseif ($inAppById.ContainsKey($Id)) { $inAppById[$Id] }
                           elseif ($appContentsUnknown.Count) { "Unknown: for app '$(@($appContentsUnknown.Keys)[0])', $(@($appContentsUnknown.Values)[0])" }
                           else { "No" })
            }
            $rows.Add([pscustomobject]@{
                WorkspaceName        = $wsName
                WorkspaceId          = $wsId
                ItemName             = $Name
                ItemId               = $Id
                ItemType             = $Type
                Folder               = $Folder
                InApp                = $inApp
                Owner                = $meta.Owner
                Created              = $meta.Created
                ModifiedBy           = $meta.ModifiedBy
                Modified             = $meta.Modified
                Endorsement          = $meta.Endorsement
                SensitivityLabel     = $meta.SensitivityLabel
                Description          = $meta.Description
                WorkspacePermissions = $permissionSummary
                ItemAccess           = $Access
                RefreshSchedule      = $Schedule
                Security             = $Security
                DataSources          = $Sources
                TableCount           = $TableCount
                QueryCount           = $QueryCount
                QueryDetail          = $QueryDetail
            })
        }

        foreach ($d in $datasets) {
            if ($d.name -in "Usage Metrics Report", "Report Usage Metrics Model") { continue }
            Write-Host "  Semantic model: $($d.name)"
            $q = Get-ModelQueryInfo $d
            Add-Row $d.id $d.name "Semantic model" (Get-ItemFolder $d.id) `
                (Get-ModelSchedule "$wsId" $d.id) (Get-ModelSources $d) $q.Tables $q.Queries $q.Detail (Get-ItemAccess $d "dataset") $q.Security
        }

        foreach ($r in $reports) {
            if ((Get-Prop $r "reportType") -eq "PaginatedReport") {
                Write-Host "  Paginated report: $($r.name)"
                $src = Try-PbiGet "$ws/reports/$($r.id)/datasources"
                $sources = $(if ($src) { Join-NonEmpty "; " (@($src.value) | ForEach-Object { Format-DataSource $_ "$wsName / $($r.name)" }) } else { "Could not read: $script:LastError" })
                $rdl = Get-RdlInfo $r.id $r.name
                $modelNames = foreach ($id in $rdl.Models) { "Semantic model: $(Get-ModelLabel $id '')" }
                $sources = Join-NonEmpty "; " (@($sources) + @($modelNames))
                Add-Row $r.id $r.name "Paginated report" (Get-ItemFolder $r.id) `
                    $null $sources $rdl.Count $rdl.Count $rdl.Detail (Get-ItemAccess $r "report")
            } else {
                Write-Host "  Interactive report: $($r.name)"
                $modelId   = Get-Prop $r "datasetId"
                $modelWsId = $(if (Get-Prop $r "datasetWorkspaceId") { $r.datasetWorkspaceId } else { "$wsId" })
                if ($modelId) {
                    $label = Get-ModelLabel $modelId $modelWsId
                    $sources  = $null   # left blank: the model is named in RefreshSchedule, and its sources are on the model's own row
                    $schedule = $null   # the schedule is on the semantic model's own row
                    $detail   = "None of its own: uses semantic model $label"
                } else { $sources = "(no semantic model)"; $schedule = $null; $detail = "n/a" }
                Add-Row $r.id $r.name "Interactive report" (Get-ItemFolder $r.id) $schedule $sources $null $null $detail (Get-ItemAccess $r "report")
            }
        }

        foreach ($a in $wsApps) {
            Write-Host "  App: $($a.name)"
            $users = Try-PbiGet "admin/apps/$($a.id)/users"
            $audience = $(if ($users) {
                $text = Join-NonEmpty "; " (@($users.value) | Where-Object { $_ -and "$(Get-Prop $_ 'appUserAccessRight')" -ne "None" } | ForEach-Object {
                    "$(Format-Principal $_): $(Format-AccessRight "$($_.appUserAccessRight)")"
                } | Sort-Object)
                $(if ($text) { $text } else { "(no users returned)" })
            } else { "Could not read: $script:LastError" })
            $contents = @($inAppById.Keys | ForEach-Object {
                $id = $_
                if ($inAppById[$id] -eq $a.name) {
                    $item = (@($reports) + @($dashboards)) | Where-Object { $_.id -eq $id } | Select-Object -First 1
                    if ($item) { $(if (Get-Prop $item "name") { $item.name } else { $item.displayName }) }
                }
            } | Sort-Object)
            $lu = Get-Prop $a 'lastUpdate'   # PowerShell 7 reads it as a date, 5.1 as text
            $updated = $(if ($lu -is [datetime]) { $lu.ToString("yyyy-MM-dd") } else { "$lu".Split('T')[0] })
            $detail = Join-NonEmpty "; " @(
                $(if (Get-Prop $a "description") { "Description: $($a.description)" }),
                "Published by $(Get-Prop $a 'publishedBy')$(if ($updated) { ", last updated $updated" })",
                $(if ($appContentsUnknown.ContainsKey($a.name)) { "Contents unknown: $($appContentsUnknown[$a.name])" }
                  else { "Contains $($contents.Count) items: $($contents -join ', ')" }),
                $(if ($appUnmatched.ContainsKey($a.name)) { "also lists items not matched to this workspace: $($appUnmatched[$a.name] -join ', ')" }))
            Add-Row $a.id $a.name "App" $null $null $null $null $null $detail $audience $null
        }

        foreach ($db in $dashboards) {
            Write-Host "  Dashboard: $($db.displayName)"
            $tiles = @(Get-Prop $db "tiles") | Where-Object { $_ }
            $models = foreach ($id in @($tiles | ForEach-Object { Get-Prop $_ "datasetId" } | Where-Object { $_ } | Select-Object -Unique)) {
                "Semantic model $(Get-ModelLabel $id '')"
            }
            Add-Row $db.id $db.displayName "Dashboard" (Get-ItemFolder $db.id) `
                "Tiles refresh with their semantic models" (Join-NonEmpty "; " $models) $null $null "$($tiles.Count) tiles; queries are in the semantic models" (Get-ItemAccess $db "dashboard")
        }
    } catch {
        # One workspace failing doesn't stop the others; it gets an error row in the output.
        $inv = $_.InvocationInfo
        $msg = "$($_.Exception.Message)"
        # The line is shown for unexpected errors only, not for the script's own messages (throw "...").
        $where = $(if ($inv -and $inv.ScriptLineNumber -and $_.FullyQualifiedErrorId -ne $msg) { " (line $($inv.ScriptLineNumber): $("$($inv.Line)".Trim()))" } else { "" })
        Write-Host "  ERROR: workspace $wsId skipped: $msg$where" -ForegroundColor Red
        $failed.Add([pscustomobject]@{ WorkspaceId = $wsId; Error = $msg })
        $rows.Add([pscustomobject]@{
            WorkspaceName = $(if ($wsName) { $wsName } else { "(could not read)" }); WorkspaceId = $wsId
            ItemName = $null; ItemId = $null; ItemType = "Workspace error"; Folder = $null
            InApp = $null; WorkspacePermissions = $null; ItemAccess = $null; RefreshSchedule = $null; Security = $null; DataSources = $null
            TableCount = $null; QueryCount = $null; QueryDetail = "$msg$where"
        })
    }
}

# ---------------------------------------------------------------------------
# Gateway rows: one per gateway used by the items read, with everything this
# account can see about it, its connections and which items use them.
# ---------------------------------------------------------------------------

foreach ($gKey in @($script:gatewayUse.Keys | Sort-Object)) {
    $g = $script:gatewayById[$gKey]
    $name = $(if ($g -and (Get-Prop $g "displayName")) { $g.displayName } elseif ($g) { "Personal gateway" } else { "Gateway $gKey" })
    Write-Host ""
    Write-Host "Gateway: $name" -ForegroundColor Cyan

    # The gateway itself
    $settings = @()
    if ($g) {
        $settings += "Type: $(Get-Prop $g 'type')"
        if (Get-Prop $g "version") { $settings += "version $($g.version)" }
        if ($null -ne (Get-Prop $g "numberOfMemberGateways")) { $settings += "$($g.numberOfMemberGateways) cluster members" }
        if (Get-Prop $g "loadBalancingSetting") { $settings += "load balancing: $($g.loadBalancingSetting)" }
        if ($null -ne (Get-Prop $g "allowCloudConnectionRefresh")) { $settings += "cloud connection refresh allowed: $($g.allowCloudConnectionRefresh)" }
        if ($null -ne (Get-Prop $g "allowCustomConnectors")) { $settings += "custom connectors allowed: $($g.allowCustomConnectors)" }
        $vnet = Get-Prop $g "virtualNetworkAzureResource"
        if ($vnet) { $settings += "virtual network: subscription $(Get-Prop $vnet 'subscriptionId'), resource group $(Get-Prop $vnet 'resourceGroupName'), network $(Get-Prop $vnet 'virtualNetworkName'), subnet $(Get-Prop $vnet 'subnetName')" }
        if (Get-Prop $g "capacityId") { $settings += "capacity $($g.capacityId)" }
        if ($null -ne (Get-Prop $g "inactivityMinutesBeforeSleep")) { $settings += "sleeps after $($g.inactivityMinutesBeforeSleep) minutes idle" }
        if ("$(Get-Prop $g 'type')" -eq "OnPremises") {
            try {
                $members = @(Get-FabricPaged "$FabricApiUrl/gateways/$gKey/members")
                $settings += "members: " + (($members | ForEach-Object {
                    "$(Get-Prop $_ 'displayName') (version $(Get-Prop $_ 'version'), $(if ((Get-Prop $_ 'enabled') -eq $false) { 'disabled' } else { 'enabled' }))"
                }) -join ", ")
            } catch { $settings += "members: could not read ($(Get-ShortError $_))" }
        }
    } else {
        $settings += "No access to this gateway's details: a gateway admin can add you as Connection creator or Admin to show its name, type, version and members"
    }

    # Who can manage or use the gateway
    $admins = $(if ($g) {
        try {
            Join-NonEmpty "; " (@(Get-FabricPaged "$FabricApiUrl/gateways/$gKey/roleAssignments") | ForEach-Object {
                "$(Format-FabricPrincipal $_.principal): $($_.role)" } | Sort-Object)
        } catch { "Could not read gateway roles: $(Get-ShortError $_) (needs Connection creator or Admin on the gateway)" }
    } else { "Not available: no access to this gateway" })

    # Its connections used by the items read
    $connections = foreach ($cKey in @($script:gatewayUse[$gKey].Keys | Sort-Object)) {
        $c = $script:connectionById[$cKey]
        $source = $script:gatewaySource[$cKey]
        $usedBy = ($script:gatewayUse[$gKey][$cKey] | Sort-Object) -join ", "
        if ($c) {
            $cd = Get-Prop $c "connectionDetails"
            $lastUsed = Get-Prop (Get-Prop $c "connectionRecency") "lastCredentialUsedDateTime"
            $lastUsed = $(if ($lastUsed -is [datetime]) { $lastUsed.ToString("yyyy-MM-dd") } elseif ($lastUsed) { "$lastUsed".Split('T')[0] })
            $users = try {
                Join-NonEmpty ", " (@(Get-FabricPaged "$FabricApiUrl/connections/$cKey/roleAssignments") | ForEach-Object {
                    "$(Format-FabricPrincipal $_.principal): $($_.role)" } | Sort-Object)
            } catch { "could not read ($(Get-ShortError $_))" }
            Join-NonEmpty "; " @(
                "$(Get-ConnectionLabel $cKey)".Substring("connection: ".Length),
                "$(Get-Prop $cd 'type'): $(Get-Prop $cd 'path')",
                $(if ($lastUsed) { "credentials last used $lastUsed" }),
                $(if ($users) { "users: $users" }),
                "used by: $usedBy")
        } else {
            $label = $(if ($cKey -eq "(unknown)") { "Connection not identified by the API" } else { "Connection $cKey" })
            Join-NonEmpty "; " @($label, $source, "used by: $usedBy")
        }
    }

    $rows.Add([pscustomobject]@{
        WorkspaceName = "(all workspaces read)"; WorkspaceId = $null
        ItemName = $name; ItemId = $gKey; ItemType = "Gateway"; Folder = $null; InApp = $null
        WorkspacePermissions = $null; ItemAccess = $admins; RefreshSchedule = $null
        DataSources = (@($connections) -join " | "); TableCount = $null; QueryCount = $null
        QueryDetail = ($settings -join "; ")
    })
}

# ---------------------------------------------------------------------------
# 8. Output: one CSV, one row per item, everything about the item on that row
# ---------------------------------------------------------------------------

$maxCell = 32000   # Excel shows at most 32,767 characters in a cell
function Limit-Cell([string] $Text) {
    if ($Text.Length -le $maxCell) { return $Text }
    $Text.Substring(0, $maxCell) + "`n... (cut at $maxCell characters; Excel can't show more in one cell)"
}

[hashtable] $byItem = @{}
foreach ($t in $tableRows) {
    $key = "$($t.WorkspaceId)|$($t.ItemId)"
    if (-not $byItem.ContainsKey($key)) { $byItem[$key] = New-Object System.Collections.Generic.List[object] }
    $byItem[$key].Add($t)
}

$output = foreach ($r in $rows) {
    $parts = @()
    $key = "$($r.WorkspaceId)|$($r.ItemId)"
    if ($r.ItemId -and $byItem.ContainsKey($key)) { $parts = @($byItem[$key] | Sort-Object Kind, Name) }
    $tablesText = Join-NonEmpty "; " ($parts | ForEach-Object {
        $info = Join-NonEmpty ", " @($_.Kind, $_.Sources, $(if ($_.Hidden) { "hidden" }))
        "$($_.Name) [$info]"
    })
    $queriesText = ($parts | Where-Object { $_.Query } | ForEach-Object { "// $($_.Name) ($($_.Kind))`n$($_.Query)" }) -join "`n`n"
    [pscustomobject]@{
        WorkspaceName        = $r.WorkspaceName
        WorkspaceId          = $r.WorkspaceId
        ItemName             = $r.ItemName
        ItemType             = $r.ItemType
        ItemId               = $r.ItemId
        Folder               = $r.Folder
        InApp                = $r.InApp
        Owner                = $r.Owner
        Created              = $r.Created
        ModifiedBy           = $r.ModifiedBy
        Modified             = $r.Modified
        Endorsement          = $r.Endorsement
        SensitivityLabel     = $r.SensitivityLabel
        Description          = Limit-Cell $r.Description
        WorkspacePermissions = Limit-Cell $r.WorkspacePermissions
        ItemAccess           = Limit-Cell $r.ItemAccess
        RefreshSchedule      = $r.RefreshSchedule
        Security             = Limit-Cell $r.Security
        DataSources          = Limit-Cell $r.DataSources
        TableCount           = $r.TableCount
        QueryCount           = $r.QueryCount
        QueryDetail          = $r.QueryDetail
        Tables               = Limit-Cell $tablesText
        Queries              = Limit-Cell $queriesText
    }
}
$output = @($output | Sort-Object WorkspaceName, ItemType, Folder, ItemName)

New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
$stamp = Get-Date -Format "yyyyMMdd-HHmmss"
$outFile = Join-Path $OutputFolder "PowerBI-inventory-$stamp.csv"
# UTF-8 with a byte order mark, so Excel shows accented characters correctly
$encoding = $(if ($PSVersionTable.PSVersion.Major -ge 6) { "utf8BOM" } else { "UTF8" })
$output | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding $encoding

$output | Format-Table WorkspaceName, ItemName, ItemType, Folder, TableCount, QueryCount -AutoSize -Wrap | Out-Host

# ---------------------------------------------------------------------------
# 9. Tenant settings snapshot (Fabric admin only): every Admin portal tenant
#    setting, plus the capacity, domain and workspace overrides, in its own file.
# ---------------------------------------------------------------------------

$settingsFile = $null
$settingsNote = $null
if ($SkipTenantSettings)                          { $settingsNote = "skipped (-SkipTenantSettings)" }
elseif ($NoAdminScan -or $script:adminScanRefused) { $settingsNote = "not saved: needs a Fabric admin account" }
else {
    Write-Host ""
    Write-Host "Reading tenant settings ..." -ForegroundColor Cyan

    function ConvertTo-SettingRows($Settings, [string] $Scope, [string] $ScopeName) {
        foreach ($t in @($Settings) | Where-Object { $_ }) {
            $enabled  = (Get-Prop $t "enabled") -eq $true
            $groups   = Join-NonEmpty ", " (@(Get-Prop $t "enabledSecurityGroups")  | Where-Object { $_ } | ForEach-Object { Get-Prop $_ "name" })
            $excluded = Join-NonEmpty ", " (@(Get-Prop $t "excludedSecurityGroups") | Where-Object { $_ } | ForEach-Object { Get-Prop $_ "name" })
            $delegated = Join-NonEmpty ", " @(
                $(if ((Get-Prop $t "delegateToCapacity")  -eq $true) { "capacity admins" }),
                $(if ((Get-Prop $t "delegateToDomain")    -eq $true) { "domain admins" }),
                $(if ((Get-Prop $t "delegateToWorkspace") -eq $true) { "workspace admins" }))
            [pscustomobject]@{
                Scope        = $Scope
                ScopeName    = $ScopeName
                Group        = Get-Prop $t "tenantSettingGroup"
                Title        = Get-Prop $t "title"
                SettingName  = Get-Prop $t "settingName"
                Enabled      = $(if ($enabled) { "Enabled" } else { "Disabled" })
                AppliesTo    = $(if (-not $enabled) { $null } elseif ($groups) { "Security groups: $groups" } else { "Entire organization" })
                ExceptGroups = $excluded
                CanBeOverriddenBy = $delegated
                DelegatedFrom = Get-Prop $t "delegatedFrom"
                Properties   = Join-NonEmpty "; " (@(Get-Prop $t "properties") | Where-Object { $_ } | ForEach-Object { "$(Get-Prop $_ 'name'): $(Get-Prop $_ 'value')" })
            }
        }
    }

    $settingRows = New-Object System.Collections.Generic.List[object]
    try {
        foreach ($row in ConvertTo-SettingRows @(Get-FabricPaged "$FabricApiUrl/admin/tenantsettings") "Tenant" "") { $settingRows.Add($row) }

        # Names for the override scopes, where they can be read
        $capacityName = @{}
        $caps = Try-PbiGet "admin/capacities"
        if ($caps) { foreach ($c in @($caps.value)) { $capacityName["$($c.id)".ToLowerInvariant()] = "$(Get-Prop $c 'displayName') ($(Get-Prop $c 'sku'), $(Get-Prop $c 'region'))" } }
        $domainName = @{}
        try { foreach ($d in Get-FabricPaged "$FabricApiUrl/admin/domains" "domains") { $domainName["$($d.id)".ToLowerInvariant()] = Get-Prop $d "displayName" } } catch { }

        foreach ($scope in @(
            @{ Kind = "Capacity";  Url = "admin/capacities/delegatedTenantSettingOverrides"; Names = $capacityName },
            @{ Kind = "Domain";    Url = "admin/domains/delegatedTenantSettingOverrides";    Names = $domainName },
            @{ Kind = "Workspace"; Url = "admin/workspaces/delegatedTenantSettingOverrides"; Names = @{} })) {
            try {
                foreach ($o in Get-FabricPaged "$FabricApiUrl/$($scope.Url)") {
                    $id = "$(Get-Prop $o 'id')"
                    $label = $(if ($scope.Names.ContainsKey($id.ToLowerInvariant())) { "$($scope.Names[$id.ToLowerInvariant()]) [$id]" } else { $id })
                    foreach ($row in ConvertTo-SettingRows (Get-Prop $o "tenantSettings") "$($scope.Kind) override" $label) { $settingRows.Add($row) }
                }
            } catch { Write-Warning "$($scope.Kind) overrides couldn't be read: $(Get-ShortError $_)" }
        }

        $settingsFile = Join-Path $OutputFolder "PowerBI-tenant-settings-$stamp.csv"
        $settingRows | Sort-Object Scope, ScopeName, Group, Title | Export-Csv -LiteralPath $settingsFile -NoTypeInformation -Encoding $encoding
        $overrides = @($settingRows | Where-Object { $_.Scope -ne "Tenant" }).Count
        $settingsNote = "$($settingRows.Count - $overrides) settings, $overrides overrides"
    } catch {
        $settingsNote = "not saved: $(Get-ShortError $_)"
        Write-Warning "Tenant settings couldn't be read: $(Get-ShortError $_)"
    }
}

$okCount = $workspaceIds.Count - $failed.Count
Write-Host "Workspaces read: $okCount of $($workspaceIds.Count)"
foreach ($f in $failed) { Write-Host "  Not read: $($f.WorkspaceId): $($f.Error)" -ForegroundColor Red }
Write-Host "Items: $(@($output | Where-Object { $_.ItemType -ne 'Workspace error' }).Count)"
Write-Host "Saved: $outFile"
if ($settingsFile) { Write-Host "       $settingsFile ($settingsNote)" } else { Write-Host "Tenant settings: $settingsNote" }
Write-Host "This file contains connection details and queries: keep it within the team." -ForegroundColor Yellow

if ($PassThru) { $output }
