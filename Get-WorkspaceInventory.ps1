<#
.SYNOPSIS
    Documents every item in one or more Power BI workspaces: folder, workspace
    permissions, refresh schedule, data sources, and the tables and queries
    each semantic model or paginated report is built from. One CSV, one row
    per item.

.DESCRIPTION
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
                                       paginated report or dashboard
        Folder                         workspace folder path, e.g. Finance/Monthly
        WorkspacePermissions           who has which workspace role
        ItemAccess                     who has access to this item other than
                                       through a workspace role (shared directly,
                                       or a Viewer given more than Read), and with
                                       which permissions. All item types with a
                                       Fabric admin account; semantic models only
                                       with a standard account (needs Admin or
                                       Member on the workspace)
        RefreshSchedule                days, times, time zone
        DataSources                    servers, databases, URLs, gateways, dataflows
                                       (blank for an interactive report that uses a
                                       semantic model: see the model's own row)
        TableCount, QueryCount,        counts (see below) and their breakdown
        QueryDetail
        Tables                         each table, shared query or parameter (or
                                       paginated report dataset) with its kind
                                       and connectors
        Queries                        the full M, DAX or SQL of each of them

    A workspace that can't be read gets one "Workspace error" row with the
    reason, and the script carries on with the next workspace. Cells longer
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
    [switch] $PassThru
)

$ErrorActionPreference = "Stop"
$ScriptVersion = "2026-11-23.2"

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
    throw "The MicrosoftPowerBIMgmt module isn't installed. See power-bi-api-powershell-vscode.md."
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
    if ($msg -match '401|Unauthorized') { return "not signed in or not allowed (401)" }
    if ($msg -match '403|Forbidden')    { return "no permission (403)" }
    if ($msg -match '404|NotFound')     { return "not found (404)" }
    ($msg -split "`n")[0].Trim()
}

function Get-Prop($Object, [string] $Name) {
    if ($null -ne $Object -and $Object.PSObject.Properties[$Name]) { $Object.$Name } else { $null }
}

function Join-NonEmpty([string] $Separator, $Values) {
    (@($Values) | Where-Object { $_ -and "$_".Trim() } | Select-Object -Unique) -join $Separator
}

function Format-DataSource($S) {
    $c = Get-Prop $S "connectionDetails"
    $where = Join-NonEmpty " / " @((Get-Prop $c "server"), (Get-Prop $c "database"), (Get-Prop $c "url"),
                                   (Get-Prop $c "path"), (Get-Prop $c "account"), (Get-Prop $c "kind"))
    $gw = $(if (Get-Prop $S "gatewayId") { " [gateway]" } else { "" })
    if ($where) { "$($S.datasourceType): $where$gw" } else { "$($S.datasourceType)$gw" }
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
                            datasourceId = $id; datasourceType = Get-Prop $x "datasourceType"
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

        function Get-FabricPaged([string] $Url, [hashtable] $Headers, [string] $ListProperty = "value") {
            $next = $Url
            while ($next) {
                $page = Invoke-RestMethod -Uri $next -Method Get -Headers $Headers -ErrorAction Stop
                foreach ($v in @(Get-Prop $page $ListProperty)) { if ($v) { $v } }
                $next = Get-Prop $page "continuationUri"
            }
        }

        [hashtable] $folderOfItem = @{}
        $folderRows   = @()
        $folderSource = "Fabric API"
        $headers = @{ Authorization = (Get-PowerBIAccessToken -AsString) }

        # Folder names: needs a role in the workspace (Viewer or higher).
        [hashtable] $byId = @{}
        try {
            foreach ($f in Get-FabricPaged "$FabricApiUrl/workspaces/$wsId/folders" $headers) { $byId[$f.id] = $f }
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
        try { $items = @(Get-FabricPaged "$FabricApiUrl/workspaces/$wsId/items" $headers) }
        catch {
            try {
                if (-not $isFabricAdmin) { throw }
                $items = @(Get-FabricPaged "$FabricApiUrl/admin/items?workspaceId=$wsId" $headers "itemEntities")
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
            [pscustomobject]@{ Tables = @($tables); Expressions = @($exprs) }
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
            [pscustomobject]@{ Tables = @($tables); Expressions = @($exprs) }
        }

        function Get-DefinitionModelParts([string] $ModelId) {
            $parts = Get-ModelDefinitionParts $ModelId
            $tableKeys = @($parts.Keys | Where-Object { $_ -like "definition/tables/*.tmdl" })
            if ($tableKeys.Count) {
                $tables = foreach ($k in $tableKeys) { ConvertFrom-TmdlTable $parts[$k] }
                $exprs  = @()
                if ($parts.ContainsKey("definition/expressions.tmdl")) { $exprs = @(ConvertFrom-TmdlExpressions $parts["definition/expressions.tmdl"]) }
                return [pscustomobject]@{ Tables = @($tables); Expressions = $exprs }
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
                return $info
            }
            if ($SkipModelDefinition) {
                return [pscustomobject]@{ Tables = $null; Queries = $null; Detail = "Unknown (the scan returned no tables and -SkipModelDefinition was used)" }
            }
            try {
                $m = Get-DefinitionModelParts $D.id
                return Add-ModelRows $D.id $D.name $m.Tables $m.Expressions "model definition"
            } catch {
                $why = Get-ShortError $_
                if ($why -match '403') { $why = "$why; reading a model's definition needs Contributor or higher on the workspace" }
                return [pscustomobject]@{ Tables = $null; Queries = $null; Detail = "Unknown (scan returned no tables; definition: $why)" }
            }
        }

        function Get-ModelSources($D) {
            $list = foreach ($u in @(Get-Prop $D "datasourceUsages") | Where-Object { $_ }) {
                if ($sourceById.ContainsKey($u.datasourceInstanceId)) { Format-DataSource $sourceById[$u.datasourceInstanceId] }
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

        function Add-Row($Id, $Name, $Type, $Folder, $Schedule, $Sources, $TableCount, $QueryCount, $QueryDetail, $Access) {
            $rows.Add([pscustomobject]@{
                WorkspaceName        = $wsName
                WorkspaceId          = $wsId
                ItemName             = $Name
                ItemId               = $Id
                ItemType             = $Type
                Folder               = $Folder
                WorkspacePermissions = $permissionSummary
                ItemAccess           = $Access
                RefreshSchedule      = $Schedule
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
                (Get-ModelSchedule "$wsId" $d.id) (Get-ModelSources $d) $q.Tables $q.Queries $q.Detail (Get-ItemAccess $d "dataset")
        }

        foreach ($r in $reports) {
            if ((Get-Prop $r "reportType") -eq "PaginatedReport") {
                Write-Host "  Paginated report: $($r.name)"
                $src = Try-PbiGet "$ws/reports/$($r.id)/datasources"
                $sources = $(if ($src) { Join-NonEmpty "; " (@($src.value) | ForEach-Object { Format-DataSource $_ }) } else { "Could not read: $script:LastError" })
                $rdl = Get-RdlInfo $r.id $r.name
                $modelNames = foreach ($id in $rdl.Models) { "Semantic model: $(Get-ModelLabel $id '')" }
                $sources = Join-NonEmpty "; " (@($sources) + @($modelNames))
                Add-Row $r.id $r.name "Paginated report" (Get-ItemFolder $r.id) `
                    "None (runs its queries when opened; check subscriptions in the service)" $sources $rdl.Count $rdl.Count $rdl.Detail (Get-ItemAccess $r "report")
            } else {
                Write-Host "  Interactive report: $($r.name)"
                $modelId   = Get-Prop $r "datasetId"
                $modelWsId = $(if (Get-Prop $r "datasetWorkspaceId") { $r.datasetWorkspaceId } else { "$wsId" })
                if ($modelId) {
                    $label = Get-ModelLabel $modelId $modelWsId
                    $sources  = $null   # left blank: the model is named in RefreshSchedule, and its sources are on the model's own row
                    $schedule = "From semantic model ${label}: $(Get-ModelSchedule $modelWsId $modelId)"
                    $detail   = "None of its own: uses semantic model $label"
                } else { $sources = "(no semantic model)"; $schedule = "n/a"; $detail = "n/a" }
                Add-Row $r.id $r.name "Interactive report" (Get-ItemFolder $r.id) $schedule $sources $null $null $detail (Get-ItemAccess $r "report")
            }
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
            WorkspacePermissions = $null; ItemAccess = $null; RefreshSchedule = $null; DataSources = $null
            TableCount = $null; QueryCount = $null; QueryDetail = "$msg$where"
        })
    }
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
        WorkspacePermissions = Limit-Cell $r.WorkspacePermissions
        ItemAccess           = Limit-Cell $r.ItemAccess
        RefreshSchedule      = $r.RefreshSchedule
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
$outFile = Join-Path $OutputFolder ("PowerBI-inventory-{0}.csv" -f (Get-Date -Format "yyyyMMdd-HHmmss"))
# UTF-8 with a byte order mark, so Excel shows accented characters correctly
$encoding = $(if ($PSVersionTable.PSVersion.Major -ge 6) { "utf8BOM" } else { "UTF8" })
$output | Export-Csv -LiteralPath $outFile -NoTypeInformation -Encoding $encoding

$output | Format-Table WorkspaceName, ItemName, ItemType, Folder, TableCount, QueryCount -AutoSize -Wrap | Out-Host

$okCount = $workspaceIds.Count - $failed.Count
Write-Host "Workspaces read: $okCount of $($workspaceIds.Count)"
foreach ($f in $failed) { Write-Host "  Not read: $($f.WorkspaceId): $($f.Error)" -ForegroundColor Red }
Write-Host "Items: $(@($output | Where-Object { $_.ItemType -ne 'Workspace error' }).Count)"
Write-Host "Saved: $outFile"
Write-Host "This file contains connection details and queries: keep it within the team." -ForegroundColor Yellow

if ($PassThru) { $output }
