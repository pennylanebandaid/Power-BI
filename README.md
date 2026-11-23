# Workspace inventory with PowerShell

Get-WorkspaceInventory.ps1 documents what's in one or more Power BI workspaces, as they are in the service. It writes **one CSV with one row per item**: semantic models, interactive reports, paginated reports, dashboards, the workspace app, and the gateways they use. Everything about an item is on its row: folder, permissions, refresh schedule, subscriptions, data sources with their gateways and connections, and the tables and queries it's built from.

It only reads. Nothing in Power BI is created, changed, refreshed or run.

## What you get

One CSV with these columns. Open it in Excel; long text (queries) shows on several lines inside one cell.

| Column | What it holds |
|---|---|
| WorkspaceName, WorkspaceId | The workspace |
| ItemName, ItemType, ItemId | The item: *Semantic model*, *Interactive report*, *Paginated report*, *Dashboard*, *App*, *Gateway*, or *Workspace error* |
| Folder | Workspace folder path, e.g. Finance/Monthly, or (workspace root) |
| InApp | Reports and dashboards: the workspace app they're published in, or No |
| Owner | Semantic model: who it's configured by. Report: who created it. App: who published it |
| Created, ModifiedBy, Modified | When the item was created, and who last changed it and when. Times are in UTC |
| Endorsement | *Promoted* or *Certified* (with who certified it), or blank |
| SensitivityLabel | The sensitivity label's ID. Label names are in Microsoft Purview |
| Description | The item's description |
| WorkspacePermissions | Who has which workspace role, with email addresses |
| ItemAccess | Who has access other than through a workspace role (shared directly, or a Viewer given more than Read), with their permissions. On an *App* row: everyone the app is shared with. On a *Gateway* row: its admins and connection creators |
| RefreshSchedule | Days, times, time zone. Reports show their model's schedule |
| Subscriptions | Reports and dashboards: each email subscription with frequency, state, format, dates and recipients |
| Security | Semantic models: each row-level security role with its filters (DAX), any object-level security (hidden tables and columns), and its members. Members are people and groups in this tenant, so each needs a matching identity in a new one |
| DataSources | Servers, databases, URLs, dataflows. A source that goes through a gateway also shows the gateway and connection, e.g. Sql: srv / db [gateway: Finance-OnPrem (OnPremises); connection: Finance SQL (Windows, privacy Organizational)]. Blank for an interactive report on a semantic model: see the model's row. On a *Gateway* row: each connection with its users and the items that use it |
| TableCount, QueryCount | See below |
| QueryDetail | The breakdown behind the counts, or why they're unknown. On an *App* row: description, publisher, last update and contents. On a *Gateway* row: type, version, cluster members and settings |
| Tables | Each table, shared query and parameter (or paginated report dataset), with its kind and connectors |
| Queries | The full M, DAX or SQL of each of them |

### Sample

Selected columns from sample-output.csv, a run with a Fabric admin account against made-up data:

| ItemName | ItemType | Folder | InApp | ItemAccess | RefreshSchedule | Subscriptions | TableCount | QueryCount |
|---|---|---|---|---|---|---|---|---|
| Finance-OnPrem | Gateway | | | BI Developers (group): ConnectionCreator; Pat Admin (pat@x.com): Admin | | | | |
| Gateway 9b9b9b9b-… | Gateway | | | Not available: no access to this gateway | | | | |
| Finance | App | | | Finance Readers (group): Read; Jane Smith (jane@x.com): Read, Build; Pat Admin (pat@x.com): All | | | | |
| Exec | Dashboard | Finance | Finance | None beyond workspace roles | Tiles refresh with their semantic models | None | | |
| Budget Report | Interactive report | Finance/Monthly | Finance | Entire organization: Read; Sales Team (group): Read, Reshare | From semantic model 'Budget': Scheduled refresh on; Monday,…,Friday; 07:00,13:00; Eastern Standard Time | Weekly budget (Weekly, enabled, page 'Summary'); to: cfo@x.com | | |
| Invoice | Paginated report | Finance | Finance | Bob (bob@x.com): Read | None (runs its queries when opened; see Subscriptions) | Monthly invoices (Monthly, enabled, PDF, 1/1/2026 to 12/31/2026); to: ap@x.com, Finance Team (group) \| Old one (Daily, disabled, Excel) | 2 | 2 |
| Budget | Semantic model | (workspace root) | | None beyond workspace roles | Scheduled refresh on; Monday,…,Friday; 07:00,13:00; Eastern Standard Time | | 4 | 3 |
| Sales | Semantic model | Finance/Monthly | | Finance Viewers (group): Viewer role; on this item: Read, Build; Jane Smith (jane@x.com): Read, Reshare, Build | Scheduled refresh on; Monday,Friday; 06:00; Eastern Standard Time | | 2 | 2 |

DataSources from the same run:

| Item | DataSources |
|---|---|
| Sales (semantic model) | Sql: srv / db [gateway: Finance-OnPrem (OnPremises); connection: Finance SQL (Windows, privacy Organizational)] |
| Invoice (paginated report) | Oracle: ora-prod [gateway 9b9b9b9b-… (no access to its details); connection c2c2c2c2-… (no access to its details)]; Sql: srv / db [gateway: Finance-OnPrem (OnPremises); connection: Finance SQL (Windows, privacy Organizational)]; Semantic model: 'Sales' |
| Finance-OnPrem (gateway) | Finance SQL (Windows, privacy Organizational); SQL: srv;db; credentials last used 2026-11-20; users: BI Developers (group): User, Pat Admin (pat@x.com): Owner; used by: Finance Dev / Invoice, Finance Dev / Sales |
| Gateway 9b9b9b9b-… | Connection c2c2c2c2-… (no access to its details); Oracle: ora-prod; used by: Finance Dev / Invoice |

The *Finance-OnPrem* gateway's QueryDetail: *Type: OnPremises; version 3000.290.7; 2 cluster members; load balancing: Failover; cloud connection refresh allowed: False; custom connectors allowed: True; members: FINGW01 (version 3000.290.7, enabled), FINGW02 (version 3000.280.5, disabled)*.

The *Budget* model row, columns not shown above:

| Column | Value |
|---|---|
| Security | Role 'No Salary' (read): no filters; hides (object-level security): Budget[Amount], Date Table; members: none assigned \| Role 'West Region' (read): filters: Budget: [Region] = "West"; Fact Sales: VAR r = USERPRINCIPALNAME() RETURN [Owner] = r; members: jane@x.com, Sales West (group) |
| DataSources | From M code: Sql.Database, Excel.Workbook, File.Contents, Csv.Document |
| QueryDetail | 2 Power Query tables, 1 DAX calculated tables, 1 calculation groups, 1 shared queries, 1 parameters, 0 auto date tables (not counted) (from model definition) |
| Tables | Time Intel [Calculation group]; Date Table [DAX calculated table]; RootPath [Parameter]; Budget [Power Query, Excel.Workbook, File.Contents]; Fact Sales [Power Query (incremental refresh), Sql.Database, hidden]; Helper Query [Shared query (not loaded as a table), Csv.Document, File.Contents] |
| Queries | // Budget (Power Query) followed by its M code, then the same for each table, query and parameter |

### Tenant settings snapshot

Signed in with the **Fabric admin account**, the script also saves PowerBI-tenant-settings-date-time.csv, a record of the Admin portal
settings to set up the new tenant the same way. One row per setting:

| Column | What it holds |
|---|---|
| Scope, ScopeName | *Tenant*, or a *Capacity*, *Domain* or *Workspace override* with that capacity's, domain's or workspace's name and ID |
| Group, Title, SettingName | The setting as it's grouped and named in the Admin portal, and its internal name |
| Enabled | *Enabled* or *Disabled* |
| AppliesTo, ExceptGroups | *Entire organization* or the security groups it applies to, and any groups excluded |
| CanBeOverriddenBy | Capacity, domain or workspace admins, if the setting is delegated to them |
| DelegatedFrom | For an override: where it was delegated from |
| Properties | The setting's extra options, such as a URL or a number |

From sample-tenant-settings.csv:

| Scope | ScopeName | Title | Enabled | AppliesTo | ExceptGroups | CanBeOverriddenBy |
|---|---|---|---|---|---|---|
| Tenant | | Publish to web | Enabled | Security groups: Web Publishers | Contractors | |
| Tenant | | Export to Excel | Enabled | Entire organization | | capacity admins, workspace admins |
| Capacity override | Finance P1 (P1, East US) [CAP1] | Export to Excel | Disabled | | | |
| Domain override | Finance [dom1] | Certification | Enabled | Security groups: Finance Certifiers | | |

With a standard account, or with -NoAdminScan or -SkipTenantSettings, no snapshot is saved.

## What each account can see

| Column | Standard account, Contributor or higher | Fabric admin, no workspace role |
|---|---|---|
| Items, WorkspacePermissions, DataSources | ✅ | ✅ |
| Folder | ✅ | Folder IDs instead of names |
| RefreshSchedule | ✅ | Only on Premium or Fabric capacity |
| Model TableCount, QueryCount, Tables, Queries, Security | ✅ | ❌ Unknown |
| Paginated report datasets and queries | ✅ | ❌ Unknown |
| ItemAccess: semantic models | ✅ with Admin or Member | ✅ |
| ItemAccess: reports and dashboards | ❌ No API for a standard account | ✅ |
| Subscriptions | ❌ | ✅ |
| InApp and *App* rows | ❌ | ✅ |
| Owner, Created, Endorsement, SensitivityLabel, Description | Semantic models only (reports: Description only; dashboards: none) | ✅ All item types |
| ModifiedBy, Modified | ❌ | Reports and apps |

To fill **every** column in one run, use the Fabric admin account and give it Contributor or higher on the workspaces. Otherwise, run it once with each account.

**Gateways and connections** have their own permissions, separate from workspace roles and from the Fabric admin role. Each account sees only the gateways and connections it has a role on: a gateway's **Admin** or **Connection creator** role shows its details and roles, and a connection's **User** or **Owner** role shows the connection. You can check which ones an account has in the service under **Settings › Manage connections and gateways**. Gateways and connections the account can't see still appear, by ID, with the items that use them.

A cell that can't be filled says why, for example *Not available without a Fabric admin account* or *Could not read: no permission (403)*.

## How the counts work

| Item | TableCount | QueryCount |
|---|---|---|
| Semantic model | Tables in the finished model, not counting the hidden auto date/time tables Power BI adds | Power Query queries: tables loaded by Power Query plus shared queries (staging, helper queries, functions) not loaded as tables. Parameters and DAX calculated tables are in QueryDetail but not counted |
| Paginated report | Datasets in the .rdl | Same |
| Interactive report, dashboard | Blank: their queries are in the semantic model | Blank |

QueryDetail ends with where the counts came from: *(from admin scan)* or *(from model definition)*.

## Calls it makes

Every call is a GET except two POSTs, which only read:

| Call | When | Why it's a POST |
|---|---|---|
| Start the admin scan (admin/workspaces/getInfo) | Fabric admin, once per workspace | The scan is started, then its result is read with GETs |
| Get the model's definition (getDefinition) | Once per semantic model, when the scan has no tables | Microsoft made this call a POST; it returns a copy of the definition |

Admin calls for the tenant settings snapshot (once per run, Fabric admin only): list tenant settings (25 a minute), list capacities and domains for
their names, and list capacity, domain and workspace overrides.

Admin calls with hourly limits:

| Call | Limit | Used for |
|---|---|---|
| Get report/dashboard subscriptions as admin (preview) | 200 an hour | One per report and dashboard. Past the limit, the rest say *Not read: … run again later* |
| Get apps as admin | 200 an hour | Once per run |
| Get app users as admin | 200 an hour | One per app |

Fabric calls for gateways and connections (GET, 50 a minute each; on *too many requests* the script waits a minute and tries again):

| Call | Used for |
|---|---|
| List gateways, list connections | Once per run |
| List gateway members, list gateway role assignments | Once per gateway used |
| List connection role assignments | Once per connection used |


## Parameters

| Parameter | What it does |
|---|---|
| -WorkspaceId | One or more workspace GUIDs, separated by commas |
| -WorkspaceIdFile | A text file of workspace GUIDs. Can be combined with -WorkspaceId |
| -OutputFolder | Where to save the CSV. Defaults to the current folder. Quote it if it has a space |
| -NoAdminScan | Don't try the admin scan; use your workspace role. Use with a standard account. No admin calls are made, so no tenant settings snapshot |
| -SkipTenantSettings | Don't save the tenant settings snapshot |
| -SkipModelDefinition | Don't read model definitions. Model counts and queries show *Unknown*, and the per-model POSTs aren't made |
| -SkipRdlDownload | Don't download paginated reports. Their dataset counts show *Unknown* |
| -FabricApiUrl | Fabric API address, for folders, model definitions, gateways and connections. Leave the default |
| -PassThru | Also return the rows as objects, e.g. … -PassThru \| Out-GridView |

## What it doesn't cover

- **Members of groups**: groups show as the group, not the people in it.
- **App audiences**: an *App* row lists everyone the app is shared with, but not which audience each is in; no API returns that.
- **Sharing links**, such as *people in your organization with the link*.
- **Empty workspace folders**: only folders that hold items appear, in the Folder column.
- **Gateways no item uses**: *Gateway* rows are only for gateways used by the items read. Cloud connections show in DataSources but get no row of their own.
- **Gateway machine details** beyond the member name and version (server name, Windows service account, recovery key): these are on the gateway machines and in the gateway app, not in any API.
- **Very long cells**: Excel shows at most 32,767 characters in a cell. Longer text, usually Queries, is cut at 32,000 characters with a note.

The CSV holds connection details and queries: keep it within the team.
