<#
.SYNOPSIS
  Read-only audit of Entra ID directory role assignments.

.DESCRIPTION
  Connects to Microsoft Graph (Directory.Read.All) and exports directory
  role memberships to CSV/JSON, with logging and a run summary.

  SAFE BY DESIGN:
  - No write operations
  - No role changes
  - No user modifications

.REQUIREMENTS
  - Microsoft.Graph PowerShell SDK
  - Permission: Directory.Read.All (may require admin consent)

.USAGE
  pwsh ./src/Get-EntraRoleAudit.ps1 -OutputRoot ./reports
#>

[CmdletBinding()]
param(
    [string]$OutputRoot = "./reports",
    [string]$TenantId,
    [switch]$NoDisconnect
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

# ---------------- Helper Functions ----------------

function Ensure-Module {
    param([Parameter(Mandatory)][string]$Name)

    if (-not (Get-Module -ListAvailable -Name $Name)) {
        throw "Required module '$Name' not found. Install with: Install-Module $Name -Scope CurrentUser"
    }
}

function New-RunFolder {
    param([Parameter(Mandatory)][string]$Root)

    if (-not (Test-Path -LiteralPath $Root)) {
        New-Item -ItemType Directory -Path $Root | Out-Null
    }

    $stamp = Get-Date -Format "yyyyMMdd-HHmmss"
    $path  = Join-Path $Root "entra-role-audit-$stamp"

    New-Item -ItemType Directory -Path $path | Out-Null
    return $path
}

function Write-Log {
    param(
        [string]$LogFile,
        [string]$Message,
        [ValidateSet("INFO","WARN","ERROR")]
        [string]$Level = "INFO"
    )

    $line = "{0} [{1}] {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $Message
    Add-Content -Path $LogFile -Value $line

    switch ($Level) {
        "INFO"  { Write-Host $Message }
        "WARN"  { Write-Warning $Message }
        "ERROR" { Write-Error $Message }
    }
}

function Get-AdditionalProp {
    param($Obj, [string]$Key)

    if ($Obj.AdditionalProperties -and $Obj.AdditionalProperties.ContainsKey($Key)) {
        return $Obj.AdditionalProperties[$Key]
    }
    return $null
}

function Connect-GraphReadOnly {
    param([string]$Tenant)

    $scopes = @("Directory.Read.All")

    if ([string]::IsNullOrWhiteSpace($Tenant)) {
        Connect-MgGraph -Scopes $scopes | Out-Null
    } else {
        Connect-MgGraph -TenantId $Tenant -Scopes $scopes | Out-Null
    }

    # Some Graph module versions include Select-MgProfile; others don't.
    # Default is typically v1.0, so this is optional.
    if (Get-Command -Name Select-MgProfile -ErrorAction SilentlyContinue) {
        Select-MgProfile -Name "v1.0" | Out-Null
    }
}

# ---------------- Main Execution ----------------

$runFolder = New-RunFolder -Root $OutputRoot
$logFile   = Join-Path $runFolder "run.log"

Write-Log -LogFile $logFile -Message "Starting Entra role audit"
Write-Log -LogFile $logFile -Message "Host: $env:COMPUTERNAME | PowerShell: $($PSVersionTable.PSVersion)"

try {
    Ensure-Module -Name "Microsoft.Graph"

    Write-Log -LogFile $logFile -Message "Connecting to Microsoft Graph (read-only)"
    Connect-GraphReadOnly -Tenant $TenantId

    $ctx = Get-MgContext
    Write-Log -LogFile $logFile -Message "Connected as $($ctx.Account) to tenant $($ctx.TenantId)"

    Write-Log -LogFile $logFile -Message "Retrieving directory roles"
    $roles = Get-MgDirectoryRole -All

    if (-not $roles) {
        Write-Log -LogFile $logFile -Level "WARN" -Message "No directory roles returned"
        $roles = @()
    }

    $results = New-Object System.Collections.Generic.List[object]
    $errors  = New-Object System.Collections.Generic.List[object]

    foreach ($role in $roles) {
        Write-Log -LogFile $logFile -Message "Role: $($role.DisplayName)"

        try {
            $members = Get-MgDirectoryRoleMember -DirectoryRoleId $role.Id -All

            foreach ($m in $members) {
                $results.Add([pscustomobject]@{
                    RoleName          = $role.DisplayName
                    RoleId            = $role.Id
                    MemberId          = $m.Id
                    MemberType        = $m.'@odata.type'
                    MemberDisplayName = Get-AdditionalProp $m "displayName"
                    MemberUPN         = Get-AdditionalProp $m "userPrincipalName"
                    MemberMail        = Get-AdditionalProp $m "mail"
                })
            }

            Write-Log -LogFile $logFile -Message "  Members found: $($members.Count)"
        }
        catch {
            Write-Log -LogFile $logFile -Level "ERROR" -Message "Failed on role $($role.DisplayName): $($_.Exception.Message)"
            $errors.Add([pscustomobject]@{
                RoleName = $role.DisplayName
                RoleId   = $role.Id
                Error    = $_.Exception.Message
            })
        }
    }

    # -------- Exports --------

    $csvPath  = Join-Path $runFolder "role-assignments.csv"
    $jsonPath = Join-Path $runFolder "role-assignments.json"
    $sumPath  = Join-Path $runFolder "summary.txt"

    Write-Log -LogFile $logFile -Message "Exporting CSV"
    $results | Sort-Object RoleName, MemberDisplayName | Export-Csv -NoTypeInformation -Path $csvPath

    Write-Log -LogFile $logFile -Message "Exporting JSON"
    $results | ConvertTo-Json -Depth 6 | Set-Content -Path $jsonPath -Encoding UTF8

    $summary = @(
        "Entra Role Audit Summary",
        "Run folder: $runFolder",
        "Tenant ID:  $($ctx.TenantId)",
        "Account:    $($ctx.Account)",
        "Roles:      $($roles.Count)",
        "Rows:       $($results.Count)",
        "Errors:     $($errors.Count)"
    ) -join [Environment]::NewLine

    $summary | Set-Content -Path $sumPath -Encoding UTF8
    Write-Log -LogFile $logFile -Message "Audit complete"
}
finally {
    if (-not $NoDisconnect) {
        try {
            Disconnect-MgGraph | Out-Null
            Write-Log -LogFile $logFile -Message "Disconnected from Microsoft Graph"
        } catch {}
    }
}
