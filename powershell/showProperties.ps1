<#
    ClientSideComponentProperties der UpdateMenu Extension anzeigen
#>

[CmdletBinding()]
param (
    [Parameter(mandatory)]
    [string]
    $SiteCollectionUrl,

    [Parameter(mandatory)]
    [string]
    $ClientId
)

try {
    Connect-PnPOnline -Url $SiteCollectionUrl -Interactive -ClientId $ClientId
}
catch {
    Write-Error "Verbindung zu '$SiteCollectionUrl' fehlgeschlagen: $_"
    exit 1
}

$action = Get-PnPCustomAction -Scope Site | Where-Object { $_.ClientSideComponentId -eq "c85aa61b-e557-495a-9351-9c17c20ab9eb" }

if (-not $action) {
    Write-Warning "Keine UpdateMenu Extension auf '$SiteCollectionUrl' gefunden."
    exit 0
}

Write-Host "UpdateMenu Extension gefunden auf '$SiteCollectionUrl':" -ForegroundColor Green
$action.ClientSideComponentProperties | ConvertFrom-Json | ConvertTo-Json -Depth 10
