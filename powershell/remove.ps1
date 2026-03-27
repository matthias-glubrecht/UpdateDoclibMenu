<#
    Extension aus einer Site Collection entfernen
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

try {
    $actions = Get-PnPCustomAction -Scope Site | Where-Object { $_.ClientSideComponentId -eq "c85aa61b-e557-495a-9351-9c17c20ab9eb" }

    if (-not $actions) {
        Write-Warning "Keine UpdateMenu Extension auf '$SiteCollectionUrl' gefunden."
        exit 0
    }

    $actions | Remove-PnPCustomAction -Force
    Write-Host "UpdateMenu Extension erfolgreich entfernt von '$SiteCollectionUrl'." -ForegroundColor Green
}
catch {
    Write-Error "Entfernung der Extension fehlgeschlagen: $_"
    exit 1
}
