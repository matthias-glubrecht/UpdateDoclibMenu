<#
    Extension in einer Site Collection registrieren
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

$configuration = @"
{
    "configs": [
        {
            "Libraries":["Project Documents","Contract Documents","ERP Documents","Lead Documents","Opportunity Documents","Account Documents"],
            "AllowedCommands":{
                "Level0":["newFolderCommand"],
                "Level1":["newFolderCommand","uploadFile"],
                "Level2":["newFolderCommand","uploadFile","uploadFolder","NewDOCCustomerDocument"]
            }
        }
    ]
}
"@

try {
    Connect-PnPOnline -Url $SiteCollectionUrl -Interactive -ClientId $ClientId
}
catch {
    Write-Error "Verbindung zu '$SiteCollectionUrl' fehlgeschlagen: $_"
    exit 1
}

$existing = Get-PnPCustomAction -Scope Site | Where-Object { $_.ClientSideComponentId -eq "c85aa61b-e557-495a-9351-9c17c20ab9eb" }
if ($existing) {
    Write-Warning "UpdateMenu Extension ist bereits auf '$SiteCollectionUrl' registriert. Keine Aktion durchgefuehrt."
    exit 0
}

try {
    Add-PnPCustomAction -Name "UpdateMenu" `
        -Title "UpdateMenu" `
        -Location "ClientSideExtension.ListViewCommandSet.CommandBar" `
        -ClientSideComponentId "c85aa61b-e557-495a-9351-9c17c20ab9eb" `
        -RegistrationType List -RegistrationId 101 `
        -Scope Site `
        -ClientSideComponentProperties $configuration

    Write-Host "UpdateMenu Extension erfolgreich registriert auf '$SiteCollectionUrl'." -ForegroundColor Green
}
catch {
    Write-Error "Registrierung der Extension fehlgeschlagen: $_"
    exit 1
}