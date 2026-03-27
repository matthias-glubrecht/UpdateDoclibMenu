# UpdateDoclibMenu

SPFx ListViewCommandSet extension that controls which menu items are visible in the "New" menu of SharePoint document libraries, based on the library name and the current folder depth.

## Authors

- Matthias Glubrecht
- GitHub Copilot

## How it works

The extension injects CSS rules to hide menu items (identified by their `data-automationid`) that are not explicitly allowed for the current folder level. It intercepts `history.pushState` / `replaceState` and listens for `popstate` events to react to in-library navigation without full page reloads.

## Configuration

The extension is configured via `ClientSideComponentProperties` when registering the Custom Action. The configuration uses a pure whitelist approach — all menu items with a `data-automationid` are hidden by default, and only explicitly listed items are shown.

### `configs`

An array of library-specific configurations:

| Property | Type | Description |
|---|---|---|
| `Libraries` | `string[]` | Library display names this config applies to |
| `AllowedCommands` | `object` | An object with `Level0`, `Level1`, `Level2`, ... keys |

Each `LevelN` key contains a `string[]` of `data-automationid` values that are **allowed** (visible) at that folder depth. All other menu items are hidden. If a level is not specified, all menu items are hidden at that depth.

Orphaned menu dividers (separators between hidden groups) are automatically cleaned up via a DOM observer.

### Known `data-automationid` values

| automationid | Menu item |
|---|---|
| `newFolderCommand` | Folder |
| `uploadFile` | Files upload |
| `uploadFolder` | Folder upload |
| `newWordDocument` | Word document |
| `newExcelWorkbook` | Excel workbook |
| `newPowerPointPresentation` | PowerPoint presentation |
| `newOneNoteNotebook` | OneNote notebook |
| `newVisioDrawing` | Visio drawing |
| `CreateClipChampCommand` | Clipchamp video |
| `NewDOCCustomerDocument` | CustomerDocument (custom template) |
| `CreateShortcutCommand` | Link |
| `editNewMenuCommand` | Edit New menu |
| `addTemplate` | Add Template |

Custom document templates use the pattern `NewDOC` followed by the template name (e.g. `NewDOCCustomerDocument`). The exact IDs depend on which templates are configured in the library.

### Example configuration

```json
{
  "configs": [
    {
      "Libraries": ["Project Documents", "Contract Documents"],
      "AllowedCommands": {
        "Level0": ["newFolderCommand"],
        "Level1": ["newFolderCommand", "uploadFile"],
        "Level2": ["newFolderCommand", "uploadFile", "uploadFolder", "NewDOCCustomerDocument"]
      }
    }
  ]
}
```

In this example:

- **Level 0** (library root): only "New Folder" is visible
- **Level 1** (first subfolder level): "New Folder" and "Files upload"
- **Level 2+** (deeper): "New Folder", "Files upload", "Folder upload", and the custom document template
- Libraries not listed in any config are not affected
- Any new menu items Microsoft adds in the future will be automatically hidden

## Compatibility

| Requirement | Version |
|---|---|
| SharePoint | SharePoint Online |
| SPFx | 1.21.1 |
| Node.js | >=22.14.0 <23.0.0 |
| TypeScript | ~5.3.3 |

## Deployment

### Prerequisites

- [Node.js](https://nodejs.org/) v22.14.0+
- [PnP PowerShell](https://pnp.github.io/powershell/)

### Build

```bash
npm install
gulp bundle --ship
gulp package-solution --ship
```

### Install

1. Upload `sharepoint/solution/update-doclib-menu.sppkg` to the tenant App Catalog
2. Deploy (approve) the app — `skipFeatureDeployment` is enabled, so no site-level install is needed

### Register on a Site Collection

```powershell
.\powershell\deploy.ps1 -SiteCollectionUrl "https://contoso.sharepoint.com/sites/yoursite"
```

This registers the extension as a Custom Action with `-Scope Site`, so it applies to all document libraries (Template ID 101) across the entire site collection.

### Unregister

```powershell
.\powershell\remove.ps1 -SiteCollectionUrl "https://contoso.sharepoint.com/sites/yoursite"
```

## Version

1.0.0
