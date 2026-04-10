import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';

const LOG_SOURCE: string = 'UpdateMenuCommandSet';
const STYLE_ELEMENT_ID: string = 'b3f7e2a1-8c4d-4f9e-a6d5-1e0b3c7f2d84';

export interface IAllowedCommands {
  [level: string]: string[];
}

export interface ILibraryConfig {
  Libraries: string[];
  AllowedCommands: IAllowedCommands;
}

export interface IUpdateMenuCommandSetProperties {
  configs: ILibraryConfig[];
}

export default class UpdateMenuCommandSet extends BaseListViewCommandSet<IUpdateMenuCommandSetProperties> {

  private _lastUrl: string = '';
  private _menuObserver: MutationObserver | undefined;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized UpdateMenuCommandSet');

    this._lastUrl = window.location.href;
    this.updateStyles();
    this.startMenuObserver();

    // SharePoint navigates within doc libs using history.pushState/replaceState,
    // which does NOT trigger listViewStateChangedEvent. We intercept those calls.
    const originalPushState = history.pushState.bind(history);
    const originalReplaceState = history.replaceState.bind(history);

    history.pushState = (...args: Parameters<History['pushState']>): void => {
      originalPushState(...args);
      this.onUrlChanged();
    };

    history.replaceState = (...args: Parameters<History['replaceState']>): void => {
      originalReplaceState(...args);
      this.onUrlChanged();
    };

    // Handle browser back/forward buttons
    window.addEventListener('popstate', () => this.onUrlChanged());

    // Also update on list view state changes
    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    // Hide the dummy command
    const hiddenCommand = this.tryGetCommand('HIDDEN_COMMAND');
    if (hiddenCommand) {
      hiddenCommand.visible = false;
    }

    return Promise.resolve();
  }

  public onExecute(_event: IListViewCommandSetExecuteEventParameters): void {
    // No commands to execute – this extension only injects CSS
  }

  protected onDispose(): void {
    if (this._menuObserver) {
      this._menuObserver.disconnect();
      this._menuObserver = undefined;
    }
    this.removeStyles();
  }

  private _onListViewStateChanged = (_args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');
    this.updateStyles();
  };

  private onUrlChanged(): void {
    const currentUrl: string = window.location.href;
    if (currentUrl !== this._lastUrl) {
      this._lastUrl = currentUrl;
      Log.info(LOG_SOURCE, `URL changed: ${currentUrl}`);
      this.updateStyles();
    }
  }

  private getFolderDepth(): number {
    const list = this.context.pageContext.list;
    if (!list) {
      return 0;
    }
    const listRootPath: string = list.serverRelativeUrl.replace(/\/+$/, '');

    const url: URL = new URL(window.location.href);
    const folderPath: string = (url.searchParams.get('id') || url.searchParams.get('RootFolder') || '')
      .replace(/\/+$/, '');

    if (!folderPath || folderPath === listRootPath) {
      return 0;
    }

    const relativePath: string = folderPath.indexOf(listRootPath) === 0
      ? folderPath.substring(listRootPath.length)
      : folderPath;
    const segments: string[] = relativePath.split('/').filter(s => s.length > 0);
    return segments.length;
  }

  private findConfigForCurrentLibrary(): ILibraryConfig | undefined {
    const list = this.context.pageContext.list;
    if (!list) {
      return undefined;
    }
    const configs: ILibraryConfig[] = this.properties.configs || [];
    const matches: ILibraryConfig[] = configs.filter(c => (c.Libraries || []).indexOf(list.title) !== -1);
    return matches.length > 0 ? matches[0] : undefined;
  }

  private getAllowedCommands(config: ILibraryConfig, depth: number): string[] {
    const key: string = 'Level' + depth;
    const allowed: IAllowedCommands = config.AllowedCommands || {};
    return allowed[key] || [];
  }

  private updateStyles(): void {
    if (!this.context.pageContext.list) {
      this.removeStyles();
      return;
    }

    const config: ILibraryConfig | undefined = this.findConfigForCurrentLibrary();
    if (!config) {
      Log.info(LOG_SOURCE, 'No config found for library "' + this.context.pageContext.list.title + '"');
      this.removeStyles();
      return;
    }

    const depth: number = this.getFolderDepth();
    Log.info(LOG_SOURCE, 'Library: "' + this.context.pageContext.list.title + '", Folder depth: ' + depth);

    const allowedCommands: string[] = this.getAllowedCommands(config, depth);

    // Hide all menu items with a data-automationid...
    const rules: string[] = [
      'div#command-bar-menu-id div[class*="contextMenu_"] li.ms-ContextualMenu-item:has(> button[data-automationid]) { display: none !important; }'
    ];

    // ...then show back the whitelisted ones
    for (let i: number = 0; i < allowedCommands.length; i++) {
      rules.push('div#command-bar-menu-id div[class*="contextMenu_"] li.ms-ContextualMenu-item:has(> button[data-automationid="' + allowedCommands[i] + '"]) { display: list-item !important; }');
    }

    const css: string = rules.join('\n');

    let style: HTMLStyleElement | null = document.getElementById(STYLE_ELEMENT_ID) as HTMLStyleElement | null;
    if (!style) {
      style = document.createElement('style');
      style.id = STYLE_ELEMENT_ID;
      document.head.appendChild(style);
    }
    if (style.innerHTML !== css) {
      style.innerHTML = css;
      Log.info(LOG_SOURCE, 'Updated styles');
    }
  }

  private removeStyles(): void {
    const style: HTMLElement | null = document.getElementById(STYLE_ELEMENT_ID);
    if (style) {
      style.innerHTML = '';
    }
  }

  /** Observes the DOM for context menus and cleans up orphaned dividers. */
  private startMenuObserver(): void {
    this._menuObserver = new MutationObserver(() => this.cleanupDividers());
    this._menuObserver.observe(document.body, { childList: true, subtree: true });
  }

  /** Hides dividers that are adjacent to another divider or at the start/end of the menu. */
  private cleanupDividers(): void {
    const menus: NodeListOf<Element> = document.querySelectorAll('ul.ms-ContextualMenu-list');
    for (let m: number = 0; m < menus.length; m++) {
      const items: HTMLCollection = menus[m].children;
      let lastVisibleWasDivider: boolean = true; // treat start of list as divider
      for (let i: number = 0; i < items.length; i++) {
        const item: HTMLElement = items[i] as HTMLElement;
        const isDivider: boolean = item.getAttribute('role') === 'separator';
        const isHidden: boolean = window.getComputedStyle(item).display === 'none';

        if (isHidden) {
          continue;
        }

        if (isDivider) {
          if (lastVisibleWasDivider) {
            item.style.display = 'none';
          } else {
            lastVisibleWasDivider = true;
          }
        } else {
          lastVisibleWasDivider = false;
        }
      }

      // Hide trailing divider
      for (let i: number = items.length - 1; i >= 0; i--) {
        const item: HTMLElement = items[i] as HTMLElement;
        const isHidden: boolean = item.style.display === 'none' || window.getComputedStyle(item).display === 'none';
        if (isHidden) {
          continue;
        }
        if (item.getAttribute('role') === 'separator') {
          item.style.display = 'none';
        }
        break;
      }
    }
  }
}
