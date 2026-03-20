import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  type Command,
  type IListViewCommandSetExecuteEventParameters,
  type ListViewStateChangedEventArgs
} from '@microsoft/sp-listview-extensibility';

export interface IRedirectExtensionCommandSetProperties {
  redirectPageUrl: string;
}

const LOG_SOURCE: string = 'RedirectExtensionCommandSet';

export default class RedirectExtensionCommandSet extends BaseListViewCommandSet<IRedirectExtensionCommandSetProperties> {

  private _styleElement: HTMLStyleElement | null = null;

  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized RedirectExtensionCommandSet');
    console.log('RedirectExtension - Initialized. Properties:', this.properties);

    // Hide the OOB SharePoint "New" / "+ Add new item" button via CSS
    this._hideDefaultNewButton();

    const newCommand: Command = this.tryGetCommand('NEW_ITEM');
    if (newCommand) {
      newCommand.visible = true;
    }

    const editCommand: Command = this.tryGetCommand('EDIT_ITEM');
    if (editCommand) {
      editCommand.visible = false;
    }

    const viewCommand: Command = this.tryGetCommand('VIEW_ITEM');
    if (viewCommand) {
      viewCommand.visible = false;
    }

    this.context.listView.listViewStateChangedEvent.add(this, this._onListViewStateChanged);

    return Promise.resolve();
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    const redirectPageUrl: string = this.properties.redirectPageUrl;

    Log.info(LOG_SOURCE, `Command executed: ${event.itemId}`);
    console.log('RedirectExtension - Command executed:', event.itemId, 'Properties:', this.properties);

    if (!redirectPageUrl) {
      Log.error(LOG_SOURCE, new Error('redirectPageUrl property is not configured.'));
      return;
    }

    const listId: string = this.context.pageContext.list?.id?.toString() || '';
    const siteUrl: string = this.context.pageContext.web.absoluteUrl;

    switch (event.itemId) {
      case 'NEW_ITEM': {
        const newUrl = `${siteUrl}/${redirectPageUrl}?action=new&listId=${listId}&source=${encodeURIComponent(window.location.href)}`;
        Log.info(LOG_SOURCE, `Redirecting New to: ${newUrl}`);
        window.location.href = newUrl;
        break;
      }
      case 'EDIT_ITEM': {
        const selectedItemId = this.context.listView.selectedRows?.[0]?.getValueByName('ID');
        if (!selectedItemId) {
          Log.warn(LOG_SOURCE, 'No item selected for Edit.');
          return;
        }
        const editUrl = `${siteUrl}/${redirectPageUrl}?action=edit&itemId=${selectedItemId}&listId=${listId}&source=${encodeURIComponent(window.location.href)}`;
        Log.info(LOG_SOURCE, `Redirecting Edit to: ${editUrl}`);
        window.location.href = editUrl;
        break;
      }
      case 'VIEW_ITEM': {
        const viewItemId = this.context.listView.selectedRows?.[0]?.getValueByName('ID');
        if (!viewItemId) {
          Log.warn(LOG_SOURCE, 'No item selected for View.');
          return;
        }
        const viewUrl = `${siteUrl}/${redirectPageUrl}?action=view&itemId=${viewItemId}&listId=${listId}&source=${encodeURIComponent(window.location.href)}`;
        Log.info(LOG_SOURCE, `Redirecting View to: ${viewUrl}`);
        window.location.href = viewUrl;
        break;
      }
      default:
        throw new Error('Unknown command');
    }
  }

  private _hideDefaultNewButton(): void {
    this._styleElement = document.createElement('style');
    this._styleElement.textContent = `
      /* Hide OOB SharePoint "New" button and "+ Add new item" link */
      button[data-automationid="newCommand"],
      button[aria-label="New"],
      .ms-CommandBar-mainArea button[name="New"],
      .ms-addnew {
        display: none !important;
      }
    `;
    document.head.appendChild(this._styleElement);
  }

  protected onDispose(): void {
    if (this._styleElement) {
      this._styleElement.remove();
      this._styleElement = null;
    }
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const editCommand: Command = this.tryGetCommand('EDIT_ITEM');
    if (editCommand) {
      editCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    const viewCommand: Command = this.tryGetCommand('VIEW_ITEM');
    if (viewCommand) {
      viewCommand.visible = this.context.listView.selectedRows?.length === 1;
    }

    this.raiseOnChange();
  }
}
