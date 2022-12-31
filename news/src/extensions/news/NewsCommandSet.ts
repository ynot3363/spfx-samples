import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor,
} from "@microsoft/sp-listview-extensibility";

import { override } from "@microsoft/decorators";
import * as React from "react";
import * as ReactDom from "react-dom";
import {
  CustomPanel,
  ICustomPanelProps,
} from "../../components/CustomPanel/CustomPanel";
import { SPPermission } from "@microsoft/sp-page-context";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface INewsCommandSetProperties {}

const LOG_SOURCE: string = "NewsCommandSet";

export default class NewsCommandSet extends BaseListViewCommandSet<INewsCommandSetProperties> {
  private panelPlaceHolder: HTMLDivElement = null;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized NewsCommandSet");

    // Try to get the three commands
    const compareOneCommand: Command = this.tryGetCommand("DEMOTE_NEWS_POST");
    const compareTwoCommand: Command = this.tryGetCommand("PROMOTE_NEWS_POST");
    const compareThreeCommand: Command = this.tryGetCommand(
      "UPDATE_PUBLISHING_DATE"
    );

    // Ensure that the buttons are not visible by default
    if (compareOneCommand) {
      compareOneCommand.visible = false;
    }
    if (compareTwoCommand) {
      compareTwoCommand.visible = false;
    }
    if (compareThreeCommand) {
      compareThreeCommand.visible = false;
    }

    // Create the container for our React component
    this.panelPlaceHolder = document.body.appendChild(
      document.createElement("div")
    );

    return Promise.resolve();
  }

  public onListViewUpdated(
    event: IListViewCommandSetListViewUpdatedParameters
  ): void {
    const compareOneCommand: Command = this.tryGetCommand("DEMOTE_NEWS_POST");
    const compareTwoCommand: Command = this.tryGetCommand("PROMOTE_NEWS_POST");
    const compareThreeCommand: Command = this.tryGetCommand(
      "UPDATE_PUBLISHING_DATE"
    );

    // Check to see if the user has edit permissions to the list
    const userHasEditPermissions: boolean =
      this.context.pageContext.list.permissions.hasPermission(
        SPPermission.editListItems
      );

    const allItemsPromoted =
      event.selectedRows.every(
        (row) => row.getValueByName("PromotedState") === "2"
      ) && event.selectedRows.length > 0;
    const allItemsNotPromoted =
      event.selectedRows.every(
        (row) => row.getValueByName("PromotedState") === "0"
      ) && event.selectedRows.length > 0;

    if (userHasEditPermissions) {
      if (compareOneCommand) {
        let isVisible = false;

        if (allItemsPromoted) {
          isVisible = true;
        }

        compareOneCommand.visible = isVisible;
      }

      if (compareTwoCommand) {
        let isVisible = false;

        if (allItemsNotPromoted) {
          isVisible = true;
        }

        compareTwoCommand.visible = isVisible;
      }

      if (compareThreeCommand) {
        let isVisible = false;

        if (allItemsPromoted) {
          isVisible = true;
        }

        compareThreeCommand.visible = isVisible;
      }
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    const items: RowAccessor[] = [...event.selectedRows];

    switch (event.itemId) {
      case "DEMOTE_NEWS_POST":
        try {
          this._showPanel(
            items,
            "Demote News Posts",
            "Update the pages below to no longer be classified as News posts.",
            "DEMOTE_NEWS_POST"
          );
        } catch (error) {
          Log.error(LOG_SOURCE, error);
        }
        break;
      case "PROMOTE_NEWS_POST":
        try {
          this._showPanel(
            items,
            "Promote to News Posts",
            "Update the pages below to be classified as News posts.",
            "PROMOTE_NEWS_POST"
          );
        } catch (error) {
          Log.error(LOG_SOURCE, error);
        }
        break;
      case "UPDATE_PUBLISHING_DATE":
        try {
          this._showPanel(
            items,
            "Update Publishing Date",
            "Update the publishing date for the selected pages below.",
            "UPDATE_PUBLISHING_DATE"
          );
        } catch (error) {
          Log.error(LOG_SOURCE, error);
        }
        break;
      default:
        throw new Error("Unknown command");
    }
  }

  private _showPanel(
    items: RowAccessor[],
    title: string,
    descripton: string,
    command: string
  ) {
    this._renderPanelComponent({
      context: this.context,
      items: items,
      isOpen: true,
      panelTitle: title,
      panelDescription: descripton,
      command: command,
      onRefresh: () => location.reload(),
      onClose: () => this._renderPanelComponent({ isOpen: false }),
    });
  }

  private _renderPanelComponent(props: any) {
    const element: React.ReactElement<ICustomPanelProps> = React.createElement(
      CustomPanel,
      {
        isOpen: false,
        items: [],
        panelTitle: null,
        panelDescription: null,
        onClose: null,
        onRefresh: null,
        ...props,
      }
    );
    ReactDom.render(element, this.panelPlaceHolder);
  }
}
