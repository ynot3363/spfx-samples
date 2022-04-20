import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor,
} from "@microsoft/sp-listview-extensibility";

import * as strings from "NewsCommandSetStrings";
import { override } from "@microsoft/decorators";
import * as React from "react";
import * as ReactDom from "react-dom";
import { DialogWrapper, IDialogWrapperProps } from "./components/DialogWrapper";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface INewsCommandSetProperties {}

const LOG_SOURCE: string = "NewsCommandSet";

export default class NewsCommandSet extends BaseListViewCommandSet<INewsCommandSetProperties> {
  private dialogPlaceholder: HTMLDivElement = null;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized NewsCommandSet");

    this.dialogPlaceholder = document.body.appendChild(
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

    if (compareOneCommand) {
      let isVisible = false;

      if (event.selectedRows.length === 1) {
        isVisible =
          event.selectedRows[0].getValueByName("PromotedState") === "2";
      }

      compareOneCommand.visible = isVisible;
    }

    if (compareTwoCommand) {
      let isVisible = false;

      if (event.selectedRows.length === 1) {
        isVisible =
          event.selectedRows[0].getValueByName("PromotedState") === "0";
      }

      compareTwoCommand.visible = isVisible;
    }

    if (compareThreeCommand) {
      let isVisible = false;

      if (event.selectedRows.length === 1) {
        isVisible =
          event.selectedRows[0].getValueByName("PromotedState") === "2";
      }

      compareThreeCommand.visible = isVisible;
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "DEMOTE_NEWS_POST":
        /**
         * Loop through the selectedRows, there should only ever be 1 selected item though
         * because the command is only visible when 1 item is select and it is promoted
         */
        event.selectedRows.forEach(async (item) => {
          try {
            const checkedOutUser = item.getValueByName("CheckoutUser");
            const checkedOutUserId = item.getValueByName("CheckedOutUserId");
            const currentUserId =
              this.context.pageContext.legacyPageContext.userId.toString();
            const promotedState = item.getValueByName("PromotedState");

            if (promotedState === "2") {
              /**
               * Ensure that the page is not currently checked out to anyone before attempting to demote it.
               */
              if (checkedOutUserId && checkedOutUserId !== currentUserId) {
                this.renderDialogWrapper(
                  "Page is currently checked out",
                  `The page is currently checked out to ${checkedOutUser[0].title}. You need to take ownership of the page before you can demote it.`,
                  true
                );
                return;
              }

              const demoteNewsResponse = await this.demoteNewsPost(item);

              if (demoteNewsResponse.ok) {
                const data = await demoteNewsResponse.json();
                let hasException = false;
                let errorMessage = "";

                /**
                 * Loop through the fields in the response and ensure there are no errors
                 * if there are set hasException to true and errorMessage
                 */
                data.value.forEach((field: any) => {
                  if (field.HasException) {
                    hasException = true;
                    errorMessage = field.ErrorMessage;
                  }
                });

                if (hasException) {
                  this.renderDialogWrapper(
                    "Error demoting page",
                    `There was an error demoting the new post: ${errorMessage}`,
                    true
                  );
                  return;
                } else {
                  const publishPageResponse = await this.publishPage(item);

                  if (publishPageResponse.ok) {
                    this.renderDialogWrapper(
                      "Page demoted successfully",
                      `The news post, ${item.getValueByName(
                        "FileLeafRef"
                      )}, has been successfully demoted.`,
                      true
                    );
                    return;
                  } else {
                    this.renderDialogWrapper(
                      "Error publishing page",
                      `The news post, ${item.getValueByName(
                        "FileLeafRef"
                      )}, has been demoted but there was an error publishing the page. Please re-publish the page.`,
                      true
                    );
                    return;
                  }
                }
              } else {
                this.renderDialogWrapper(
                  "Error demoting page",
                  `There was an error trying to demote the news post, ${demoteNewsResponse.statusText}.`,
                  true
                );
                return;
              }
            }
          } catch (error) {
            Log.error(LOG_SOURCE, error);
            this.renderDialogWrapper("Error", `${error}`, true);
            return;
          }
        });

        break;
      case "PROMOTE_NEWS_POST":
        /**
         * Loop through the selectedRows, there should only ever be 1 selected item though
         * because the command is only visible when 1 item is select and it is promoted
         */
        event.selectedRows.forEach(async (item) => {
          try {
            const checkedOutUser = item.getValueByName("CheckoutUser");
            const checkedOutUserId = item.getValueByName("CheckedOutUserId");
            const currentUserId =
              this.context.pageContext.legacyPageContext.userId.toString();
            const promotedState = item.getValueByName("PromotedState");

            if (promotedState === "0") {
              /**
               * Ensure that the page is not currently checked out to anyone before attempting to promoting it.
               */
              if (checkedOutUserId && checkedOutUserId !== currentUserId) {
                this.renderDialogWrapper(
                  "Page is currently checked out",
                  `The page is currently checked out to ${checkedOutUser[0].title}. You need to take ownership of the page before you can promote it.`,
                  true
                );
                return;
              }

              const promoteNewsResponse = await this.promoteNewsPost(item);

              if (promoteNewsResponse.ok) {
                const data = await promoteNewsResponse.json();
                let hasException = false;
                let errorMessage = "";

                /**
                 * Loop through the fields in the response and ensure there are no errors
                 * if there are set hasException to true and errorMessage
                 */
                data.value.forEach((field: any) => {
                  if (field.HasException) {
                    hasException = true;
                    errorMessage = field.ErrorMessage;
                  }
                });

                if (hasException) {
                  this.renderDialogWrapper(
                    "Error promoting page",
                    `There was an error promoting the page to anew post: ${errorMessage}`,
                    true
                  );
                  return;
                } else {
                  const publishPageResponse = await this.publishPage(item);

                  if (publishPageResponse.ok) {
                    this.renderDialogWrapper(
                      "Page promoted successfully",
                      `The page post, ${item.getValueByName(
                        "FileLeafRef"
                      )}, has been successfully promoted to a news post.`,
                      true
                    );
                    return;
                  } else {
                    this.renderDialogWrapper(
                      "Error publishing page",
                      `The page, ${item.getValueByName(
                        "FileLeafRef"
                      )}, has been promoted to a news post but there was an error publishing the page. Please re-publish the page.`,
                      true
                    );
                    return;
                  }
                }
              } else {
                this.renderDialogWrapper(
                  "Error promoting page",
                  `There was an error trying to demote the news post, ${promoteNewsResponse.statusText}.`,
                  true
                );
                return;
              }
            }
          } catch (error) {
            Log.error(LOG_SOURCE, error);
            this.renderDialogWrapper("Error", `${error}`, true);
            return;
          }
        });

        break;
      case "UPDATE_PUBLISHING_DATE":
        break;
      default:
        throw new Error("Unknown command");
    }
  }

  /**
   * Method will update the page and set the promoted state to 0
   * @param item The item in the row that is currently being looped through
   * @returns response from the ValidateUpdateListItem() call
   */
  private async demoteNewsPost(
    item: RowAccessor
  ): Promise<SPHttpClientResponse> {
    const url = `${
      this.context.pageContext.site.absoluteUrl
    }/_api/web/lists(guid'${
      this.context.pageContext.list.id
    }')/items(${item.getValueByName("ID")})/ValidateUpdateListItem()`;
    const body = {
      formValues: [
        {
          FieldName: "PromotedState",
          FieldValue: "0",
        },
      ],
      bNewDocumentUpdate: false,
    };

    return await this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "Content-Type": "application/json;odata=nometadata",
          "Content-Length": JSON.stringify(body).length.toString(),
        },
        body: JSON.stringify(body),
      }
    );
  }

  /**
   * Method will update the page and set the promoted state to 2
   * @param item The item in the row that is currently being looped through
   * @returns response from the ValidateUpdateListItem() call
   */
  private async promoteNewsPost(
    item: RowAccessor
  ): Promise<SPHttpClientResponse> {
    const url = `${
      this.context.pageContext.site.absoluteUrl
    }/_api/web/lists(guid'${
      this.context.pageContext.list.id
    }')/items(${item.getValueByName("ID")})/ValidateUpdateListItem()`;
    const body = {
      formValues: [
        {
          FieldName: "PromotedState",
          FieldValue: "2",
        },
      ],
      bNewDocumentUpdate: false,
    };

    return await this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "Content-Type": "application/json;odata=nometadata",
          "Content-Length": JSON.stringify(body).length.toString(),
        },
        body: JSON.stringify(body),
      }
    );
  }

  /**
   * Method will publish a page based on the server relative url of the item
   * @param item The item in the row that is currently being looped through
   * @returns response from the Publish() call
   */
  private async publishPage(item: RowAccessor): Promise<SPHttpClientResponse> {
    const url = `${
      this.context.pageContext.site.absoluteUrl
    }/_api/web/GetFileByServerRelativeUrl('${item.getValueByName(
      "FileRef"
    )}')/Publish()`;

    const body = { comment: "Demoted News Post" };

    return await this.context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "Content-Type": "application/json;odata=nometadata",
        },
        body: JSON.stringify(body),
      }
    );
  }

  private closeErrorDialog = () => {
    this.renderDialogWrapper("", "", false);
  };

  private renderDialogWrapper(title: string, subText: string, show: boolean) {
    const element: React.ReactElement<IDialogWrapperProps> =
      React.createElement(DialogWrapper, {
        title: title,
        subText: subText,
        showDialog: show,
        closeDialog: this.closeErrorDialog,
      });

    ReactDom.render(element, this.dialogPlaceholder);
  }
}
