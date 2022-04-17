import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters,
  RowAccessor,
} from "@microsoft/sp-listview-extensibility";
import { Dialog } from "@microsoft/sp-dialog";

import * as strings from "NewsCommandSetStrings";
import { override } from "@microsoft/decorators";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface INewsCommandSetProperties {}

const LOG_SOURCE: string = "NewsCommandSet";

export default class NewsCommandSet extends BaseListViewCommandSet<INewsCommandSetProperties> {
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, "Initialized NewsCommandSet");
    return Promise.resolve();
  }

  public onListViewUpdated(
    event: IListViewCommandSetListViewUpdatedParameters
  ): void {
    const compareOneCommand: Command = this.tryGetCommand("DEMOTE_NEWS_POST");

    if (compareOneCommand) {
      let isVisible = false;

      if (event.selectedRows.length === 1) {
        isVisible =
          event.selectedRows[0].getValueByName("PromotedState") === "2";
      }

      compareOneCommand.visible = isVisible;
    }
  }

  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case "DEMOTE_NEWS_POST":
        const selectedItems = event.selectedRows;

        /**
         * Loop through the selectedItems, there should only ever be 1 selected item though
         * because the command is only visible when 1 item is select and it is promoted
         */
        selectedItems.forEach(async (item) => {
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
                Dialog.alert(
                  `The page is currently checked out to ${checkedOutUser[0].title}. You need to take ownership of the page before you can demote it.`
                );
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
                  Dialog.alert(
                    `There was an error demoting the new post: ${errorMessage}`
                  );
                } else {
                  const publishPageResponse = await this.publishPage(item);

                  if (publishPageResponse.ok) {
                    Dialog.alert(
                      `The news post, ${item.getValueByName(
                        "FileLeafRef"
                      )}, has been successfully demoted.`
                    );
                    location.reload();
                  } else {
                    Dialog.alert(
                      `There was an error trying to demote the news post, ${publishPageResponse.statusText}.`
                    );
                  }
                }
              } else {
                Dialog.alert(
                  `There was an error trying to demote the news post, ${demoteNewsResponse.statusText}.`
                );
              }
            }
          } catch (error) {
            Log.error(LOG_SOURCE, error);
            Dialog.alert(error);
          }
        });

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
}
