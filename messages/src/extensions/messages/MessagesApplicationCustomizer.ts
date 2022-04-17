import * as React from "react";
import * as ReactDom from "react-dom";
import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import * as strings from "MessagesApplicationCustomizerStrings";
import { override } from "@microsoft/decorators";
import { ISPItem } from "../../models/ISPItem";
import { IMessage } from "../../models/IMessage";
import Message, { IMessageProps } from "./components/Message/Message";

const LOG_SOURCE: string = "MessagesApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IMessagesApplicationCustomizerProperties {
  // The link to the URL where the messages list is on
  intranetUrl: string;
  // The GUID for the list that holds the messages
  messageListId: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class MessagesApplicationCustomizer extends BaseApplicationCustomizer<IMessagesApplicationCustomizerProperties> {
  /**
   * Holds an internal collection of messages and will be populated with items from the associated
   * SharePoint list.
   */
  private _messages: IMessage[] = [];

  /**
   * Overrides the onInit function and searchs for the canvas location to insert the
   * footer into the page.
   */
  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    /**
     * Looks for a specific element with id spPageCanvasContent in the SharePoint page DOM to prepend our
     * messageContainer div that we will render our Messages into.
     */
    const container: Element = document.getElementById("spPageCanvasContent")
      .parentElement.parentElement;

    if (container) {
      this._messages = await this.getItems();
    }

    this.context.application.navigatedEvent.add(this, this.renderMessages);

    return Promise.resolve();
  }

  /**
   * Searches for the canvas location to insert the footer into the page.
   */
  private renderMessages(): void {
    const messageContainerElement: Element =
      document.getElementById("messageContainer");
    /**
     * You need to get the container on the page in the renderMessages method because on partial
     * page loads this is the only function that runs and if you get it once you will have a
     * reference to an old HTML element
     */
    const container = document.getElementById("spPageCanvasContent")
      .parentElement.parentElement;

    if (!messageContainerElement && container && this._messages.length > 0) {
      const messageContainer: HTMLElement = document.createElement("div");
      const element: React.ReactElement<IMessageProps> = React.createElement(
        Message,
        {
          messages: this._messages,
        }
      );

      container.prepend(messageContainer);
      ReactDom.render(element, messageContainer);
    }
  }

  /**
   * Makes a REST call to the associated list that holds the footer links
   * @returns A collection of IMessage
   */
  private async getItems(): Promise<IMessage[]> {
    const now = new Date();
    const url: string = `${this.properties.intranetUrl}/_api/web/lists(guid'${
      this.properties.messageListId
    }')/items?$filter=msg_publishDate le datetime'${now.toISOString()}' and msg_expirationDate gt datetime'${now.toISOString()}'&$orderBy=msg_publishDate`;

    return new Promise<IMessage[]>((resolve, reject) => {
      this.context.spHttpClient
        .get(url, SPHttpClient.configurations.v1)
        .then(async (response: SPHttpClientResponse) => {
          if (response.ok) {
            const spItems: ISPItem[] = (await response.json()).value;

            /**
             * Map through the SharePoint list data and map it to our IMessage data model
             */
            const messageItems: IMessage[] = spItems.map((item) => {
              return {
                message: item.Title || "",
                details: item.msg_details || "",
                link: item.msg_link
                  ? { url: item.msg_link.Url, desc: item.msg_link.Description }
                  : null,
                type: item.msg_type || "",
                publishDate: item.msg_publishDate,
                expirationDate: item.msg_expirationDate,
                id: item.ID,
              };
            });

            resolve(messageItems);
          } else {
            reject(response.statusText);
          }
        })
        .catch((error: Error) => {
          reject(error.message);
        });
    });
  }
}
