import * as React from "react";
import * as ReactDom from "react-dom";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
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
  private _topPlaceholder: PlaceholderContent | undefined;

  /**
   * Overrides the onInit function and searchs for the canvas location to insert the
   * footer into the page.
   */
  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    if (!this.properties.intranetUrl || !this.properties.messageListId) {
      const error: Error = new Error(
        "Missing required configuration properties."
      );
      Log.error(LOG_SOURCE, error);
    }

    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );
    }

    this._messages = await this._getItems();

    this.context.application.navigatedEvent.add(this, this._renderMessages);

    this._renderMessages();

    return Promise.resolve();
  }

  private _onDispose(placeholderContent: PlaceholderContent): void {
    ReactDom.unmountComponentAtNode(placeholderContent.domElement);
  }

  /**
   * Searches for the canvas location to insert the footer into the page.
   */
  private _renderMessages(): void {
    const element: React.ReactElement<IMessageProps> = React.createElement(
      Message,
      {
        messages: this._messages,
      }
    );

    ReactDom.render(element, this._topPlaceholder.domElement);
  }

  /**
   * Makes a REST call to the associated list that holds the footer links
   * @returns A collection of IMessage
   */
  private async _getItems(): Promise<IMessage[]> {
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
