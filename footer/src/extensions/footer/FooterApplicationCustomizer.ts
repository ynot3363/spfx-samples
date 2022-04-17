import { ISPItem } from "./../../models/ISPItem";
import * as React from "react";
import * as ReactDom from "react-dom";
import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import * as strings from "FooterApplicationCustomizerStrings";
import { override } from "@microsoft/decorators";
import { IFooterLink } from "../../models/IFooterLink";
import Footer, { IFooterProps } from "./components/Footer";

const LOG_SOURCE: string = "FooterApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFooterApplicationCustomizerProperties {
  // The link to the URL where the footer list is on
  intranetUrl: string;
  // The GUID for the list that holds the footer links
  footerListId: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class FooterApplicationCustomizer extends BaseApplicationCustomizer<IFooterApplicationCustomizerProperties> {
  /**
   * Holds a collection of footer links from the associated list
   */
  private _footerItems: IFooterLink[] = [];

  /**
   * Overrides the onInit function and searchs for the canvas location to insert the
   * footer into the page.
   */
  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    /**
     * Looks for a specific element with id spPageCanvasContent in the SharePoint page DOM to append our
     * customFooter div that we will render our Footer into.
     */
    const container: Element = document.querySelector(
      `div[data-automation-id="mainScrollRegionInnerContent"]`
    ).parentElement;

    if (container) {
      this._footerItems = await this.getItems();
    }

    this.context.application.navigatedEvent.add(this, this.renderFooter);

    return Promise.resolve();
  }

  /**
   * Searches for the canvas location to insert the footer into the page.
   */
  private renderFooter(): void {
    const footerElement: Element = document.querySelector("#customFooter");
    /**
     * You need to get the container on the page in the renderFooter method because on partial
     * page loads this is the only function that runs and if you get it once you will have a
     * reference to an old HTML element
     */
    const container = document.querySelector(
      `div[data-automation-id="mainScrollRegionInnerContent"]`
    ).parentElement;

    if (!footerElement && container && this._footerItems.length > 0) {
      const footer: HTMLElement = document.createElement("footer");
      const element: React.ReactElement<IFooterProps> = React.createElement(
        Footer,
        {
          footerLinks: this._footerItems,
        }
      );

      container.appendChild(footer);
      ReactDom.render(element, footer);
    }
  }

  /**
   * Makes a REST call to the associated list that holds the footer links
   * @returns A collection of IFooterLink
   */
  private async getItems(): Promise<IFooterLink[]> {
    const url: string = `${this.properties.intranetUrl}/_api/web/lists(guid'${this.properties.footerListId}')/items?$orderBy=linkOrder`;

    return new Promise<IFooterLink[]>((resolve, reject) => {
      const now = new Date();
      const twentyMinutesLater = now.setMinutes(now.getMinutes() + 20);
      /**
       * Retrieve footerLinks from localStorage if they exist
       */
      const checkLocalStoreItem = localStorage.getItem("footerLinks");

      if (checkLocalStoreItem) {
        const footerLinksLocalStore = JSON.parse(checkLocalStoreItem);

        /**
         * Ensure the local storage item is not older than 20 minutes
         */
        if (now.getTime() < footerLinksLocalStore.expires) {
          resolve(footerLinksLocalStore.value);
        }
      }

      this.context.spHttpClient
        .get(url, SPHttpClient.configurations.v1)
        .then(async (response: SPHttpClientResponse) => {
          if (response.ok) {
            const spItems: ISPItem[] = (await response.json()).value;

            /**
             * Map through the SharePoint list data and map it to our IFooterLink data model
             */
            const footerItems: IFooterLink[] = spItems.map((item) => {
              return {
                name: item.Title || "",
                link: item.link
                  ? { url: item.link.Url, desc: item.link.Description }
                  : null,
                icon: item.icon
                  ? { url: item.icon.Url, desc: item.icon.Description }
                  : null,
                order: item.linkOrder || null,
                id: item.ID,
              };
            });

            /**
             * Create an item to store in local storage and add an expiration time of 20 minutes
             */
            const localStoreItem = {
              value: footerItems,
              expires: twentyMinutesLater,
            };

            localStorage.setItem("footerLinks", JSON.stringify(localStoreItem));

            resolve(footerItems);
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
