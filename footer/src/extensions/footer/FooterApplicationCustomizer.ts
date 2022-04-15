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
   * Looks for a specific element with class Canvas in the SharePoint page DOM to append our
   * customFooter div that we will render our Footer into.
   */
  private _container: Element = document.querySelectorAll(".Canvas")[0];
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
     * Only make the call to SP if the container we can append to exists
     */
    if (this._container) {
      this._footerItems = await this.getItems();
      this.context.application.navigatedEvent.add(this, this.onNavigated);
    }

    return Promise.resolve();
  }

  /**
   * Event that is executed when the site navigates to a new container.
   */
  private async onNavigated(): Promise<void> {
    /**
     * Need to delay the call because of a defect in SharePoint where the navigation event fires before
     * page navigation
     */
    setTimeout(this.renderFooter.bind(this), 3000);
  }

  /**
   * Searches for the canvas location to insert the footer into the page.
   */
  private async renderFooter(): Promise<void> {
    const footerElement: Element = document.querySelector("#customFooter");

    if (!footerElement && this._footerItems.length > 0 && this._container) {
      const footer: HTMLElement = document.createElement("footer");
      const element: React.ReactElement<IFooterProps> = React.createElement(
        Footer,
        {
          footerLinks: this._footerItems,
        }
      );

      this._container.appendChild(footer);
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
