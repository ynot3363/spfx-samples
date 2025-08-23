import * as React from "react";
import * as ReactDom from "react-dom";
import * as strings from "FooterApplicationCustomizerStrings";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import { Log } from "@microsoft/sp-core-library";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { override } from "@microsoft/decorators";
import { IFooterLink } from "./models/IFooterLink";
import { ISPItem } from "./models/ISPItem";
import Footer, { IFooterProps } from "./components/Footer";

const LOG_SOURCE: string = "FooterApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IFooterApplicationCustomizerProperties {
  // The link to the URL where the footer list is on
  siteUrl: string;
  // The GUID for the list that holds the footer links
  listGuid: string;
  // The ID of the footer element
  footerElementId: string;
  // The cacheKey to use for storing footer links in local storage
  cacheKey: string;
  // Whether to show a copyright message in the footer, defaults to false
  showCopyright?: boolean;
  // Company name
  companyName?: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class FooterApplicationCustomizer extends BaseApplicationCustomizer<IFooterApplicationCustomizerProperties> {
  private _footerLinks: IFooterLink[] = [];
  private _bottomPlaceholder: PlaceholderContent | undefined;

  @override
  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    if (
      !this.properties.siteUrl ||
      !this.properties.listGuid ||
      !this.properties.cacheKey ||
      !this.properties.footerElementId
    ) {
      const error: Error = new Error(
        "Missing required configuration properties."
      );
      Log.error(LOG_SOURCE, error);
      return Promise.reject(error);
    }

    if (this.properties.showCopyright === undefined) {
      this.properties.showCopyright = false;
    }

    if (!this.properties.companyName) {
      this.properties.companyName = "";
    }

    /**
     * Even though we are not using the bottom placeholder, we need to create it
     * otherwise the extension will fail to load on the page.
     */
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Bottom,
          { onDispose: this._onDispose }
        );
    }

    try {
      this._footerLinks = await this._getItems();
    } catch (error) {
      Log.error(LOG_SOURCE, error);
    }

    this._renderFooter();

    this.context.application.navigatedEvent.add(this, this._renderFooter);

    return Promise.resolve();
  }

  private _renderFooter(): void {
    const { companyName, footerElementId, showCopyright } = this.properties;

    /**
     * Helper function to render the footer so we can use the observer pattern to
     * load the footer onto the page when the adjacent elements are available.
     * @returns A boolean indicating whether the footer was successfully rendered
     */
    const tryRenderFooter = (): boolean => {
      if (document.getElementById(footerElementId)) {
        return true;
      }
      try {
        const mainContent = document.querySelector(".mainContent");
        if (!mainContent) return false;

        const container = mainContent.querySelector(
          `div[data-automation-id="contentScrollRegion"]`
        );
        if (!container) return false;

        const footer: HTMLElement = document.createElement("footer");
        footer.id = footerElementId;

        const element: React.ReactElement<IFooterProps> = React.createElement(
          Footer,
          {
            footerLinks: this._footerLinks,
            showCopyright: showCopyright,
            companyName: companyName,
          }
        );

        // Append the footer to the first child of the container, if available.
        if (container.children.length > 0) {
          container.children[0].appendChild(footer);
        } else {
          container.appendChild(footer);
        }

        ReactDom.render(element, footer);

        return true;
      } catch (error) {
        console.error(error);
        return false;
      }
    };

    // Attempt to render immediately.
    if (tryRenderFooter()) {
      return;
    }

    // If not rendered yet, set up a MutationObserver to wait for the container.
    const observer = new MutationObserver(() => {
      if (tryRenderFooter()) {
        observer.disconnect();
      }
    });

    observer.observe(document.body, {
      childList: true,
      subtree: true,
    });
  }

  private _onDispose(placeholderContent: PlaceholderContent): void {
    const footer = document.getElementById(this.properties.footerElementId);
    if (footer) {
      ReactDom.unmountComponentAtNode(footer);
    }
    ReactDom.unmountComponentAtNode(placeholderContent.domElement);
  }

  /**
   * Makes a REST call to the associated list that holds the footer links
   * @returns A collection of IFooterLink
   */
  private async _getItems(): Promise<IFooterLink[]> {
    const { siteUrl, listGuid, cacheKey } = this.properties;
    const url: string = `${siteUrl}/_api/web/lists(guid'${listGuid}')/items?$orderBy=linkOrder`;

    return new Promise<IFooterLink[]>((resolve, reject) => {
      const now = new Date();
      const expirationTime = now.setMinutes(now.getMinutes() + 20); // Arbitratry cache time of 20 minutes change as needed
      const cachedLinks = localStorage.getItem(cacheKey);

      if (cachedLinks) {
        try {
          const cachedLinksJSON = JSON.parse(cachedLinks);

          if (now.getTime() < cachedLinksJSON.expires) {
            resolve(cachedLinksJSON.value);
          }
        } catch (error) {
          console.error(
            "Error parsing cached footer links, go fetch them",
            error
          );
        }
      }

      this.context.spHttpClient
        .get(url, SPHttpClient.configurations.v1)
        .then(async (response: SPHttpClientResponse) => {
          if (response.ok) {
            const spItems: ISPItem[] = (await response.json()).value;

            const links: IFooterLink[] = spItems
              .filter((link) => !!link.Title && !!link.link)
              .map((item, index) => {
                return {
                  name: item.Title || "",
                  link: { url: item.link.Url, desc: item.link.Description },
                  icon: item.icon
                    ? { url: item.icon.Url, desc: item.icon.Description }
                    : undefined,
                  order: item.linkOrder || index,
                  id: item.ID,
                };
              });

            const localStoreItem = {
              value: links,
              expires: expirationTime,
            };

            localStorage.setItem(cacheKey, JSON.stringify(localStoreItem));

            resolve(links);
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
