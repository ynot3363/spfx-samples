import * as React from "react";
import * as ReactDom from "react-dom";
import * as strings from "ReduxExampleWebPartWebPartStrings";
import { Version } from "@microsoft/sp-core-library";
import { type IPropertyPaneConfiguration } from "@microsoft/sp-property-pane";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import ReduxProvider from "./components/ReduxProvider";
import { IReduxExampleWebPartProps } from "./components/ReduxExampleWebPart";
import { ServiceManager } from "./services/ServiceManager";
import { customElementHelper, Providers } from "@microsoft/mgt-element";
import { SharePointProvider } from "@microsoft/mgt-sharepoint-provider";
import { lazyLoadComponent } from "@microsoft/mgt-spfx-utils";

const ReduxExampleWebPart = React.lazy(
  () =>
    import(
      /* webpackChunkName: 'mgt-redux-example-component' */ "./components/ReduxExampleWebPart"
    )
);

export interface IReduxExampleWebPartWebPartProps {}

export default class ReduxExampleWebPartWebPart extends BaseClientSideWebPart<IReduxExampleWebPartWebPartProps> {
  public render(): void {
    const element = lazyLoadComponent<IReduxExampleWebPartProps>(
      ReduxExampleWebPart,
      {
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
      }
    );

    const ReduxProviderElement = React.createElement(
      ReduxProvider,
      {
        initialState: {
          user: {
            id: this.context.pageContext.aadInfo.userId.toString(),
            name: this.context.pageContext.user.displayName,
            givenName: "",
            surname: "",
            email: this.context.pageContext.user.email,
            department: "",
            jobTitle: "",
            refreshing: false,
          },
          webpart: {
            instanceId: this.context.instanceId,
            url: this.context.pageContext.web.absoluteUrl,
          },
        },
      },
      element
    );

    ReactDom.render(ReduxProviderElement, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();

    ServiceManager.initialize(this.context.serviceScope);

    if (!Providers.globalProvider) {
      Providers.globalProvider = new SharePointProvider(this.context);
    }
    customElementHelper.withDisambiguation("redux-solution");

    return Promise.resolve();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    const { semanticColors } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty(
        "--bodyText",
        semanticColors.bodyText || null
      );
      this.domElement.style.setProperty("--link", semanticColors.link || null);
      this.domElement.style.setProperty(
        "--linkHovered",
        semanticColors.linkHovered || null
      );
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [],
            },
          ],
        },
      ],
    };
  }
}
