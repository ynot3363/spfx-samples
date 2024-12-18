import * as React from "react";
import * as ReactDom from "react-dom";
import { DisplayMode, Version } from "@microsoft/sp-core-library";
import {
  type IPropertyPaneConfiguration,
  IPropertyPaneField,
  PropertyPaneDropdown,
  PropertyPaneToggle,
} from "@microsoft/sp-property-pane";
import {
  IDynamicDataCallables,
  IDynamicDataPropertyDefinition,
} from "@microsoft/sp-dynamic-data";
import { BaseClientSideWebPart } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import * as strings from "TabsWebPartStrings";
import { Tabs, ITabsProps } from "./components/Tabs";
import { ITab } from "./models/ITab";
import set from "lodash/set";
import { FONT_SIZES } from "./constants/FontSizes";
import { NestedKeyOf, NestedValueType } from "./types/helperTypes";
import { PropertyPaneDescription } from "./controls/Description/Description";
import { PropertyPaneColorPickerField } from "./controls/ColorPicker/PropertyPaneColorPicker";
import { PropertyPaneMonacoEditorField } from "./controls/Monaco/PropertyPaneMonacoEditor";
import { IPropertyPaneMonacoEditorProps } from "./controls/Monaco/IPropertyPaneMonacoEditorProps";

export interface ITabsWebPartProps {
  /** Collection of tabs */
  tabs: ITab[];
  /** The font-size of the tab name */
  fontSize: string;
  /** Determines if the color palette for the web part, theme or custom */
  themeBased: boolean;
  /** Color of the line under the tab */
  activeTabColor: string;
}

export default class TabsWebPart
  extends BaseClientSideWebPart<ITabsWebPartProps>
  implements IDynamicDataCallables
{
  private _theme: IReadonlyTheme;

  public render(): void {
    let activeTabColor = "#041e42";

    if (this.properties.themeBased) {
      activeTabColor = this?._theme?.palette?.themePrimary || "#041e42";
    } else {
      activeTabColor = this.properties.activeTabColor || "#041e42";
    }

    /** Hide the web part if there are no tabs and the page is in read more. */
    if (
      this.displayMode === DisplayMode.Read &&
      this.properties.tabs.length === 0
    ) {
      const parentElement =
        this.domElement.parentElement?.parentElement?.parentElement;
      if (parentElement) {
        parentElement.style.setProperty("visibility", "hidden");
        parentElement.style.setProperty("display", "none");
      }

      return;
    }

    const element: React.ReactElement<ITabsProps> = React.createElement(Tabs, {
      tabs: this.properties.tabs ?? [],
      fontSize: this.properties.fontSize,
      themeBased: this.properties.themeBased,
      activeTabColor: activeTabColor,
      theme: this._theme,
      displayMode: this.displayMode,
      domElement: this.domElement,
      updateProperty: this._updateWebPartProperty.bind(this),
    });

    ReactDom.render(element, this.domElement);
  }

  public getPropertyDefinitions(): readonly IDynamicDataPropertyDefinition[] {
    return [{ id: "instanceId", title: "Instance Id" }];
  }

  public getPropertyValue(propertyId: string): string {
    switch (propertyId) {
      case "instanceId":
        return this.context.instanceId;
    }

    throw new Error("Bad property id");
  }

  protected onInit(): Promise<void> {
    this.context.dynamicDataSourceManager.initializeSource(this);

    return super.onInit();
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._theme = currentTheme;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }

  protected _updateWebPartProperty<T extends NestedKeyOf<ITabsWebPartProps>>(
    propertyPath: T,
    oldValue: NestedValueType<ITabsWebPartProps, T>,
    newValue: NestedValueType<ITabsWebPartProps, T>
  ): void {
    set(this.properties, propertyPath.toString(), newValue);

    switch (propertyPath) {
      case "themeBased":
        if (!!newValue && this._theme) {
          set(
            this.properties,
            "activeTabColor",
            this?._theme?.palette?.themePrimary || "#041e42"
          );
        }
        break;
    }

    this.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.context.propertyPane.refresh();
    this.render();
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const basicGroupFields: IPropertyPaneField<any>[] = [
      PropertyPaneDropdown("fontSize", {
        ariaLabel: strings.FontSizeFieldLabel,
        label: strings.FontSizeFieldLabel,
        selectedKey: this.properties.fontSize,
        options: FONT_SIZES,
      }),
      PropertyPaneToggle("themeBased", {
        key: "themeBased",
        label: strings.ThemeToggleLabel,
        ariaLabel: strings.ThemeToggleLabel,
        onText: strings.ThemeToggleOnText,
        onAriaLabel: strings.ThemeToggleOnText,
        offText: strings.ThemeToggleOffText,
        offAriaLabel: strings.ThemeToggleOffText,
        checked: this.properties.themeBased,
      }),
      PropertyPaneDescription("themeDesc", {
        text: strings.ThemeToggleDescription,
      }),
    ];

    if (!this.properties.themeBased) {
      basicGroupFields.push(
        PropertyPaneColorPickerField("activeTabColor", {
          key: "activeTabColor",
          label: strings.ActiveTabColorLabel,
          description: strings.ActiveTabColorDescription,
          color: this.properties.activeTabColor,
          properties: this.properties,
          onPropertyChange: this._updateWebPartProperty.bind(this),
        })
      );
    }

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          displayGroupsAsAccordion: true,
          groups: [
            {
              groupName: strings.BasicGroupName,
              isCollapsed: false,
              groupFields: [...basicGroupFields],
            },
            {
              groupName: strings.AdvancedGroupName,
              isCollapsed: true,
              groupFields: [
                PropertyPaneMonacoEditorField("monaco", {
                  key: "monaco",
                  properties: this.properties,
                  onUpdate: () => {
                    this.context.propertyPane.refresh();
                    this.render();
                  },
                } as IPropertyPaneMonacoEditorProps<ITabsWebPartProps>),
              ],
            },
          ],
        },
      ],
    };
  }
}
