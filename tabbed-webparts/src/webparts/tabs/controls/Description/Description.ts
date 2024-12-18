import * as React from "react";
import * as ReactDom from "react-dom";
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
} from "@microsoft/sp-property-pane";
import { IPropertyPaneDescriptionInternalProps } from "./IPropertyPaneDescriptionInternalProps";
import { IPropertyPaneDescriptionProps } from "./IPropertyPaneDescriptionProps";
import {
  Description,
  IDescriptionProps,
} from "../../components/Description/Description";

class DescriptionField
  implements IPropertyPaneField<IPropertyPaneDescriptionProps>
{
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneDescriptionInternalProps;

  private elem: HTMLElement;

  constructor(
    targetProperty: string,
    properties: IPropertyPaneDescriptionProps
  ) {
    this.targetProperty = targetProperty;
    this.properties = {
      key: properties.text,
      text: properties.text,
      onRender: this.onRender.bind(this),
      onDispose: this.onDispose.bind(this),
    };
  }

  public render(): void {
    if (!this.elem) {
      return;
    }

    this.onRender(this.elem);
  }

  private onRender(elem: HTMLElement): void {
    if (!this.elem) {
      this.elem = elem;
    }

    const element: React.FunctionComponentElement<IDescriptionProps> =
      React.createElement(Description, { text: this.properties.text });

    ReactDom.render(element, elem);
  }

  private onDispose(element: HTMLElement): void {
    ReactDom.unmountComponentAtNode(element);
  }
}

/** Method is creating a new instance of the property pane control */
export function PropertyPaneDescription(
  targetProperty: string,
  properties: IPropertyPaneDescriptionProps
): IPropertyPaneField<IPropertyPaneDescriptionProps> {
  const newProperties: IPropertyPaneDescriptionProps = {
    text: properties.text,
  };

  return new DescriptionField(targetProperty, newProperties);
}
