import * as React from "react";
import * as ReactDom from "react-dom";
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
} from "@microsoft/sp-property-pane";
import { IPropertyPaneColorPickerInternalProps } from "./IPropertyPaneColorPickerInternalProps";
import { IPropertyPaneColorPickerProps } from "./IPropertyPaneColorPickerProps";
import { ColorPickerField } from "../../components/ColorPicker/ColorPicker";
import set from "lodash/set";

class PropertyPaneColorPicker<T extends object>
  implements IPropertyPaneField<IPropertyPaneColorPickerProps<T>>
{
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneColorPickerInternalProps<T>;

  private elem: HTMLElement;
  private color: string;
  private timer: ReturnType<typeof setTimeout>;

  constructor(
    targetProperty: string,
    properties: IPropertyPaneColorPickerProps<T>
  ) {
    this.targetProperty = targetProperty;
    this.color = properties.color;
    this.properties = {
      key: properties.key,
      label: properties.label,
      properties: properties.properties,
      description: properties.description || "",
      color: properties.color,
      showPreview: properties.showPreview || true,
      alphaType: properties.alphaType || "alpha",
      theme: properties.theme || undefined,
      columnCount: properties.columnCount || 9,
      colorSwatchs: properties.colorSwatchs || undefined,
      debounce: properties?.debounce || 500,
      onPropertyChange: properties.onPropertyChange,
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

  private onChange(newColor: string): void {
    const newValue = newColor;
    const oldValue = this.color;

    this.color = newColor;

    if (this.properties.debounce && this.properties.debounce > 0) {
      clearTimeout(this.timer);
      this.timer = setTimeout(() => {
        this.onChangeInternal(oldValue, newValue);
      }, this.properties.debounce);
    } else {
      this.onChangeInternal(oldValue, newValue);
    }
  }

  private onChangeInternal(oldValue: string, newValue: string): void {
    if (!this.properties) return;

    set(this.properties.properties, this.targetProperty, newValue);
    this.properties.onPropertyChange(this.targetProperty, oldValue, newValue);
  }

  private onRender(elem: HTMLElement): void {
    if (!this.elem) {
      this.elem = elem;
    }

    const element = React.createElement(ColorPickerField, {
      label: this.properties.label,
      description: this.properties.description,
      color: this.color,
      showPreview: this.properties.showPreview,
      alphaType: this.properties.alphaType,
      theme: this.properties.theme,
      columnCount: this.properties.columnCount,
      colorSwatchs: this.properties.colorSwatchs,
      onChange: this.onChange.bind(this),
    });

    ReactDom.render(element, elem);
  }

  private onDispose(element: HTMLElement): void {
    ReactDom.unmountComponentAtNode(element);
  }
}

/** Method is creating a new instance of the property pane control */
export function PropertyPaneColorPickerField<T extends object>(
  targetProperty: string,
  properties: IPropertyPaneColorPickerProps<T>
): IPropertyPaneField<IPropertyPaneColorPickerProps<T>> {
  return new PropertyPaneColorPicker(targetProperty, properties);
}
