import * as React from "react";
import * as ReactDom from "react-dom";
import {
  IPropertyPaneField,
  PropertyPaneFieldType,
} from "@microsoft/sp-property-pane";
import { IPropertyPaneMonacoEditorInternalProps } from "./IPropertyPaneMonacoEditorInternalProps";
import { IPropertyPaneMonacoEditorProps } from "./IPropertyPaneMonacoEditorProps";
import { MonacoEditorModal } from "../../components/MonacoEditorModal/MonacoEditorModal";
import { DynamicProperty } from "@microsoft/sp-component-base";
import get from "lodash/get";
import set from "lodash/set";

class PropertyPaneMonacoEditor<T extends object>
  implements IPropertyPaneField<IPropertyPaneMonacoEditorProps<T>>
{
  public type: PropertyPaneFieldType = PropertyPaneFieldType.Custom;
  public targetProperty: string;
  public properties: IPropertyPaneMonacoEditorInternalProps<T>;

  private elem: HTMLElement;

  constructor(
    targetProperty: string,
    properties: IPropertyPaneMonacoEditorProps<T>
  ) {
    this.targetProperty = targetProperty;
    this.properties = {
      key: properties.key,
      properties: properties.properties,
      onUpdate: properties.onUpdate,
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

  private onApplyChanges(newValue: string): void {
    const newProps = JSON.parse(newValue);
    const currProps = this.properties.properties;

    for (const key in newProps) {
      if (!Object.hasOwnProperty.call(newProps, key)) {
        continue;
      }

      const newValue = get(newProps, key);
      const currentValue = get(currProps, key);

      if (currentValue?.__type === "DynamicProperty") {
        const currVal = currentValue as DynamicProperty<unknown>;

        if (newValue.value !== null && newValue.value !== "undefined") {
          currVal.setValue(newValue.value);
        }

        if (newValue.reference !== null && newValue.reference !== "undefined") {
          currVal.setReference(newValue.reference._reference);
        }
      } else {
        set(currProps, key, newValue);
      }
    }

    this.properties.onUpdate();
  }

  private onRender(elem: HTMLElement): void {
    if (!this.elem) {
      this.elem = elem;
    }

    const element = React.createElement(MonacoEditorModal, {
      properties: this.properties.properties,
      applyChanges: this.onApplyChanges.bind(this),
    });

    ReactDom.render(element, elem);
  }

  private onDispose(element: HTMLElement): void {
    ReactDom.unmountComponentAtNode(element);
  }
}

/** Method is creating a new instance of the property pane control */
export function PropertyPaneMonacoEditorField<T extends object>(
  targetProperty: string,
  properties: IPropertyPaneMonacoEditorProps<T>
): IPropertyPaneField<IPropertyPaneMonacoEditorProps<T>> {
  return new PropertyPaneMonacoEditor(targetProperty, properties);
}
