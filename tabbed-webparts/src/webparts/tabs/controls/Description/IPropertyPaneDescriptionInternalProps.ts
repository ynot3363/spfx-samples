import { IPropertyPaneCustomFieldProps } from "@microsoft/sp-property-pane";
import { IPropertyPaneDescriptionProps } from "./IPropertyPaneDescriptionProps";

export interface IPropertyPaneDescriptionInternalProps
  extends IPropertyPaneDescriptionProps,
    IPropertyPaneCustomFieldProps {}
