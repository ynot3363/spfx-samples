import { IPropertyPaneCustomFieldProps } from "@microsoft/sp-property-pane";
import { IPropertyPaneColorPickerProps } from "./IPropertyPaneColorPickerProps";

export interface IPropertyPaneColorPickerInternalProps<T extends object>
  extends IPropertyPaneColorPickerProps<T>,
    IPropertyPaneCustomFieldProps {}
