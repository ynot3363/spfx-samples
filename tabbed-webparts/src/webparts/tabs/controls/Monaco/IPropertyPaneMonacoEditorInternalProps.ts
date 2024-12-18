import { IPropertyPaneCustomFieldProps } from "@microsoft/sp-property-pane";
import { IPropertyPaneMonacoEditorProps } from "./IPropertyPaneMonacoEditorProps";

export interface IPropertyPaneMonacoEditorInternalProps<T extends object>
  extends IPropertyPaneMonacoEditorProps<T>,
    IPropertyPaneCustomFieldProps {}
