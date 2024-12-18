export interface IPropertyPaneMonacoEditorProps<T extends object> {
  /** Unique Key */
  key: string;
  /* Reference to web part properties */
  properties: T;
  /** Callback to refresh the web part */
  onUpdate: () => void;
}
