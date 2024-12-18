import { IColorCellProps } from "@fluentui/react/lib/SwatchColorPicker";
import { ITheme } from "@fluentui/react/lib/Theme";

export interface IPropertyPaneColorPickerProps<T extends object> {
  /** Unique key to identify the component */
  key: string;
  /** Web part properties */
  properties: T;
  /** Label for the control */
  label: string;
  /** The selected color */
  color: string;
  /** Description for the control */
  description?: string;
  /** Determines if the preview box should be shown or not */
  showPreview?: boolean;
  /** Delay updates from the control by x milliseconds, default 500 */
  debounce?: 500;
  /**
   * alpha - (the default) means display a slider and text field for editing alpha values.
   * transparency - also displays a slider and textfield but for editing transparency values.
   * none - hides these controls.
   * Alpha represents the opacity of the color, whereas transparency represents the transparentness
   * of the color: e.g.a 30% transparent color has 70% opaqueness.
   */
  alphaType?: "alpha" | "transparency" | "none";
  /** Ability to pass the current theme */
  theme?: ITheme;
  /** Ability to set your own color swatches */
  colorSwatchs?: IColorCellProps[];
  /** Ability to control the column count default is 9 */
  columnCount?: number;
  /** Callback for change events */
  onPropertyChange: (
    propertyPath: string,
    oldValue: string,
    newValue: string
  ) => void;
}
