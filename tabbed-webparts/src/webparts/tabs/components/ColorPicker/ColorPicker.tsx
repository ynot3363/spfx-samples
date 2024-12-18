import * as React from "react";
import styles from "./ColorPicker.module.scss";
import { useBoolean } from "@fluentui/react-hooks/lib/useBoolean";
import { ColorPicker } from "@fluentui/react/lib/ColorPicker";
import { IColor } from "@fluentui/react/lib/Color";
import {
  IColorCellProps,
  SwatchColorPicker,
} from "@fluentui/react/lib/SwatchColorPicker";
import { ActionButton } from "@fluentui/react/lib/Button";
import { ITheme } from "@fluentui/react/lib/Theme";

export interface IColorPickerFieldProps {
  /** Label for the control */
  label: string;
  /** The selected color */
  color: string;
  /** Description for the control */
  description?: string;
  /** Determine if the selected color should be hidden or not */
  hideSelectedColorValue?: boolean;
  /** The default control to show */
  showSwatchOrPicker?: "picker" | "swatch";
  /** Determines if the preview box should be shown or not */
  showPreview?: boolean;
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
  /** Callback for change events */
  onChange: (newColor: string) => void;
  /** Ability to set your own color swatches */
  colorSwatchs?: IColorCellProps[];
  /** Ability to control the column count default is 9 */
  columnCount?: number;
}

export const ColorPickerField: React.FC<IColorPickerFieldProps> =
  function ColorPickerField(props: IColorPickerFieldProps): JSX.Element {
    const {
      alphaType,
      color,
      colorSwatchs,
      columnCount,
      description,
      hideSelectedColorValue,
      label,
      showPreview,
      showSwatchOrPicker,
      onChange,
    } = props;
    const [selectedColor, setSelectedColor] = React.useState(color ?? "");
    const [showColorSwatches, { toggle: toggleShowColorSwatches }] = useBoolean(
      showSwatchOrPicker && showSwatchOrPicker === "swatch" ? true : false
    );
    const [showColorPicker, { toggle: toggleShowColorPicker }] = useBoolean(
      showSwatchOrPicker && showSwatchOrPicker === "picker" ? true : false
    );
    const defaultColorSwatches: IColorCellProps[] = [
      { id: "#FF1921", label: "Red", color: "#FF1921" },
      { id: "#FFC12E", label: "Orange", color: "#FFC12E" },
      { id: "#FEFF37", label: "Yellow", color: "#FEFF37" },
      { id: "#90D057", label: "Light Green", color: "#90D057" },
      { id: "#00B053", label: "Green", color: "#00B053" },
      { id: "#00AFED", label: "Light Blue", color: "#00AFED" },
      { id: "#006EBD", label: "Blue", color: "#006EBD" },
      { id: "#011F5E", label: "Dark Blue", color: "#011F5E" },
      { id: "#712F9E", label: "Purple", color: "#712F9E" },
    ];

    return (
      <div className={styles.container}>
        {!!label && (
          <>
            <span className={styles.label}>{label}</span>
            <br />
          </>
        )}
        {!hideSelectedColorValue && (
          <div className={styles.colorContainer}>
            <div
              className={styles.colorSquare}
              style={{ backgroundColor: selectedColor }}
            />
            <span className={styles.hexCode}>{selectedColor}</span>
          </div>
        )}
        <div>
          <ActionButton
            iconProps={{ iconName: "FontColorSwatch" }}
            onClick={() => {
              if (showColorPicker) toggleShowColorPicker();
              toggleShowColorSwatches();
            }}
          >
            Swatches
          </ActionButton>
          <ActionButton
            iconProps={{ iconName: "Color" }}
            onClick={() => {
              if (showColorSwatches) toggleShowColorSwatches();
              toggleShowColorPicker();
            }}
          >
            Color Picker
          </ActionButton>
        </div>
        <span className={styles.description}>{description}</span>
        {showColorSwatches && (
          <SwatchColorPicker
            columnCount={columnCount ?? 9}
            cellHeight={24}
            cellWidth={24}
            cellShape="square"
            colorCells={colorSwatchs ?? defaultColorSwatches}
            selectedId={selectedColor}
            onChange={(
              event: React.FormEvent<HTMLElement>,
              id: string | undefined,
              color: string | undefined
            ) => {
              if (id !== undefined) {
                setSelectedColor(id);
                onChange(id);
              }
            }}
          />
        )}
        {showColorPicker && (
          <ColorPicker
            alphaType={alphaType ?? "alpha"}
            color={selectedColor}
            showPreview={showPreview ?? true}
            onChange={(
              ev: React.SyntheticEvent<HTMLElement, Event>,
              color: IColor
            ) => {
              setSelectedColor(color.str);
              onChange(color.str);
            }}
          />
        )}
      </div>
    );
  };
