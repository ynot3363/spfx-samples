import * as React from "react";
import styles from "./CustomPanel.module.scss";
import { Panel } from "@fluentui/react/lib/Panel";
import { Text } from "@fluentui/react/lib/Text";
import {
  DefaultButton,
  PrimaryButton,
} from "@fluentui/react/lib/components/Button";
import {
  ListViewCommandSetContext,
  RowAccessor,
} from "@microsoft/sp-listview-extensibility";
import { DocumentRow } from "../DocumentRow/DocumentRow";
import {
  DatePicker,
  defaultDatePickerStrings,
} from "@fluentui/react/lib/components/DatePicker";
import { DayOfWeek } from "@fluentui/react";

export interface ICustomPanelState {
  formState: string;
  callsMade: number;
  selectedDate?: Date;
}

export interface ICustomPanelProps {
  context: ListViewCommandSetContext;
  isOpen: boolean;
  items: RowAccessor[];
  panelTitle: string;
  panelDescription: string;
  command: string;
  onClose: () => void;
  onRefresh: () => void;
}

export class CustomPanel extends React.Component<
  ICustomPanelProps,
  ICustomPanelState
> {
  constructor(props: ICustomPanelProps) {
    super(props);
    this.state = {
      formState: "Open",
      callsMade: 0,
      selectedDate: null,
    };

    this._renderDocuments.bind(this);
  }

  public render(): React.ReactElement<ICustomPanelProps> {
    const { isOpen, items, panelTitle, panelDescription, command } = this.props;
    const { selectedDate } = this.state;

    return (
      <Panel
        isOpen={isOpen}
        isFooterAtBottom={true}
        headerText={panelTitle}
        onDismiss={
          this.state.formState === "Open"
            ? this.props.onClose
            : this.props.onRefresh
        }
        onRenderFooterContent={this._renderFooter.bind(this)}
      >
        {panelDescription}
        {command === "UPDATE_PUBLISHING_DATE" && (
          <div>
            <div className={styles.label}>
              <Text variant="mediumPlus">
                <b>Publish Date:</b>
              </Text>
            </div>
            <div className={styles.datePickerContainer}>
              <DatePicker
                firstDayOfWeek={DayOfWeek.Sunday}
                placeholder="Select a date..."
                ariaLabel="Select a date"
                strings={defaultDatePickerStrings}
                onSelectDate={(date: Date) => {
                  if (date) {
                    this.setState({
                      selectedDate: date,
                    });
                  }
                }}
                value={selectedDate}
              />
            </div>
          </div>
        )}
        <div className={styles.label}>
          <Text variant="mediumPlus">
            <b>Selected Page{items.length > 1 ? "s" : ""}:</b>
          </Text>
        </div>
        {this._renderDocuments()}
      </Panel>
    );
  }

  private _renderFooter(): React.ReactFragment {
    const buttonStyles = { root: { marginRight: 8 } };
    const { command } = this.props;
    const { formState, selectedDate } = this.state;
    let element: JSX.Element = (
      <>
        <PrimaryButton
          styles={buttonStyles}
          disabled={
            formState === "Submitting" ||
            (command === "UPDATE_PUBLISHING_DATE" && !selectedDate)
          }
          onClick={() => {
            this.setState({
              formState: "Submitting",
            });
          }}
        >
          Submit
        </PrimaryButton>
        <DefaultButton
          disabled={formState === "Submitting"}
          onClick={() => {
            this.setState(
              {
                formState: "Open",
                callsMade: 0,
              },
              () => {
                this.props.onClose();
              }
            );
          }}
        >
          Cancel
        </DefaultButton>
      </>
    );

    if (formState === "Complete") {
      element = (
        <PrimaryButton onClick={this.props.onRefresh}>Close</PrimaryButton>
      );
    }
    return element;
  }

  private _renderDocuments(): React.ReactFragment {
    const { items, command, context } = this.props;
    const { selectedDate } = this.state;

    return (
      <div className={styles.documentContainer}>
        {items.map((item: RowAccessor) => {
          return (
            <DocumentRow
              item={item}
              formState={this.state.formState}
              command={command}
              context={context}
              date={selectedDate}
              updateCallsMade={() =>
                this.setState(
                  (prevState) => {
                    return { callsMade: prevState.callsMade + 1 };
                  },
                  () => {
                    if (this.state.callsMade === this.props.items.length) {
                      this.setState({
                        formState: "Complete",
                      });
                    }
                  }
                )
              }
            />
          );
        })}
      </div>
    );
  }
}
