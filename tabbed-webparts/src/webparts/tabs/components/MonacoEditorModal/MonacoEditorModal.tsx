import * as React from "react";
import styles from "./MonacoEditorModal.module.scss";
import Modal from "@fluentui/react/lib/Modal";
import {
  CompoundButton,
  DefaultButton,
  IButtonStyles,
  IconButton,
  PrimaryButton,
} from "@fluentui/react/lib/Button";
import { Stack, StackItem } from "@fluentui/react/lib/Stack";
import { Toggle } from "@fluentui/react/lib/Toggle";
import { useBoolean } from "@fluentui/react-hooks/lib/useBoolean";
import MonacoEditor from "./MonacoEditor/MonacoEditor";

export interface IMonacoEditorModalProps<T extends object> {
  properties: T;
  applyChanges: (newValue: string) => void;
}

export const MonacoEditorModal = function MonacoEditorModal<T extends object>(
  props: IMonacoEditorModalProps<T>
): React.ReactElement<IMonacoEditorModalProps<T>> {
  const [showModal, { toggle: toggleModal }] = useBoolean(false);
  const [wrapText, { toggle: toggleWrapText }] = useBoolean(false);
  const [jsonString, setJSONString] = React.useState(
    JSON.stringify(props.properties, null, "\t")
  );
  const buttonStyles: IButtonStyles = {
    root: {
      margin: "10px 0px",
    },
  };

  let fileRef: HTMLInputElement;

  const onUpload = (): void => {
    if (!fileRef || !fileRef.files || fileRef.files.length === 0) return;

    if (fileRef.files[0].type === "application/json") {
      const fileReader: FileReader = new FileReader();
      fileReader.readAsText(fileRef.files[0]);
      fileReader.onload = () => {
        let _jsonString = fileReader.result as string;
        const json = JSON.parse(_jsonString); //normalize as an object
        _jsonString = JSON.stringify(json, null, "\t"); // format back to string
        fileRef.value = "";
        setJSONString(_jsonString);
        toggleModal();
      };
    } else {
      alert("You can only upload JSON files.");
    }
  };

  const onDismiss = (): void => {
    setJSONString(JSON.stringify(props.properties, null, "\t"));
    toggleModal();
  };

  return (
    <>
      <Stack horizontalAlign="stretch" verticalAlign="stretch">
        <CompoundButton
          iconProps={{ iconName: "BuildDefinition" }}
          onClick={() => {
            setJSONString(JSON.stringify(props.properties, null, "\t"));
            toggleModal();
          }}
          styles={buttonStyles}
          text="Configure Properties"
          secondaryText="Open an editor to modify the web part properties"
        />
        <CompoundButton
          iconProps={{ iconName: "Download" }}
          onClick={() => {
            const json = JSON.stringify(JSON.parse(jsonString), null, "\t");
            const blob = new Blob([json], { type: "octet/stream" });
            const url = window.URL.createObjectURL(blob);
            const a = document.createElement("a");
            a.setAttribute("style", "display:none");
            a.setAttribute("data-interception", "off");
            a.href = url;
            a.download = "webpartproperties.json";
            document.body.appendChild(a);
            a.click();
            window.URL.revokeObjectURL(url);
          }}
          styles={buttonStyles}
          text="Download Configuration"
          secondaryText="Saves the web part's properties to a JSON file."
        />
        <CompoundButton
          iconProps={{ iconName: "Upload" }}
          onClick={() => {
            fileRef.click();
          }}
          styles={buttonStyles}
          text="Upload Configuration"
          secondaryText="Allows you to upload a JSON configuration file. Will open the modal for inspection."
        />
      </Stack>
      <Modal isOpen={showModal} isBlocking={true} onDismiss={onDismiss}>
        <div className={styles.configureContainer}>
          <div className={styles.configureHeader}>
            <h2 id="modalTitle">Edit Web Part Properties</h2>
            <IconButton
              className={styles.dismissIconButton}
              iconProps={{ iconName: "Cancel" }}
              aria-label="Cloase popup modal"
              onClick={onDismiss}
            />
          </div>
          <div className={styles.configureBody}>
            <Stack horizontal horizontalAlign="end">
              <StackItem>
                <Toggle
                  label="Wrap text"
                  inlineLabel
                  onText="On"
                  offText="Off"
                  onChange={toggleWrapText}
                />
              </StackItem>
            </Stack>
            <div className={styles.containerClass}>
              <MonacoEditor
                jsonString={jsonString}
                readOnly={false}
                wrapText={wrapText}
                onChange={(currentValue) => setJSONString(currentValue)}
              />
            </div>
          </div>
          <div className={styles.configureFooter}>
            <PrimaryButton
              onClick={() => {
                props.applyChanges(jsonString);
                toggleModal();
              }}
            >
              Apply
            </PrimaryButton>
            <DefaultButton onClick={onDismiss}>Close</DefaultButton>
          </div>
        </div>
      </Modal>
      <input
        type="file"
        accept=".json"
        multiple={false}
        id="uploadwebpartjson"
        ref={(element: HTMLInputElement) => {
          fileRef = element;
        }}
        style={{ display: "none" }}
        onChange={() => {
          onUpload();
        }}
      />
    </>
  );
};
