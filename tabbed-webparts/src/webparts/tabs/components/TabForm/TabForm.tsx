import * as React from "react";
import styles from "./TabForm.module.scss";
import { ITab } from "../../models/ITab";
import { AppContext } from "../AppContext";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";
import { TextField } from "@fluentui/react/lib/TextField";
import { Stack } from "@fluentui/react/lib/Stack";
import { DefaultButton, PrimaryButton } from "@fluentui/react/lib/Button";
import { Guid } from "@microsoft/sp-core-library";
import {
  highlightWebPart,
  removeWebPartHighlight,
} from "../../utilities/highlightWebPart";

export interface ITabFormProps {
  tab?: ITab;
  onDismiss: () => void;
}

const SECTION_CLASS = "CanvasSection";
const WEBPART_CLASS = "ControlZone";

export const TabForm: React.FC<ITabFormProps> = function TabForm(
  props: ITabFormProps
) {
  const { tab, onDismiss } = props;
  const appContext = React.useContext(AppContext);
  const zone = appContext.domElement?.closest(`.${SECTION_CLASS}`);
  const webParts = Array.from(
    zone?.querySelectorAll(`.${WEBPART_CLASS}`) || []
  );
  const knownIds = [
    appContext.domElement?.closest(`.${WEBPART_CLASS}`)?.id || "",
    ...appContext.tabs.map((t) => t.targetElement),
  ];
  const availableWebParts = webParts.filter(
    (webPart) =>
      !knownIds.includes(webPart.id) || tab?.targetElement === webPart.id
  );
  const targetElemDropdownOptions: IDropdownOption[] = [
    ...availableWebParts.map((wp, i) => {
      const webPartTitle = wp.children[0].textContent;
      const optionTitle = `${webPartTitle} Web Part (${wp.nodeName.toLocaleLowerCase()}${
        wp.id ? `#${wp.id}` : ""
      })`;

      return { key: wp.id, text: optionTitle };
    }),
  ];

  const [selectedWebPart, setSelectedWebPart] = React.useState(
    tab?.targetElement ?? undefined
  );
  const [tabName, setTabName] = React.useState(tab?.tabName ?? "");
  const [tabId, setTabId] = React.useState(tab?.tabId ?? "");
  const [invalidForm, setInvalidForm] = React.useState(false);

  React.useEffect(() => {
    return removeWebPartHighlight;
  }, []);

  return (
    <>
      {availableWebParts.length === 0 && (
        <MessageBar messageBarType={MessageBarType.blocked}>
          There are no available web part. Please add a web part to the column
          first.
        </MessageBar>
      )}
      <form onSubmit={(event) => event.preventDefault()}>
        <Dropdown
          options={targetElemDropdownOptions}
          label="Web Part"
          placeholder="Select a web part"
          defaultSelectedKey={selectedWebPart}
          required
          errorMessage={
            invalidForm && !selectedWebPart
              ? "You must select a web part for the tab."
              : undefined
          }
          onChange={(
            event: React.FormEvent<HTMLDivElement>,
            option?: IDropdownOption | undefined,
            index?: number | undefined
          ) => {
            if (option) {
              highlightWebPart(`${option.key}`);
              setSelectedWebPart(
                typeof option.key === "number"
                  ? option.key.toString()
                  : option.key
              );
            } else {
              setSelectedWebPart(undefined);
            }
          }}
        />
        <TextField
          label="Tab Name"
          description="The name you would like to appear on the tab."
          value={tabName}
          required
          errorMessage={
            invalidForm && tabName?.trim().length === 0
              ? "Tab name is required."
              : undefined
          }
          onChange={(
            event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
            newValue?: string | undefined
          ) => {
            setTabName(newValue ?? "");
          }}
        />
        <TextField
          label="Tab Id"
          description="A unique Id for the tab that can be used to deep link to the tab."
          value={tabId}
          required
          errorMessage={
            invalidForm && tabId?.trim().length === 0
              ? "Tab id is required."
              : undefined
          }
          onChange={(
            event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
            newValue?: string | undefined
          ) => {
            setTabId(newValue ?? "");
          }}
        />
        <Stack
          horizontal
          horizontalAlign="end"
          verticalAlign="center"
          verticalFill
          wrap
          tokens={{ childrenGap: "10" }}
          className={styles.buttonsContainer}
        >
          <PrimaryButton
            onClick={() => {
              const validForm =
                selectedWebPart &&
                tabName?.trim().length > 0 &&
                tabId?.trim().length > 0;

              if (!validForm) {
                setInvalidForm(true);
                return;
              }

              if (tab) {
                const insertIndex = appContext.tabs.findIndex(
                  (t) => t.key === tab.key
                );
                const filteredTabs = appContext.tabs.filter(
                  (t) => t.key !== tab.key
                );

                filteredTabs.splice(insertIndex, 0, {
                  key: tab.key,
                  tabName: tabName,
                  tabId: tabId,
                  targetElement: selectedWebPart,
                });

                appContext.updateProperty(
                  "tabs",
                  [...appContext.tabs],
                  [...filteredTabs]
                );
              } else {
                appContext.updateProperty(
                  "tabs",
                  [...appContext.tabs],
                  [
                    ...appContext.tabs,
                    {
                      key: Guid.newGuid().toString(),
                      tabName: tabName,
                      tabId: tabId,
                      targetElement: selectedWebPart,
                    },
                  ]
                );
              }
              onDismiss();
            }}
          >
            {!!tab ? "Update Tab" : "Add Tab"}
          </PrimaryButton>
          <DefaultButton onClick={onDismiss}>Cancel</DefaultButton>
        </Stack>
      </form>
    </>
  );
};
