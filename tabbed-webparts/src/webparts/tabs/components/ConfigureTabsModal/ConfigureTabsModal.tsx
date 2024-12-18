import * as React from "react";
import { AppContext } from "../AppContext";
import { ITab } from "../../models/ITab";
import { IDragOptions } from "@fluentui/react/lib/Modal";
import { ContextualMenu } from "@fluentui/react/lib/ContextualMenu";
import { Dropdown, IDropdownOption } from "@fluentui/react/lib/Dropdown";
import {
  DetailsList,
  IColumn,
  Selection,
  SelectionMode,
} from "@fluentui/react/lib/DetailsList";
import { Stack, StackItem } from "@fluentui/react/lib/Stack";
import { ActionButton } from "@fluentui/react/lib/Button";
import { TooltipHost } from "@fluentui/react/lib/Tooltip";
import { TabForm } from "../TabForm/TabForm";
import {
  highlightWebPart,
  removeWebPartHighlight,
} from "../../utilities/highlightWebPart";
import { CustomModal } from "../CustomModal/CustomModal";

export interface IConfigureTabsModalProps {
  isOpen: boolean;
  onDismiss: () => void;
}

const SECTION_CLASS = "CanvasSection";
const WEBPART_CLASS = "ControlZone";

export const ConfigureTabsModal: React.FC<IConfigureTabsModalProps> =
  function ConfigureTabsModal(props: IConfigureTabsModalProps) {
    const { isOpen, onDismiss } = props;
    const appContext = React.useContext(AppContext);
    const [selectedTab, setSelectedTab] = React.useState<ITab | undefined>(
      undefined
    );
    const [showTabForm, setShowTabForm] = React.useState(false);
    const dragOptions = React.useMemo(
      (): IDragOptions => ({
        moveMenuItemText: "Move",
        closeMenuItemText: "Close",
        menu: ContextualMenu,
        keepInBounds: true,
        dragHandleSelector: ".ms-Modal-scrollableContent > div:first-child",
      }),
      []
    );
    const sortOptions: IDropdownOption[] = Array.from(
      appContext.tabs,
      (v, k) => {
        const num = k + 1;
        return { key: num, text: num.toString() };
      }
    );
    const columns: IColumn[] = [
      {
        key: "sort",
        name: "Sort Order",
        minWidth: 100,
        onRender(item, index) {
          return (
            <Dropdown
              data-selection-disabled={true}
              options={sortOptions}
              selectedKey={index !== undefined ? index + 1 : undefined}
              onChange={(
                event: React.FormEvent<HTMLDivElement>,
                option?: IDropdownOption
              ) => {
                if (option) {
                  const insertIndex = (option.key as number) - 1;
                  const filteredTabs = appContext.tabs.filter(
                    (t) => t.key !== item.key
                  );

                  filteredTabs.splice(insertIndex, 0, item);

                  appContext.updateProperty(
                    "tabs",
                    [...appContext.tabs],
                    filteredTabs
                  );
                }
              }}
            />
          );
        },
      },
      { key: "tabName", name: "Tab Name", minWidth: 200 },
      { key: "tabId", name: "Tab Id", minWidth: 100 },
      { key: "targetElement", name: "Target Element", minWidth: 300 },
    ];
    const selection = React.useMemo(
      () =>
        new Selection<ITab>({
          onSelectionChanged: () => {
            const selectedItems = selection.getSelection();
            if (selectedItems.length === 0) {
              setSelectedTab(undefined);
              removeWebPartHighlight();
            }
            if (selectedItems.length === 1) {
              highlightWebPart(`${selectedItems[0].targetElement}`);
              setSelectedTab({ ...selectedItems[0] });
            }
          },
          selectionMode: SelectionMode.single,
        }),
      []
    );
    const zone = appContext.domElement?.closest(`.${SECTION_CLASS}`);
    const webParts = Array.from(
      zone?.querySelectorAll(`.${WEBPART_CLASS}`) || []
    );
    const tabWebPartId =
      appContext.domElement?.closest(`.${WEBPART_CLASS}`)?.id || "";
    const knownIds = [
      tabWebPartId,
      ...appContext.tabs.map((t) => t.targetElement),
    ];
    const availableWebParts = webParts.filter(
      (webPart) => !knownIds.includes(webPart.id)
    );

    React.useEffect(() => {
      return removeWebPartHighlight;
    }, []);

    return (
      <CustomModal
        isOpen={isOpen}
        isBlocking={true}
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        dragOptions={dragOptions as any}
        onDismiss={onDismiss}
      >
        <CustomModal.Header
          title={
            showTabForm === false
              ? "Configure Tabs"
              : selectedTab === undefined
              ? "Add Tab"
              : "Edit Tab"
          }
          onDismiss={onDismiss}
        >
          <Stack horizontal>
            {showTabForm === false && (
              <StackItem>
                <TooltipHost
                  content={
                    availableWebParts.length === 0
                      ? "There are no available web parts to add. Please add a new web part to the column first."
                      : "Click to add a new tab."
                  }
                  id={`AddButton-${tabWebPartId}`}
                  calloutProps={{ gapSpace: 0 }}
                >
                  <ActionButton
                    disabled={availableWebParts.length === 0}
                    iconProps={{ iconName: "Add" }}
                    onClick={() => {
                      const currentSelection = selection.getSelection();
                      if (currentSelection.length === 1) {
                        selection.setKeySelected(
                          currentSelection[0].key,
                          false,
                          false
                        );
                      }
                      setShowTabForm(true);
                    }}
                  >
                    Add Tab
                  </ActionButton>
                </TooltipHost>
              </StackItem>
            )}
            {showTabForm === false && selectedTab !== undefined && (
              <>
                <StackItem>
                  <ActionButton
                    iconProps={{ iconName: "Edit" }}
                    onClick={() => {
                      setShowTabForm(true);
                    }}
                  >
                    Edit Tab
                  </ActionButton>
                </StackItem>
                <StackItem>
                  <ActionButton
                    iconProps={{ iconName: "Delete" }}
                    onClick={() => {
                      const filteredTabs = appContext.tabs.filter(
                        (t) => t.key !== selectedTab.key
                      );
                      appContext.updateProperty(
                        "tabs",
                        appContext.tabs,
                        filteredTabs
                      );
                    }}
                  >
                    Delete Tab
                  </ActionButton>
                </StackItem>
              </>
            )}
          </Stack>
        </CustomModal.Header>
        <CustomModal.Body>
          <>
            {showTabForm === false && (
              <DetailsList
                setKey="items"
                items={[...appContext.tabs]}
                columns={columns}
                selection={selection}
                onRenderItemColumn={(
                  item?: ITab,
                  index?: number | undefined,
                  column?: IColumn | undefined
                ) => {
                  const key = column?.key as keyof ITab;
                  if (key && item) {
                    return String(item[key]);
                  }
                }}
                ariaLabelForSelectionColumn="Toggle selection"
                ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                checkButtonAriaLabel="select row"
              />
            )}
            {showTabForm === true && (
              <TabForm
                tab={selectedTab}
                onDismiss={() => {
                  const currentSelection = selection.getSelection();
                  if (currentSelection.length === 1) {
                    selection.setKeySelected(
                      currentSelection[0].key,
                      false,
                      false
                    );
                  }
                  setShowTabForm(false);
                  return null;
                }}
              />
            )}
          </>
        </CustomModal.Body>
        <CustomModal.Footer
          closeModalLabel="Close"
          onDismiss={() => {
            onDismiss();
            setShowTabForm(false);
          }}
        />
      </CustomModal>
    );
  };
