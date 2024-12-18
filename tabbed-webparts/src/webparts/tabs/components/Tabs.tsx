import * as React from "react";
import styles from "./Tabs.module.scss";
import { ITabsWebPartProps } from "../TabsWebPart";
import { DisplayMode } from "@microsoft/sp-core-library";
import { IReadonlyTheme } from "@microsoft/sp-component-base";
import { AppContext } from "./AppContext";
import { ITheme } from "@fluentui/react/lib/Theme";
import { ITab } from "../models/ITab";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { Pivot, PivotItem } from "@fluentui/react/lib/Pivot";
import { ConfigureTabsModal } from "./ConfigureTabsModal/ConfigureTabsModal";
import { NestedKeyOf, NestedValueType } from "../types/helperTypes";

export interface ITabsProps extends ITabsWebPartProps {
  displayMode: DisplayMode;
  domElement: HTMLElement | undefined;
  theme: IReadonlyTheme | undefined;
  updateProperty: <T extends NestedKeyOf<ITabsWebPartProps>>(
    propertyPath: T,
    oldValue: NestedValueType<ITabsWebPartProps, T>,
    newValue: NestedValueType<ITabsWebPartProps, T>
  ) => void;
}

export const Tabs: React.FC<ITabsProps> = function Tabs(props: ITabsProps) {
  const tabId = location.hash ? location.hash.replace("#", "") : undefined;
  const knownTabId =
    tabId === undefined
      ? false
      : props.tabs.map((t) => t.tabId).includes(tabId);
  const [showTabConfig, setShowTabConfig] = React.useState(false);
  const [selectedTab, setSelectedTab] = React.useState(
    knownTabId
      ? tabId
      : props.tabs.length === 0
      ? undefined
      : props.tabs[0].tabId
  );

  function _getTabId(itemKey: string | undefined): string {
    return itemKey ?? "";
  }

  React.useEffect(() => {
    if (props.displayMode === DisplayMode.Read) {
      const tabPanel = props.domElement?.querySelector(`.${styles.tabPanel}`);

      props.tabs.forEach((tab, index) => {
        const targetElement = document.getElementById(tab.targetElement);

        if (targetElement === null) return;

        if (knownTabId) {
          if (tab.tabId === tabId) {
            targetElement.classList.add(styles.activeTab);
            setTimeout(() => {
              props.domElement?.scrollIntoView({
                behavior: "smooth",
                block: "start",
                inline: "nearest",
              });
            }, 500);
          } else {
            targetElement.classList.add(styles.inactiveTab);
          }
        } else {
          if (index === 0) {
            targetElement.classList.add(styles.activeTab);
          } else {
            targetElement.classList.add(styles.inactiveTab);
          }
        }

        tabPanel?.insertAdjacentElement("beforeend", targetElement);
      });
    }

    return () => {
      props.tabs.forEach((tab) => {
        const targetElement = document.getElementById(tab.targetElement);
        if (targetElement === null) return;
        targetElement.classList.remove(styles.activeTab);
        targetElement.classList.remove(styles.inactiveTab);
      });
    };
  }, [props.displayMode]);

  return (
    <AppContext.Provider value={props}>
      <section className={styles.tabs}>
        {props.displayMode === DisplayMode.Edit && (
          <>
            <div className={styles.editMode}>
              <h2>Web Parts Tabs</h2>
              <div className={styles.description}>
                <span>
                  The Web Parts Tabs solution allows you to collapse the web
                  parts in the same column into tabs when the page is in Read
                  mode.
                </span>
                <br />
                <br />
                <span>
                  When editing the page the tabs are shown so you can configure
                  the user interface however the web parts will not appear under
                  the tabs. This allows you to still configure and edit those
                  web parts without an issue.
                </span>
                <br />
                <br />
                <strong>
                  This solution will only allow you to collapse web parts within
                  the same column. Ensure there are web parts added before
                  trying to configure the solution.
                </strong>
                <br />
                <br />
                <span>
                  Click on{" "}
                  {props.tabs.length === 0 ? "Add Tabs" : "Configure Tabs"}{" "}
                  button to get started.
                </span>
              </div>
              <div className={styles.configureTabsButton}>
                <PrimaryButton
                  onClick={() => {
                    setShowTabConfig(!showTabConfig);
                  }}
                  theme={props.theme as ITheme}
                >
                  {props.tabs.length === 0 ? "Add Tabs" : "Configure Tabs"}
                </PrimaryButton>
              </div>
            </div>
            {props.tabs.length > 0 && <hr />}
            {showTabConfig && (
              <ConfigureTabsModal
                isOpen={showTabConfig}
                onDismiss={() => setShowTabConfig(false)}
              />
            )}
          </>
        )}
        {props.tabs.length > 0 && (
          <>
            <Pivot
              className={styles.pivotTabs}
              headersOnly={true}
              getTabId={_getTabId}
              selectedKey={selectedTab}
              onLinkClick={(item?: PivotItem | undefined) => {
                if (item === undefined) return;

                if (props.displayMode === DisplayMode.Read) {
                  const activeTabs = props.domElement?.querySelectorAll(
                    `.${styles.activeTab}`
                  );
                  activeTabs?.forEach((el) => {
                    el.classList.remove(styles.activeTab);
                    el.classList.add(styles.inactiveTab);
                  });

                  const selectedTab = props.tabs.find(
                    (tab: ITab) => tab.tabId === item.props.itemKey
                  );

                  if (selectedTab === undefined) return;

                  const elem = document.getElementById(
                    selectedTab?.targetElement
                  );
                  elem?.classList.add(styles.activeTab);
                  elem?.classList.remove(styles.inactiveTab);
                }

                setSelectedTab(item.props.itemKey);
                history.pushState({}, "", `#${item.props.itemKey}`);
              }}
              styles={{
                root: {
                  display: "flex",
                  alignItems: "stretch",
                  whiteSpace: "normal",
                },
                linkIsSelected: {
                  fontSize: props.fontSize,
                  lineHeight: props.fontSize,
                  minHeight: 40,
                  height: "auto",
                  padding: "16px 8px",
                  width: `${100 / props.tabs.length}%`,
                  backgroundColor:
                    props.theme?.semanticColors?.inputBackground || "#FFFFFF",
                  border: `1px solid ${props.activeTabColor}`,
                  borderBottom: "none",
                  borderTopLeftRadius: 10,
                  borderTopRightRadius: 10,
                  ":before": {
                    backgroundColor: props.activeTabColor,
                    height: 3,
                  },
                  color: props.theme?.semanticColors?.inputText || "#000000",
                },
                link: {
                  fontSize: props.fontSize,
                  lineHeight: props.fontSize,
                  minHeight: 40,
                  height: "auto",
                  padding: "16px 8px",
                  width: `${100 / props.tabs.length}%`,
                  backgroundColor:
                    props.theme?.semanticColors?.buttonBackgroundDisabled ||
                    "#F3F2F1",
                  border: `1px solid ${
                    props.theme?.semanticColors?.disabledBodyText || "#F3F2F1"
                  }`,
                  borderBottom: "none",
                  borderTopLeftRadius: 10,
                  borderTopRightRadius: 10,
                  color:
                    props.theme?.semanticColors?.buttonTextDisabled ||
                    "#AEAEAE",
                },
              }}
            >
              {props.tabs.map((tab) => {
                return (
                  <PivotItem
                    key={tab.key}
                    headerText={tab.tabName}
                    itemKey={tab.tabId}
                  />
                );
              })}
            </Pivot>
            {props.displayMode === DisplayMode.Read && (
              <div
                className={styles.tabPanel}
                role="tabpanel"
                aria-hidden="false"
                aria-labelledby={_getTabId(selectedTab)}
              />
            )}
          </>
        )}
      </section>
    </AppContext.Provider>
  );
};
