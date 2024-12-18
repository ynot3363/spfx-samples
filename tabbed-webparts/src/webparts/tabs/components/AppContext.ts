import { createContext } from "react";
import { ITabsProps } from "./Tabs";

export const AppContext = createContext<ITabsProps>({
  tabs: [],
  activeTabColor: "#041e42",
  displayMode: 1,
  fontSize: "16px",
  themeBased: true,
  domElement: undefined,
  theme: undefined,
  updateProperty: () => {
    return null;
  },
});
