declare interface ITabsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  FontSizeFieldLabel: string;
  FontSizeFieldDescription: string;
  ThemeToggleLabel: string;
  ThemeToggleOnText: string;
  ThemeToggleOffText: string;
  ThemeToggleDescription: string;
  ActiveTabColorLabel: string;
  ActiveTabColorDescription: string;
  AdvancedGroupName: string;
}

declare module "TabsWebPartStrings" {
  const strings: ITabsWebPartStrings;
  export = strings;
}
