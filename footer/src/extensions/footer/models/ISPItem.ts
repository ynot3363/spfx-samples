import { IODataListItem } from "@microsoft/sp-odata-types";

export interface ISPItem extends IODataListItem {
  link?: ISPUrl;
  icon?: ISPUrl;
  linkOrder?: number;
}

interface ISPUrl {
  Url: string;
  Description: string;
}
