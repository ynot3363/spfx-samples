import { IODataListItem } from "@microsoft/sp-odata-types";

export interface ISPItem extends IODataListItem {
  msg_details: string;
  msg_link: ISPUrl;
  msg_type: string;
  msg_publishDate: string;
  msg_expirationDate: string;
}

interface ISPUrl {
  Url: string;
  Description: string;
}
