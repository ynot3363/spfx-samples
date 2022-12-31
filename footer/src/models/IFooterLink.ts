export interface IFooterLink {
  /**
   * The name of the link
   */
  name: string;
  /**
   * A url that points to where the link should take the user
   */
  link: IUrl;
  /**
   * A url that points to the icon to load for the footer
   */
  icon: IUrl;
  /**
   * the order of the link in the list
   */
  order: number;
  /**
   * the id of the link in the associated list
   */
  id: number;
}

interface IUrl {
  /**
   * The url the link navigates to
   */
  url: string;
  /**
   * The description of the url
   */
  desc: string;
}
