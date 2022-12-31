export interface IMessage {
  /**
   * The main message
   */
  message: string;
  /**
   * The details around the message
   */
  details: string;
  /**
   * A url that points to where the link should take the user
   */
  link: IUrl;
  /**
   * Defines the type of message to render
   */
  type: string;
  /**
   * Defines the publish date of when the message should show for users
   */
  publishDate: string;
  /**
   * Defines the expiration date of when the message should show for users
   */
  expirationDate: string;
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
