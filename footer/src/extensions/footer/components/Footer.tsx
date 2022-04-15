import * as React from "react";
import style from "./Footer.module.scss";
import { IFooterLink } from "../../../models/IFooterLink";
import styles from "./Footer.module.scss";

export interface IFooterProps {
  /**
   * A collection of IFooterLink to render
   */
  footerLinks: IFooterLink[];
}

declare global {
  interface Window {
    __themeState__: any;
  }
}

const Footer = ({ footerLinks }: IFooterProps) => {
  /**
   * Set the background color of the container based on the site theme otherwise default to blue
   */
  const backgroundColor =
    window.__themeState__.theme.suiteBarBackground || "#0086bd";

  /**
   * Set the link text color based on the site theme otherwise default to white
   */
  const linkTextColor = window.__themeState__.theme.suiteBarText || "#ffffff";

  return (
    <div
      id="customFooter"
      style={{ backgroundColor: backgroundColor }}
      className={styles.footerContainer}
    >
      <div className={styles.stack}>
        {footerLinks.map((link) => {
          const linkUrl = new URL(link.link.url);
          const isSharePointLink =
            linkUrl.hostname.indexOf("sharepoint") !== -1;

          return (
            <div key={link.id} order={link.order} className={styles.stackItem}>
              <a
                href={linkUrl.href}
                target={isSharePointLink ? "_self" : "_blank"}
                data-interception={isSharePointLink ? "on" : "off"}
                className={styles.link}
                style={{
                  color: linkTextColor,
                }}
              >
                <img role="icon" src={link.icon.url} title={link.name} />
                <div>{link.name}</div>
              </a>
            </div>
          );
        })}
      </div>
    </div>
  );
};

export default Footer;
