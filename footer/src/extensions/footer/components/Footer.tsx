import * as React from "react";
import styles from "./Footer.module.scss";
import { IFooterLink } from "../models/IFooterLink";

export interface IFooterProps {
  /**
   * A collection of IFooterLink to render
   */
  footerLinks: IFooterLink[];
  /**
   * Whether to show the copyright message in the footer
   */
  showCopyright: boolean;
  /**
   * The company name to show in the copyright message
   */
  companyName: string;
}

declare global {
  interface Window {
    __themeState__: {
      theme: {
        primaryButtonBackground: string;
        primaryButtonText: string;
      };
    };
  }
}

const Footer = ({ companyName, footerLinks, showCopyright }: IFooterProps) => {
  const backgroundColor =
    window.__themeState__.theme.primaryButtonBackground || "#041e42";
  const linkTextColor =
    window.__themeState__.theme.primaryButtonText || "#ffffff";

  return (
    <div
      style={{ backgroundColor: backgroundColor }}
      className={styles.footerContainer}
    >
      <div className={styles.stack}>
        <div className={styles.stackItem} style={{ flex: 1 }}>
          <img
            src={require("../assets/FooterLogo.png")}
            alt="Footer Logo"
            title="Footer Logo"
            width="300px"
          />
        </div>
        {footerLinks.map((link) => {
          const linkUrl = new URL(link.link.url);
          const isSharePointLink =
            linkUrl.hostname.indexOf("sharepoint") !== -1;

          return (
            <div key={link.id} className={styles.stackItem}>
              <a
                href={linkUrl.href}
                target={isSharePointLink ? "_self" : "_blank"}
                rel={"noreferrer"}
                data-interception={isSharePointLink ? "on" : "off"}
                className={styles.link}
                style={{
                  color: linkTextColor,
                }}
              >
                {!!link.icon && (
                  <img
                    aria-label={link.icon.desc}
                    role="img"
                    src={link.icon.url}
                    title={link.icon.desc}
                  />
                )}
                <div>{link.name}</div>
              </a>
            </div>
          );
        })}
      </div>
      {showCopyright && (
        <div className={styles.copyright}>
          <span style={{ color: linkTextColor }}>
            © {new Date().getFullYear()} {companyName} All rights reserved.{" "}
          </span>
        </div>
      )}
    </div>
  );
};

export default Footer;
