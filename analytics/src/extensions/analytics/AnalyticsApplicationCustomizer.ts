import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";

declare global {
  interface Window {
    clarity: (action: string) => void;

    gtag: (...args: unknown[]) => void;

    dataLayer?: unknown[];
  }
}

import * as strings from "AnalyticsApplicationCustomizerStrings";

const LOG_SOURCE: string = "AnalyticsApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAnalyticsApplicationCustomizerProperties {}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AnalyticsApplicationCustomizer extends BaseApplicationCustomizer<IAnalyticsApplicationCustomizerProperties> {
  private clarityId: string = "{YOUR_CLARITY_ID}"; // Replace with your Clarity ID
  private clarityScriptId: string = "clarity-tracking-script";
  private gaID: string = "{YOUR_GA_ID}"; // Replace with your Google Analytics ID
  private gaScriptId: string = "ga-tracking-script";

  public async onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    // Listen for navigation changes
    this.context.placeholderProvider.changedEvent.add(
      this,
      this.applyTrackingCode
    );
    /* This event is triggered when user navigate between the pages */
    this.context.application.navigatedEvent.add(this, this.applyTrackingCode);
    return Promise.resolve();
  }

  private async applyTrackingCode(): Promise<void> {
    this.loadClarity();
    this.loadGA();
  }

  private loadGA(): void {
    if (!document.getElementById(this.gaScriptId)) {
      const script = document.createElement("script");
      script.id = this.gaScriptId;
      script.async = true;
      script.src = `https://www.googletagmanager.com/gtag/js?id=G-${this.gaID}`;
      script.onload = () => {
        // Safe to initialize here
        window.dataLayer = window.dataLayer || [];
        function gtag(...args: unknown[]): void {
          window.dataLayer?.push(args);
        }
        gtag("js", new Date());
        gtag("config", `G-${this.gaID}`, {
          siteName: this.context.pageContext.web.title,
        });
      };

      // The commented-out code below is an alternative way to load Google Analytics, however
      // because it creates the script element dynamically and passes the function as a string
      // to the text property it will trigger a CSP violation in SPO now.
      /**
      const gaScript = document.createElement("script");
      gaScript.id = this.gaScriptId;
      gaScript.type = "text/javascript";
      gaScript.async = true;
      gaScript.src = `https://www.googletagmanager.com/gtag/js?id=G-${this.gaID}`;

      const gaScript2 = document.createElement("script");
      gaScript2.id = this.gaScriptId2;
      gaScript2.type = "text/javascript";
      gaScript2.text = `window.dataLayer = window.dataLayer || [];function gtag(){dataLayer.push(arguments);}gtag('js', new Date()); gtag('config', 'G-${
        this.gaID
      }',{'siteName': '${JSON.stringify(
        this.context.pageContext.web.title
      )}'});`;

      document.head.appendChild(gaScript);
      document.head.appendChild(gaScript2);
      */

      document.head.appendChild(script);
      Log.info(LOG_SOURCE, "Google Analytics script loaded");
    } else {
      Log.info(LOG_SOURCE, "Google Analytics script already loaded");
    }
  }

  private loadClarity(): void {
    if (!document.getElementById(this.clarityScriptId)) {
      const script = document.createElement("script");
      script.id = this.clarityScriptId;
      script.async = true;
      script.src = `https://www.clarity.ms/tag/${this.clarityId}`;
      document.head.appendChild(script);
      Log.info(LOG_SOURCE, "Clarity tracking script loaded");
    } else {
      Log.info(LOG_SOURCE, "Clarity tracking script already loaded");
    }
  }
}
