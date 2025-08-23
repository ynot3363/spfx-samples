import { BaseApplicationCustomizer } from "@microsoft/sp-application-base";
import { Log } from "@microsoft/sp-core-library";
import * as strings from "AnalyticsApplicationCustomizerStrings";

declare global {
  interface Window {
    clarity: {
      (...args: unknown[]): void;
      q?: unknown[][];
    };
    gtag: (...args: unknown[]) => void;
    dataLayer?: unknown[];
    [key: string]: boolean;
  }
}

const LOG_SOURCE: string = "AnalyticsApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IAnalyticsApplicationCustomizerProperties {}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class AnalyticsApplicationCustomizer extends BaseApplicationCustomizer<IAnalyticsApplicationCustomizerProperties> {
  private _clarityTrackingId: string = "[YOUR CLARITY CODE HERE]"; // Replace with your Clarity ID
  private _clarityScriptId: string = "clarity-tracking-script";
  private _gaTrackingID: string = "[YOUR G-CODE HERE]"; // Replace with your Google Analytics ID
  private _gaScriptId: string = "ga-tracking-script";
  private _calledOnce: boolean = false;

  public async onInit(): Promise<void> {
    await super.onInit();
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    if (!this._calledOnce) {
      // Listen for placeholder placement changes
      this.context.placeholderProvider.changedEvent.add(
        this,
        this.applyTrackingCode
      );
      /* This event is triggered when user navigate between the pages */
      this.context.application.navigatedEvent.add(this, this.applyTrackingCode);
      this._calledOnce = true;
    }
    return Promise.resolve();
  }

  private async applyTrackingCode(): Promise<void> {
    const someTestCondition = true; // Replace with your actual condition

    if (someTestCondition) {
      this._addClarityAnalytics();
      this._addGoogleAnalytics();
    } else {
      this._removeClarityAnalytics();
      this._removeGoogleAnalytics();
    }
  }

  private _addGoogleAnalytics(): void {
    if (!document.getElementById(this._gaScriptId)) {
      //First we need to define the dataLayer and gtag function globally
      window.dataLayer = window.dataLayer || [];
      window.gtag = function (): void {
        // eslint-disable-next-line prefer-rest-params
        window.dataLayer!.push(arguments); // If you change this from arguments to ...args, it breaks
      };

      //Create and load GA script
      const gaScript = document.createElement("script");
      gaScript.id = this._gaScriptId;
      gaScript.type = "text/javascript";
      gaScript.async = true;
      gaScript.src = `https://www.googletagmanager.com/gtag/js?id=${this._gaTrackingID}`;

      gaScript.onload = () => {
        window.gtag("js", new Date());
        window.gtag("config", `${this._gaTrackingID}`);
        Log.info(LOG_SOURCE, "Google Analytics initialized");
      };

      document.head.appendChild(gaScript);
      Log.info(LOG_SOURCE, "Google Analytics script loaded");
    } else {
      Log.info(LOG_SOURCE, "Google Analytics script already loaded");
    }
  }

  private _removeGoogleAnalytics(): void {
    const gaScript = document.getElementById(this._gaScriptId);
    const gaDisableID = `ga-disable-${this._gaTrackingID}`;
    window[gaDisableID] = true; // Disable tracking
    if (gaScript) {
      gaScript.remove();
      Log.info(LOG_SOURCE, "Google Analytics script removed");
    } else {
      Log.info(LOG_SOURCE, "Google Analytics script not found");
    }
  }

  private _addClarityAnalytics(): void {
    if (!document.getElementById(this._clarityScriptId)) {
      // First define the clarity function if it doesn't exist
      window.clarity =
        window.clarity ||
        function (...args: unknown[]): void {
          (window.clarity.q = window.clarity.q || []).push(args);
        };

      // Create and load Clarity script
      const clarityScript = document.createElement("script");
      clarityScript.id = this._clarityScriptId;
      clarityScript.async = true;
      clarityScript.src = `https://www.clarity.ms/tag/${this._clarityTrackingId}`;
      document.head.appendChild(clarityScript);
      Log.info(LOG_SOURCE, "Clarity tracking script loaded");
    } else {
      Log.info(LOG_SOURCE, "Clarity tracking script already loaded");
    }
  }

  private _removeClarityAnalytics(): void {
    const clarityScript = document.getElementById(this._clarityScriptId);
    if (clarityScript) {
      window.clarity("stop");
      clarityScript.remove();
      Log.info(LOG_SOURCE, "Clarity tracking script removed");
    } else {
      Log.info(LOG_SOURCE, "Clarity tracking script not found");
    }
  }
}
