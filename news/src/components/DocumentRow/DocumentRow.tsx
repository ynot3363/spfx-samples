import * as React from "react";
import styles from "./DocumentRow.module.scss";
import {
  ListViewCommandSetContext,
  RowAccessor,
} from "@microsoft/sp-listview-extensibility";
import { Icon } from "@fluentui/react";
import { Text } from "@fluentui/react/lib/Text";
import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { Log } from "@microsoft/sp-core-library";
import { IListItemFormUpdateValue } from "../../models/IListItemFormUpdateValue";

export interface IDocumentRowProps {
  context: ListViewCommandSetContext;
  item: RowAccessor;
  formState: string;
  command: string;
  date?: Date;
  updateCallsMade: () => void;
}

interface IDocumentRowState {
  documentState: string;
  errorMessage: string;
}

const LOG_SOURCE: string = "NewsCommandSet:DocumentRow";

export class DocumentRow extends React.Component<
  IDocumentRowProps,
  IDocumentRowState
> {
  constructor(props: IDocumentRowProps) {
    super(props);
    this.state = {
      documentState: "Pending",
      errorMessage: null,
    };

    this._updateDocument.bind(this);
  }

  public componentDidUpdate(prevProps: Readonly<IDocumentRowProps>): void {
    if (
      this.props.formState !== prevProps.formState &&
      this.props.formState === "Submitting"
    ) {
      if (this.state.documentState === "Pending") {
        this.setState(
          {
            documentState: "Submitting",
          },
          this._updateDocument
        );
      }
    }
  }

  public render(): React.ReactElement<IDocumentRowProps> {
    const { item } = this.props;
    const { documentState, errorMessage } = this.state;
    return (
      <>
        <div key={item.getValueByName("ID")} className={styles.documentRow}>
          <Icon
            className={`${styles.icon} ${
              documentState === "Submitting" ? styles.refreshStart : ""
            } ${documentState === "Successful" ? styles.successful : ""} ${
              documentState === "Failed" ? styles.failed : ""
            }
          `}
            iconName={
              documentState === "Successful"
                ? "SkypeCircleCheck"
                : documentState === "Failed"
                ? "ErrorBadge"
                : documentState === "Submitting"
                ? "Refresh"
                : "PageList"
            }
          />
          <Text
            title={item.getValueByName("FileLeafRef")}
            className={styles.fileName}
            variant="mediumPlus"
          >
            {item.getValueByName("FileLeafRef")}
          </Text>
        </div>
        {errorMessage && (
          <div className={styles.errorMessage}>
            <Text variant="smallPlus">
              <b>Error Message: </b>
              {errorMessage}
            </Text>
          </div>
        )}
        <div className={styles.horizontalRule} />
      </>
    );
  }

  private async _updateDocument(): Promise<void> {
    const { command, item, updateCallsMade } = this.props;

    let hasException = false;
    let fieldErrorMessage = "";

    //TODO Look into refactoring the code to reduce redundancy
    switch (command) {
      case "DEMOTE_NEWS_POST":
        try {
          const demotePageCall: SPHttpClientResponse =
            await this._demoteNewsPost(item);
          const demotePageJSON = await demotePageCall.json();
          if (demotePageCall.ok) {
            //Check to make sure there are no errors passed back in the call per field
            demotePageJSON.value.forEach((field: IListItemFormUpdateValue) => {
              if (field.HasException) {
                hasException = true;
                fieldErrorMessage = field.ErrorMessage;
              }
            });

            if (hasException) {
              this.setState({
                documentState: "Failed",
                errorMessage: fieldErrorMessage,
              });

              updateCallsMade();
              return;
            }

            const publishPageCall: SPHttpClientResponse =
              await this._publishPage(item);

            if (publishPageCall.ok) {
              this.setState({
                documentState: "Successful",
              });
              updateCallsMade();
              return;
            } else {
              this.setState({
                documentState: "Failed",
                errorMessage: `The ${item.getValueByName(
                  "FileLeafRef"
                )} page was successfully demoted but an error occurred publishing the page. Please publish the page manually.`,
              });

              updateCallsMade();
              return;
            }
          } else {
            const errorMessage =
              demotePageJSON.error && demotePageJSON.error.message
                ? demotePageJSON.error.message
                : demotePageCall.statusText;
            if (demotePageCall.statusText === "Locked") {
              this.setState({
                documentState: "Failed",
                errorMessage: `${errorMessage} Please reach out to the user.`,
              });
            } else {
              this.setState({
                documentState: "Failed",
                errorMessage: `${errorMessage} Please refresh the page and try again. If the error continues please reach out to IT.`,
              });
            }

            updateCallsMade();
            return;
          }
        } catch (error) {
          Log.error(LOG_SOURCE, error);
          this.setState({
            documentState: "Failed",
            errorMessage:
              "An error occurred trying to update the page. Please refresh the page and try again. If the error continues please reach out to IT",
          });

          updateCallsMade();
        }
        break;
      case "PROMOTE_NEWS_POST":
        try {
          const promotePageCall: SPHttpClientResponse =
            await this._promoteNewsPost(item);
          const promotePageJSON = await promotePageCall.json();
          if (promotePageCall.ok) {
            promotePageJSON.value.forEach((field: IListItemFormUpdateValue) => {
              //Check to make sure there are no errors passed back in the call per field
              if (field.HasException) {
                hasException = true;
                fieldErrorMessage = field.ErrorMessage;
              }
            });

            if (hasException) {
              this.setState({
                documentState: "Failed",
                errorMessage: fieldErrorMessage,
              });

              updateCallsMade();
              return;
            }

            const publishPageCall: SPHttpClientResponse =
              await this._publishPage(item);

            if (publishPageCall.ok) {
              this.setState({
                documentState: "Successful",
              });
              updateCallsMade();
              return;
            } else {
              this.setState({
                documentState: "Failed",
                errorMessage: `The ${item.getValueByName(
                  "FileLeafRef"
                )} page was successfully promoted but an error occurred publishing the page. Please publish the page manually.`,
              });

              updateCallsMade();
              return;
            }
          } else {
            const errorMessage =
              promotePageJSON.error && promotePageJSON.error.message
                ? promotePageJSON.error.message
                : promotePageCall.statusText;
            if (promotePageCall.statusText === "Locked") {
              this.setState({
                documentState: "Failed",
                errorMessage: `${errorMessage} Please reach out to the user.`,
              });
            } else {
              this.setState({
                documentState: "Failed",
                errorMessage: `${errorMessage} Please refresh the page and try again. If the error continues please reach out to IT.`,
              });
            }

            updateCallsMade();
            return;
          }
        } catch (error) {
          Log.error(LOG_SOURCE, error);
          this.setState({
            documentState: "Failed",
            errorMessage:
              "An error occurred trying to update the page. Please refresh the page and try again. If the error continues please reach out to IT.",
          });

          updateCallsMade();
        }
        break;
      case "UPDATE_PUBLISHING_DATE":
        try {
          const updatePublishDatePageCall: SPHttpClientResponse =
            await this._updatePublishDate(item);
          const updatePublishDateJSON = await updatePublishDatePageCall.json();
          if (updatePublishDatePageCall.ok) {
            updatePublishDateJSON.value.forEach(
              (field: IListItemFormUpdateValue) => {
                //Check to make sure there are no errors passed back in the call per field
                if (field.HasException) {
                  hasException = true;
                  fieldErrorMessage = field.ErrorMessage;
                }
              }
            );

            if (hasException) {
              this.setState({
                documentState: "Failed",
                errorMessage: fieldErrorMessage,
              });

              updateCallsMade();
              return;
            }

            const publishPageCall: SPHttpClientResponse =
              await this._publishPage(item);

            if (publishPageCall.ok) {
              this.setState({
                documentState: "Successful",
              });
              updateCallsMade();
              return;
            } else {
              this.setState({
                documentState: "Failed",
                errorMessage: `The ${item.getValueByName(
                  "FileLeafRef"
                )} page's publish date was successfully updated but an error occurred publishing the page. Please publish the page manually.`,
              });

              updateCallsMade();
              return;
            }
          } else {
            const errorMessage =
              updatePublishDateJSON.error && updatePublishDateJSON.error.message
                ? updatePublishDateJSON.error.message
                : updatePublishDatePageCall.statusText;
            if (updatePublishDatePageCall.statusText === "Locked") {
              this.setState({
                documentState: "Failed",
                errorMessage: `${errorMessage}. Please reach out to the user.`,
              });
            } else {
              this.setState({
                documentState: "Failed",
                errorMessage: `${errorMessage} Please refresh the page and try again. If the error continues please reach out to IT.`,
              });
            }

            updateCallsMade();
            return;
          }
        } catch (error) {
          Log.error(LOG_SOURCE, error);
          this.setState({
            documentState: "Failed",
            errorMessage:
              "An error occurred trying to update the page. Please refresh the page and try again. If the error continues please reach out to IT.",
          });

          updateCallsMade();
        }
        break;
      default:
        throw new Error("Unknown command");
    }
  }

  /**
   * Method will update the page and set the promoted state to 0
   * @param item The item in the row that is currently being looped through
   * @returns response from the ValidateUpdateListItem() call
   */
  private async _demoteNewsPost(
    item: RowAccessor
  ): Promise<SPHttpClientResponse> {
    const { context } = this.props;
    const url = `${context.pageContext.site.absoluteUrl}/_api/web/lists(guid'${
      context.pageContext.list.id
    }')/items(${item.getValueByName("ID")})/ValidateUpdateListItem()`;
    const body = {
      formValues: [
        {
          FieldName: "PromotedState",
          FieldValue: "0",
        },
      ],
      bNewDocumentUpdate: false,
    };

    return await context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "Content-Type": "application/json;odata=nometadata",
          "Content-Length": JSON.stringify(body).length.toString(),
        },
        body: JSON.stringify(body),
      }
    );
  }

  /**
   * Method will update the page and set the promoted state to 2
   * @param item The item in the row that is currently being looped through
   * @returns response from the ValidateUpdateListItem() call
   */
  private async _promoteNewsPost(
    item: RowAccessor
  ): Promise<SPHttpClientResponse> {
    const { context } = this.props;
    const url = `${context.pageContext.site.absoluteUrl}/_api/web/lists(guid'${
      context.pageContext.list.id
    }')/items(${item.getValueByName("ID")})/ValidateUpdateListItem()`;
    const body = {
      formValues: [
        {
          FieldName: "PromotedState",
          FieldValue: "2",
        },
      ],
      bNewDocumentUpdate: false,
    };

    return await context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "Content-Type": "application/json;odata=nometadata",
          "Content-Length": JSON.stringify(body).length.toString(),
        },
        body: JSON.stringify(body),
      }
    );
  }

  /**
   * MEthod will update the page's publish date to the date the user had selected
   * @param item The item in the row that is currently being looped through
   */
  private async _updatePublishDate(item: RowAccessor) {
    const { context, date } = this.props;
    if (date) {
      //Format the date in the way SharePoint expects it MM/DD/YY HH:MM A. Update based on the locale settings of the site.
      const formattedDate = `${date.toLocaleDateString([
        "en-US",
      ])} ${date.toLocaleString(["en-US"], {
        hour: "2-digit",
        minute: "2-digit",
        hour12: true,
      })}`;
      const url = `${
        context.pageContext.site.absoluteUrl
      }/_api/web/lists(guid'${
        context.pageContext.list.id
      }')/items(${item.getValueByName("ID")})/ValidateUpdateListItem()`;
      const body = {
        formValues: [
          {
            FieldName: "FirstPublishedDate",
            FieldValue: formattedDate,
          },
        ],
        bNewDocumentUpdate: false,
      };

      return await context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            "Content-Type": "application/json;odata=nometadata",
            "Content-Length": JSON.stringify(body).length.toString(),
          },
          body: JSON.stringify(body),
        }
      );
    }

    throw new Error("No date provided");
  }

  /**
   * Method will publish a page based on the server relative url of the item
   * @param item The item in the row that is currently being looped through
   * @returns response from the Publish() call
   */
  private async _publishPage(item: RowAccessor): Promise<SPHttpClientResponse> {
    const { context } = this.props;
    const url = `${
      context.pageContext.site.absoluteUrl
    }/_api/web/GetFileByServerRelativeUrl('${item.getValueByName(
      "FileRef"
    )}')/Publish()`;

    const body = { comment: "Demoted News Post" };

    return await context.spHttpClient.post(
      url,
      SPHttpClient.configurations.v1,
      {
        headers: {
          "Content-Type": "application/json;odata=nometadata",
        },
        body: JSON.stringify(body),
      }
    );
  }
}
