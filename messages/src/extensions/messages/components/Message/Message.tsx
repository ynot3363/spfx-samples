import { IIconProps, MessageBar, MessageBarType } from "@fluentui/react";
import * as React from "react";
import { IMessage } from "../../../../models/IMessage";

export interface IMessageProps {
  /**
   * A collection of IMessage to render
   */
  messages: IMessage[];
}

const Message = ({ messages }: IMessageProps) => {
  const [_messages, setMessage] = React.useState(messages);
  const messageBarStyles = {
    root: { boxSizing: "border-box", padding: 10 },
    icon: { fontSize: 18 },
    text: { fontSize: 14 },
    dismissal: {
      flexContainer: { icon: { fontSize: 16, height: 16, lineHeight: 16 } },
    },
  };

  const dismissIconProps: IIconProps = {
    iconName: "Clear",
    styles: { root: { fontSize: "14px !important" } },
  };

  return (
    <div id="messageContainer">
      {_messages.map((message: IMessage) => {
        const linkUrl = new URL(message.link.url);
        const isSharePointLink = linkUrl.hostname.indexOf("sharepoint") !== -1;

        return (
          <MessageBar
            key={message.id}
            delayedRender={false}
            messageBarType={
              MessageBarType[message.type as keyof typeof MessageBarType]
            }
            isMultiline={true}
            dismissButtonAriaLabel="Close"
            dismissIconProps={dismissIconProps}
            onDismiss={() => {
              setMessage(_messages.filter((msg) => msg.id !== message.id));
            }}
            styles={messageBarStyles}
          >
            <strong>{message.message} </strong>
            <span>{message.details}</span>
            <a
              href={message.link.url}
              target={isSharePointLink ? "_self" : "_blank"}
              rel="noreferrer"
              data-interception={isSharePointLink ? "on" : "off"}
            >
              {message.link.desc}
            </a>
          </MessageBar>
        );
      })}
    </div>
  );
};

export default Message;
