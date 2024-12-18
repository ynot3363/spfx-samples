import * as React from "react";
import styles from "./CustomModal.module.scss";
import { IDragOptions, Modal } from "@fluentui/react/lib/Modal";
import { DefaultButton, IconButton } from "@fluentui/react/lib/Button";

export interface ICustomModalProps {
  isOpen: boolean;
  isBlocking: boolean;
  dragOptions?: IDragOptions;
  onDismiss: () => void;
  children?: string | JSX.Element | JSX.Element[] | (() => JSX.Element);
}

export const CustomModal = function CustomModal(
  props: ICustomModalProps
): JSX.Element {
  const { isOpen, isBlocking, dragOptions, onDismiss, children } = props;

  return (
    <Modal
      isOpen={isOpen}
      isBlocking={isBlocking}
      onDismiss={onDismiss}
      styles={{ scrollableContent: { maxHeight: "90vh" } }}
      dragOptions={dragOptions || undefined}
    >
      <div className={styles.configureContainer}>{children}</div>
    </Modal>
  );
};

const header = (props: {
  title: string;
  onDismiss: () => void;
  children?: string | JSX.Element | JSX.Element[] | (() => JSX.Element);
}): JSX.Element => (
  <div className={styles.configureHeader}>
    <div className={styles.titleSection}>
      <h2 id="modalTitle">{props.title}</h2>
      <IconButton
        className={styles.dismissIconButton}
        iconProps={{ iconName: "Cancel" }}
        aria-label="Close popup modal"
        onClick={props.onDismiss}
      />
    </div>
    {props.children}
  </div>
);
CustomModal.Header = header;

const body = (props: {
  children?: string | JSX.Element | JSX.Element[] | (() => JSX.Element);
}): JSX.Element => <div className={styles.configureBody}>{props.children}</div>;
CustomModal.Body = body;

const footer = (props: {
  closeModalLabel: string;
  onDismiss: () => void;
  children?: string | JSX.Element | JSX.Element[] | (() => JSX.Element);
}): JSX.Element => (
  <div className={styles.configureFooter}>
    {props.children}
    <DefaultButton onClick={props.onDismiss}>
      {props.closeModalLabel}
    </DefaultButton>
  </div>
);
CustomModal.Footer = footer;
