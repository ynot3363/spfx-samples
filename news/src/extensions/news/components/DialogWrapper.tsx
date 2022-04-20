import * as React from "react";
import { Dialog, DialogFooter, DialogType } from "@fluentui/react/lib/Dialog";
import { PrimaryButton } from "@fluentui/react/lib/Button";

const modalPropsStyles = { main: { maxWidth: 450 } };

export interface IDialogWrapperProps {
  title: string;
  subText: string;
  showDialog: boolean;
  closeDialog: () => void;
}

export const DialogWrapper: React.FunctionComponent = ({
  title,
  subText,
  showDialog,
  closeDialog,
}: IDialogWrapperProps) => {
  const modalProps = React.useMemo(
    () => ({
      isBlocking: true,
      styles: modalPropsStyles,
    }),
    []
  );

  const dialogContentProps = {
    type: DialogType.normal,
    title: title,
    subText: subText,
  };

  return (
    <>
      <Dialog
        hidden={!showDialog}
        onDismiss={() => {
          closeDialog();
        }}
        dialogContentProps={dialogContentProps}
        modalProps={modalProps}
      >
        <DialogFooter>
          <PrimaryButton
            onClick={() => {
              closeDialog();
            }}
            text="OK"
          />
        </DialogFooter>
      </Dialog>
    </>
  );
};
