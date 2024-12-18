import * as React from "react";
import { Text } from "@fluentui/react/lib/Text";

export interface IDescriptionProps {
  /** The description text */
  text: string;
}

export const Description: React.FC<IDescriptionProps> = function Description({
  text,
}: IDescriptionProps): React.ReactElement {
  return (
    <Text variant="small" style={{ color: "#949494" }}>
      {text}
    </Text>
  );
};
