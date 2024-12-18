import * as React from "react";
import {
  IProcessedStyleSet,
  mergeStyleSets,
} from "@fluentui/react/lib/Styling";

export const useMonacoEditorStyles = (): {
  controlClasses: IProcessedStyleSet<{
    containerStyles: {
      height: string;
      width: string;
    };
  }>;
} => {
  const controlClasses = React.useMemo(() => {
    return mergeStyleSets({
      containerStyles: {
        height: "70vh",
      },
    });
  }, []);
  return { controlClasses };
};
