/* eslint-disable @typescript-eslint/no-explicit-any */
import * as React from "react";
import { useMonacoEditorStyles } from "./useMonacoEditorStyles";
import { EStatus, useMonaco } from "./useMonaco";
import { Shimmer } from "@fluentui/react/lib/Shimmer";
import { ThemeProvider } from "@fluentui/react/lib/Theme";
import { mergeStyles } from "@fluentui/react/lib/Styling";
import { MessageBar, MessageBarType } from "@fluentui/react/lib/MessageBar";

export interface IMonacoEditorProps {
  jsonString: any;
  wrapText: boolean;
  readOnly: boolean;
  onChange: (newValue: string) => void;
}

const MonacoEditor: React.FC<IMonacoEditorProps> = function MonacoEditor(
  props: IMonacoEditorProps
) {
  const containerRef = React.useRef<HTMLDivElement>(null);
  const editorRef = React.useRef<any>(null);
  const { controlClasses } = useMonacoEditorStyles();
  const { monaco, status, error } = useMonaco();
  const { jsonString, readOnly, wrapText, onChange } = props;

  const onDidChangeModelContent = React.useCallback((e: any): void => {
    if (editorRef.current) {
      const currentValue: string = editorRef.current.getValue();
      if (currentValue !== jsonString) {
        const validationErrors: string[] = [];
        try {
          onChange(currentValue);
        } catch (e) {
          validationErrors.push(e.message);
        }
      }
    }
  }, []);

  React.useEffect(() => {
    if (status !== EStatus.LOADED || !containerRef.current || !monaco) return;

    monaco.editor.onDidCreateModel((m: any) => {
      m.updateOptions({
        tabSize: 2,
      });
    });

    editorRef.current = monaco.editor.create(containerRef.current, {
      value: jsonString,
      scrollBeyondLastLine: false,
      theme: "vs",
      language: "json",
      folding: true,
      readOnly: readOnly,
      lineNumbersMinChars: 4,
      lineNumbers: "on",
      formatOnPaste: true,
      minimap: {
        enabled: false,
      },
      wordWrap: wrapText ? "on" : "off",
    });

    editorRef.current.onDidChangeModelContent(onDidChangeModelContent);

    return () => {
      editorRef?.current.dispose();
    };
  }, [monaco]);

  React.useEffect(() => {
    if (status !== EStatus.LOADED) return;

    editorRef.current.updateOptions({ wordWrap: wrapText ? "on" : "off" });
  }, [wrapText]);

  if (status === EStatus.LOADING) {
    const wrapperClass = mergeStyles({
      padding: 2,
      width: "50vw",
      selectors: {
        "& > .ms-Shimmer-container": {
          margin: "10px 0px",
        },
      },
    });
    return (
      <ThemeProvider className={wrapperClass}>
        <Shimmer />
        <Shimmer width="75%" />
        <Shimmer width="50%" />
      </ThemeProvider>
    );
  }

  if (status === EStatus.ERROR) {
    return (
      <MessageBar isMultiline messageBarType={MessageBarType.error}>
        {error?.message ||
          "An error occured trying to fetch the monaco editor code."}
      </MessageBar>
    );
  }

  return (
    <>
      <div ref={containerRef} className={controlClasses.containerStyles} />
    </>
  );
};

export default MonacoEditor;
