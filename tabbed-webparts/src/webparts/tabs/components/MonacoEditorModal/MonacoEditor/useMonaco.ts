import { useState, useEffect } from "react";
import loader, { Monaco } from "@monaco-editor/loader";

const CDN_PATH_TO_MONACO_EDITOR =
  "https://cdn.jsdelivr.net/npm/monaco-editor@0.47.0/min/vs";

export enum EStatus {
  LOADING,
  LOADED,
  ERROR,
}

export const useMonaco = (): {
  monaco: Monaco | undefined;
  status: EStatus | undefined;
  error: Error | undefined;
} => {
  const [monaco, setMonaco] = useState<Monaco>();
  const [status, setStatus] = useState<EStatus>(EStatus.LOADING);
  const [error, setError] = useState<Error>();

  useEffect(() => {
    (async () => {
      try {
        loader.config({ paths: { vs: CDN_PATH_TO_MONACO_EDITOR } });
        const monacoObj = await loader.init();
        setStatus(EStatus.LOADED);
        setMonaco(monacoObj);
      } catch (error) {
        setError(error);
        setStatus(EStatus.ERROR);
        setMonaco(undefined);
      }
    })()
      .then(() => {
        /** no-op; */
      })
      .catch(() => {
        /** no-op; */
      });
  }, []);

  return {
    monaco,
    status,
    error,
  };
};
