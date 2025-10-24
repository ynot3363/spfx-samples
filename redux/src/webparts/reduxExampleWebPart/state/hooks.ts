import type { TypedUseSelectorHook } from "react-redux";
import { useDispatch, useSelector, useStore } from "react-redux";
import type { AppStore, RootState } from "./store";
import { Dispatch, ThunkDispatch, UnknownAction } from "@reduxjs/toolkit";

type AppDispatch = ThunkDispatch<RootState, undefined, UnknownAction> &
  Dispatch<UnknownAction>;

// Typed versions of useDispatch and useSelector  - v8 method way of defining types
export const useAppDispatch: () => AppDispatch = useDispatch;
export const useAppSelector: TypedUseSelectorHook<RootState> = useSelector;
export const useAppStore: () => AppStore = () => useStore() as AppStore;
