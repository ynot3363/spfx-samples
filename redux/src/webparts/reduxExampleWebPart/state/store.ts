import {
  configureStore,
  Dispatch,
  ThunkDispatch,
  UnknownAction,
} from "@reduxjs/toolkit";
import userReducer, { userListener, UserState } from "./features/userSlice";
import webpartReducer, { WebPartState } from "./features/webpartSlice";

// Define the reducer mapping
const rootReducer = {
  user: userReducer,
  webpart: webpartReducer,
};

// Define state and dispatch types based on the temp store
export type RootState = {
  user: UserState;
  webpart: WebPartState;
};

export type AppDispatch = ThunkDispatch<RootState, undefined, UnknownAction> &
  Dispatch<UnknownAction>;

// Function to create the actual store with preloaded state
// eslint-disable-next-line @typescript-eslint/explicit-function-return-type
export const createStore = (preloadedState?: Partial<RootState>) => {
  return configureStore({
    reducer: rootReducer,
    middleware: (getDefaultMiddleware) =>
      getDefaultMiddleware().prepend(userListener.middleware),
    preloadedState,
  });
};

// Store type
export type AppStore = ReturnType<typeof createStore>;
