import * as React from "react";
import { Provider } from "react-redux";
import { createStore, RootState } from "../state/store";
export interface IReduxProvider {
  initialState?: Partial<RootState>;
}
const ReduxProvider: React.FC<IReduxProvider> = ({
  children,
  initialState,
}) => {
  const store = React.useMemo(() => createStore(initialState), [initialState]);
  return <Provider store={store}>{children}</Provider>;
};

export default ReduxProvider;
