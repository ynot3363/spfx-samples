import { RootState } from "../store";
import { createSlice } from "@reduxjs/toolkit";
import type { PayloadAction } from "@reduxjs/toolkit";
export interface WebPartState {
  instanceId: string;
  url: string;
}
const initialState: WebPartState = {
  instanceId: "",
  url: "",
};
export const getInstanceId = (state: RootState): string =>
  state.webpart.instanceId;
export const getWebPartUrl = (state: RootState): string => state.webpart.url;
export const webpartSlice = createSlice({
  name: "webpart",
  initialState,
  reducers: {
    setInstanceId: (state, action: PayloadAction<string>) => {
      state.instanceId = action.payload;
    },
  },
});
// Action creators are generated for each case reducer function
export const { setInstanceId } = webpartSlice.actions;
export default webpartSlice.reducer;
