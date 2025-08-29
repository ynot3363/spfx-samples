import { RootState } from "../store";
import {
  createAsyncThunk,
  createListenerMiddleware,
  createSlice,
} from "@reduxjs/toolkit";
import type { PayloadAction } from "@reduxjs/toolkit";
import { ServiceManager } from "../../services/ServiceManager";

export interface UserState {
  id: string;
  name: string;
  email: string;
  department: string;
  givenName: string;
  jobTitle: string;
  surname: string;
  refreshing: boolean;
}

export const initialState: UserState = {
  id: "",
  name: "",
  email: "",
  department: "",
  givenName: "",
  jobTitle: "",
  surname: "",
  refreshing: false,
};

export const getUser = (state: RootState): UserState => state.user;
export const getUserFetchingStatus = (state: RootState): boolean =>
  state.user.refreshing;

export const fetchUserDetails = createAsyncThunk(
  "user/fetchDetails",
  async (_, thunkAPI) => {
    const state = thunkAPI.getState() as RootState;
    const userId = state.user.id;
    const userDetails = await ServiceManager.graphService.getUserProfile(
      userId
    );
    return userDetails;
  }
);

export const userSlice = createSlice({
  name: "user",
  initialState,
  reducers: {
    setUser: (state, action: PayloadAction<UserState>) => {
      return { ...state, ...action.payload };
    },
  },
  extraReducers: (builder) => {
    builder.addCase(fetchUserDetails.pending, (state, action) => {
      state.refreshing = true;
      return;
    });
    builder.addCase(fetchUserDetails.fulfilled, (state, action) => {
      const {
        department,
        displayName,
        email,
        givenName,
        id,
        jobTitle,
        surname,
      } = action.payload;
      state.id = id;
      state.department = department;
      state.name = displayName;
      state.email = email;
      state.givenName = givenName;
      state.jobTitle = jobTitle;
      state.surname = surname;
      state.refreshing = false;
      return;
    });
  },
});

// Action creators are generated for each case reducer function
export const { setUser } = userSlice.actions;

export const userListener = createListenerMiddleware();
userListener.startListening({
  actionCreator: setUser,
  effect: async (action, listenerApi) => {
    // Dispatch fetchUserDetails after setUser
    await listenerApi.dispatch(fetchUserDetails());
  },
});

export default userSlice.reducer;
