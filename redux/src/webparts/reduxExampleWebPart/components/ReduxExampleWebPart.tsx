import * as React from "react";
import styles from "./ReduxExampleWebPart.module.scss";
import { useAppDispatch, useAppSelector } from "../state/hooks";
import {
  fetchUserDetails,
  getUser,
  getUserFetchingStatus,
  setUser,
} from "../state/features/userSlice";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { IDynamicPerson, PeoplePicker } from "@microsoft/mgt-react";

export interface IReduxExampleWebPartProps {
  hasTeamsContext: boolean;
}

export default function ReduxExampleWebPart(
  props: IReduxExampleWebPartProps
): JSX.Element {
  const { hasTeamsContext } = props;
  const dispatch = useAppDispatch();
  const fetchingUserDetails = useAppSelector(getUserFetchingStatus);
  const user = useAppSelector(getUser);
  const userKeys = Object.keys(user) as (keyof typeof user)[];
  const firstLoad = React.useRef(true);

  return (
    <section
      className={`${styles.reduxExampleWebPart} ${
        hasTeamsContext ? styles.teams : ""
      }`}
    >
      <div>
        <h2>Redux Example Web Part</h2>
        <PeoplePicker
          defaultSelectedUserIds={[user.id]}
          type="person"
          userType="user"
          selectionMode="single"
          selectionChanged={(e: CustomEvent<IDynamicPerson[]>) => {
            if (firstLoad.current) {
              firstLoad.current = false;
              return;
            }

            if (e.detail && e.detail.length > 0) {
              const person = e.detail[0];
              dispatch(
                setUser({
                  id: person.id ?? "",
                  name: person.displayName ?? "",
                  email: "",
                  givenName: "",
                  surname: "",
                  department: "",
                  jobTitle: "",
                  refreshing: false,
                })
              );
            }
          }}
        />
        <h4>Your user profile properties</h4>
        <ul>
          {userKeys.map((key) => (
            <li key={key}>
              <strong>{key}</strong>: {user[key].toString()}
            </li>
          ))}
        </ul>
        <PrimaryButton
          disabled={fetchingUserDetails}
          onClick={() => dispatch(fetchUserDetails())}
          text="Get user details"
        />
      </div>
    </section>
  );
}
