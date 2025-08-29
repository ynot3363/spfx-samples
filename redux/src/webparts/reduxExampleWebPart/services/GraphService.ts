import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { MSGraphClientFactory } from "@microsoft/sp-http";
import { ResponseType } from "@microsoft/microsoft-graph-client";

export interface IGraphUserProfile {
  department: string;
  displayName: string;
  email: string;
  givenName: string;
  id: string;
  jobTitle: string;
  surname: string;
}

export class GraphService {
  public static readonly serviceKey: ServiceKey<GraphService> =
    ServiceKey.create<GraphService>("AaaS:GraphService", GraphService);

  private _msGraphClientFactory: MSGraphClientFactory;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._msGraphClientFactory = serviceScope.consume(
        MSGraphClientFactory.serviceKey
      );
    });
  }

  public async getUserProfile(userId: string): Promise<IGraphUserProfile> {
    const client = await this._msGraphClientFactory.getClient("3");
    try {
      const user = await client
        .api(`/users/${userId}`)
        .responseType(ResponseType.JSON)
        .select("displayName,department,id,givenName,jobTitle,mail,surname")
        .get();
      return {
        department: user.department || "",
        displayName: user.displayName || "",
        id: user.id || "",
        email: user.mail || "",
        givenName: user.givenName || "",
        jobTitle: user.jobTitle || "",
        surname: user.surname || "",
      };
    } catch (error) {
      console.error("Error fetching user profile:", error);
      return {
        department: "",
        displayName: "",
        id: "",
        email: "",
        givenName: "",
        jobTitle: "",
        surname: "",
      };
    }
  }
}
