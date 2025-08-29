import { ServiceScope } from "@microsoft/sp-core-library";
import { GraphService } from "./GraphService";

export class ServiceManager {
  private static _graphService: GraphService;

  public static initialize(serviceScope: ServiceScope): void {
    if (!this._graphService) {
      this._graphService = new GraphService(serviceScope);
    }
  }

  public static get graphService(): GraphService {
    if (!this._graphService) {
      throw new Error(
        "ServiceManager is not initialized. Call initialize() first."
      );
    }
    return this._graphService;
  }
}
