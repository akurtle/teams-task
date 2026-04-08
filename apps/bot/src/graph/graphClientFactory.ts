import {
  Client,
  GraphRequest
} from "@microsoft/microsoft-graph-client";
import {
  AuthenticationResult,
  ConfidentialClientApplication
} from "@azure/msal-node";
import { env } from "../config/env";

type EnvShape = typeof env;

export class GraphClientFactory {
  private readonly msalClient: ConfidentialClientApplication;

  public constructor(private readonly settings: EnvShape) {
    this.msalClient = new ConfidentialClientApplication({
      auth: {
        clientId: settings.azureAdClientId,
        clientSecret: settings.azureAdClientSecret,
        authority: `https://login.microsoftonline.com/${settings.azureAdTenantId}`
      }
    });
  }

  public async createClient(userAssertion?: string) {
    const token = await this.getToken(userAssertion);

    return Client.init({
      authProvider: (done) => done(null, token.accessToken)
    });
  }

  public request(client: Client, path: string): GraphRequest {
    return client.api(path).version("v1.0");
  }

  private async getToken(userAssertion?: string): Promise<AuthenticationResult> {
    if (userAssertion) {
      return this.acquireOnBehalfOfToken(userAssertion);
    }

    const result = await this.msalClient.acquireTokenByClientCredential({
      scopes: this.settings.graphScopes
    });

    if (!result) {
      throw new Error("Unable to acquire Microsoft Graph app token.");
    }

    return result;
  }

  private async acquireOnBehalfOfToken(userAssertion: string) {
    const result = await this.msalClient.acquireTokenOnBehalfOf({
      oboAssertion: userAssertion,
      scopes: this.settings.graphScopes,
      skipCache: false
    });

    if (!result) {
      throw new Error("Unable to acquire Microsoft Graph delegated token.");
    }

    return result;
  }
}
