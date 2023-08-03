import { Injectable } from '@nestjs/common';
import { Client } from '@microsoft/microsoft-graph-client';
import * as msRestNodeAuth from '@azure/ms-rest-nodeauth';

@Injectable()
export class SharePointService {
  private client: Client;

  constructor() {
    this.initialize();
  }

  async initialize() {
    const credentials = await msRestNodeAuth.loginWithServicePrincipalSecret(
      process.env.CLIENT_ID,
      process.env.CLIENT_SECRET,
      process.env.TENANT_ID,
    );

    this.client = Client.init({
      authProvider: (done) => {
        done(null, credentials.tokenCache._entries[0].accessToken);
      },
    });
  }

  async getAllSites(): Promise<any> {
    const sites = await this.client.api('/sites').get();
    return sites;
  }

  async getSearchSite(siteName: string): Promise<any> {
    const site = await this.client.api(`/sites/${siteName}`).get();
    return site;
  }
}
