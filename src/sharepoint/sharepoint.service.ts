import { Injectable } from '@nestjs/common';
import * as msRestNodeAuth from '@azure/ms-rest-nodeauth';
import { Client } from '@microsoft/microsoft-graph-client';
import axios from 'axios';

@Injectable()
export class SharePointService {
  private accessToken: string;
  private client: Client;

  constructor() {
    this.initialize();
  }

  async initialize() {
    const credentials = await msRestNodeAuth.loginWithServicePrincipalSecret(
      process.env.CLIENT_ID,
      process.env.CLIENT_SECRET,
      process.env.TENANT_ID,
      {
        tokenAudience: 'https://graph.microsoft.com',
      },
    );

    const tokenResponse = await credentials.getToken();
    this.accessToken = tokenResponse.accessToken;

    this.client = Client.init({
      authProvider: (done) => {
        done(null, this.accessToken);
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

  async search(query: string): Promise<any> {
    const response = await axios.get(
      `https://graph.microsoft.com/v1.0/sites?search=${query}`,
      {
        headers: {
          Authorization: `Bearer ${this.accessToken}`,
        },
      },
    );

    return response.data;
  }
}
