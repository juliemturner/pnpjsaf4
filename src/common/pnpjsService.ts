// <reference path="../../node_modules/@pnp/sp/fi.d.ts" />

import * as AppInsights from 'applicationinsights';
import { DefaultAzureCredential } from "@azure/identity";
import { MyItem } from './models';

export interface IPnpjsService {
  Init: () => Promise<boolean>;
  GetListItem: (id: string) => Promise<MyItem>;
}

export class PnpjsService implements IPnpjsService {
  private LOG_SOURCE = "PnpjsService";
  private _ready: boolean = false;

  private _sp = null;

  public constructor() { }

  public async Init(): Promise<boolean> {
    let retVal = false;
    try {
      const { spfi, SPFI } = await import("@pnp/sp");
      const { AzureIdentity } = await import("@pnp/azidjsclient/index.js");
      const { SPDefault } = await import("@pnp/nodejs/index.js");
      await import("@pnp/sp/webs/index.js");
      await import("@pnp/sp/lists/index.js");
      await import("@pnp/sp/items/index.js");

      const credential = new DefaultAzureCredential();
      
      this._sp = spfi(process.env.SiteUrl).using(SPDefault({}),
        AzureIdentity(credential, [`https://${process.env.Tenant}.sharepoint.com/.default`], null));

      this._ready = true;
      retVal = true;
    } catch (err) {
      AppInsights.defaultClient.trackException({
        exception: err,
        severity: AppInsights.Contracts.SeverityLevel.Critical,
        properties: { source: this.LOG_SOURCE, method: "Init" }
      });
    }
    return retVal;
  }

  public get ready(): boolean {
    return this._ready;
  }

  public async GetListItem(id: string): Promise<MyItem> {
    let retVal: MyItem = null;
    try {

    } catch (err) {
      AppInsights.defaultClient.trackException({
        exception: err,
        severity: AppInsights.Contracts.SeverityLevel.Critical,
        properties: { source: this.LOG_SOURCE, method: "GetListItem" }
      });
    }
    return retVal;
  }
}

export const pnpjs: IPnpjsService = new PnpjsService();