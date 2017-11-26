import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { IItemResponse } from './IPnPService';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IProsjektAvviksItems, IProsjektInformasjon } from '../components/IProsjektAvvikProps';
import {Web} from 'sp-pnp-js';

export default class PnPService {

  private _results: any[];
  private _site: string = "https://jadarhus.sharepoint.com/sites/avvik";
  private _listname: string = "Avviksliste";

  constructor(private _context: IWebPartContext) {
  }

  /**
   * retrieve list items by the specified query
   *
   * @param query
   * @param sorting
   * @param fields
   * @param orderBy
   * @param expand
   */
  public get(query: string, sorting: string, fields: string, expand: string, maxResults: number) {
    return new Promise<IItemResponse>(( resolve, reject) => {
      this._getItems( fields, sorting, query, maxResults, expand )
      .then( ( res: any ) => {
        if (typeof res["odata.error"] !== "undefined") {
          if (typeof res["odata.error"]["message"] !== "undefined") {
              reject(res["odata.error"]["message"].value);
              return;
          }
        }

        let resultsRetrieved = false;

        if (!this._isNull(res)) {
          resultsRetrieved = true;
        }

        const response: IItemResponse = {
          results: res,
          totalResults: res.length
        };

        resolve(response);
      }).catch((error: string) => reject(error));
    });

  }

  private _getItems( select: string, sort: string, filter: string, top: number, expand: string ): Promise<IProsjektAvviksItems[]> {
    let web = new Web(this._site);
    return web.lists.getByTitle(this._listname).items
      .select(select).top(top).orderBy(sort, false).filter(filter).expand(expand)
      .get().then( r  => {
        return r;
    }).catch(error => {
      return Promise.reject(JSON.stringify(error));
    });
  }

  public _getProjectInfo(listname: string): Promise<IProsjektInformasjon> {
    let thisSite: string = this._context.pageContext.site.absoluteUrl;
    let web = new Web(thisSite);
    return web.lists.getByTitle(listname).items.orderBy("Created",false).top(1).get().then( r => {
      if (r.length > 0) {
        return r[0];
      } else {
        return r;
      }
    }).catch(error => {
      return Promise.reject(JSON.stringify(error));
    });
  }

  private _isEmptyString(value: string): boolean {
    return value === null || typeof value === "undefined" || !value.length;
  }

  private _isNull(value: any): boolean {
    return value === null || typeof value === "undefined";
  }
}
