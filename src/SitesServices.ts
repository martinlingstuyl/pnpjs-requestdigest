import "@pnp/sp/search";
import "@pnp/sp/webs";
import { ServiceKey, ServiceScope } from "@microsoft/sp-core-library";
import { PageContext } from "@microsoft/sp-page-context";
import { RequestDigest, spfi, SPFI, SPFx } from "@pnp/sp";
import { ISearchQuery, ISearchResult } from "@pnp/sp/search";
import { IWebInfo, Web } from "@pnp/sp/webs";

export interface ISitesService {
    get: (url: string) => Promise<IWebInfo>;
    search: (searchTerm: string, resultSourceId?: string) => Promise<ISearchResult[]>;
}

export class SitesService implements ISitesService {

    public static readonly serviceKey: ServiceKey<ISitesService> =
        ServiceKey.create<ISitesService>('SPFx:SitesService', SitesService);

    private _sp: SPFI;

    public constructor(serviceScope: ServiceScope) {
        serviceScope.whenFinished(() => {
            const pageContext = serviceScope.consume(PageContext.serviceKey);
            this._sp = spfi().using(SPFx({ pageContext }));

            // The problem does not occur if you specify the RequestDigest behavior as follows.        
            //this._sp = spfi().using(RequestDigest(), SPFx({ pageContext }));
        });
    }

    public async get(url: string): Promise<IWebInfo> {
        const web = Web([this._sp.web, url]);

        const webInfo = await web.select("Title")();

        return webInfo;
    }

    public async search(searchTerm: string): Promise<ISearchResult[]> {
        const searchQuery = {
            Querytext: `"${searchTerm}*" contentclass:STS_Site`,
            RowLimit: 10,
            SelectProperties: ["Title", "Path"]
        } as ISearchQuery;

        const results = await this._sp.search(searchQuery);

        return results.PrimarySearchResults;
    }
}