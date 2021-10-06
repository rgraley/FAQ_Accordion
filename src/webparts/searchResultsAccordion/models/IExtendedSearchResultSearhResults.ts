import * as pnp from 'sp-pnp-js';

export interface ExtendedSearchResult extends pnp.SearchResult{
    bcTopics: string; //bcAnswer managed property
    ListItemID:number;
    Path:string;
    NormListID: string;
    SiteTitle: string;
}