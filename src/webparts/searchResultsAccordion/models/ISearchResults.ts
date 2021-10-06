export interface ISearchResults{
    value: ISearchResult[];
}
export interface ISearchResult {
    ListItemID: number;
    Path: string;
    NormListID: string;
    SiteTitle: string;
}
export interface ISearchService{
    GetSearchResults(query:string, sourceId:string, filterText:string, numberOfItems:string) : Promise<ISearchResult[]>;
}

