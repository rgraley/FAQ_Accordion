import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version,
  Environment,
  EnvironmentType  } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import SearchResultsAccordion from './components/SearchresultsAccordion';
import { ISearchResultsAccordionProps } from './components/ISearchResultsAccordionProps';
import {ISearchResultsAccordionWebPartProps} from './models/ISearchResultsAccordionWebPartProps';
import * as isearchResult from './models/ISearchResults';
import * as iQALists from './models/IQALists';
import MockHttpClient from './components/MockHttpClient';
import { SortDirection } from 'sp-pnp-js';
import * as sr from './models/ISearchResults';
import * as esr from './models/IExtendedSearchResultSearhResults';
import { sp } from '@pnp/sp';
import * as jQuery from 'jquery';
import 'jqueryui';
//import styles from './components/SearchResultsAccordionWebPart.module.scss';
import {
  SPHttpClient,
} from '@microsoft/sp-http';

import {SPComponentLoader} from "@microsoft/sp-loader";

//the following variables retrieve settings from the webpart properties panel
let propKeywordQuery: string;
let propSourceId: string;
let propFilterText: string ;
let propFilterTextToggle: boolean = false;
let propNumberOfItems: string;

export default class SearchResultsAccordionWebPart extends BaseClientSideWebPart<ISearchResultsAccordionWebPartProps> {
  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  public constructor(){
    super();
    SPComponentLoader.loadCss('https://maxcdn.bootstrapcdn.com/font-awesome/4.6.3/css/font-awesome.min.css');
    SPComponentLoader.loadCss('//code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css');
    require('./AccordionOverrides.css');

  }

  private _isEmpty(value): boolean {
    return value === undefined ||
      value === null ||
      value.length === 0;
  }

  private _getWebPartPropertyValues(){
      //the following variables retrieve settings from the webpart properties panel
      propKeywordQuery = this._isEmpty(this.properties.searchQuery) ? '*' : this.properties.searchQuery;
      //Personal SharePoint Online
      //let propSourceId: string = this._isEmpty(this.properties.sourceId) ? '4225801d-e615-4925-84d8-4f3ec2c22d6a' : this.properties.sourceId;
      //BayCare SharePoint Online
      propSourceId = this._isEmpty(this.properties.sourceId) ? 'bd420bb0-b7b2-4b9d-8031-3757ad3c95ce' : this.properties.sourceId;
      propFilterText = this.properties.filterText;
      propFilterTextToggle = false;
      propNumberOfItems= this._isEmpty(this.properties.numberOfItems) ? "20" : this.properties.numberOfItems;
      if (this.properties.filterTextToggle){
        propFilterTextToggle = true;
      }
      if(!this._isEmpty(propFilterText)){
        //Refiners were provided so parse them into a refinementFilter string using the _buildRefinementFilters function
        let tempQueryFilterText = this._buildRefinementFilters(propFilterText, propFilterTextToggle);
        propFilterText = tempQueryFilterText;
      }else{
        //No refiners were provided so use the default astrisk to return all items
        propFilterText = '*';
      }
  }

  private ButtonClick(oldVal: any): any {   
    this._getWebPartPropertyValues();
    this._renderListAsync(propKeywordQuery,propSourceId,propFilterText,propNumberOfItems);  
    return '' ;    
  } 

  private isNullOrEmpty( s ) 
  {
      return ( s == null || s === "" );
  }
  
  private _buildRefinementFilters(filterText:string, propFilterTextCondition:boolean){
    //instantiate ret variable as a string and set to empty string
    let returnString: string ="";
    //check to see if the filterText value is already set to '*' asterisk.  If so then keep this value; otherwise look to parse the values
    //Remove any trailing semi-colons from the string
    filterText = filterText.replace(/;+$/, "").trim();
    //Remove any astrisks from the string
    filterText = filterText.replace(/\*/g, "").trim();    
    if(filterText.indexOf(';') > -1){ //there may be one or more semi colons      
      let filterArray = filterText.split(';');
      if(filterArray.length <= 0){
          //something went wrong
          returnString = "Something Is Not Right.";
      }else if(filterArray.length == 1) {
          //Only one item found
          returnString = 'bcTopics:("' + filterArray[0] + '*")';
      }else{
          //More than one item found
          let strCondition: string = propFilterTextCondition==true ? 'and' : 'or';

          returnString = 'bcTopics:' + strCondition + '(';
          for(let i=0; i < filterArray.length; i++){
            if(!this.isNullOrEmpty(filterArray[i].trim())){
              returnString += '"' + filterArray[i].trim() + '*",';
            }
          }
          //Remove any trailing commas from the string
          returnString = returnString.replace(/,\s*$/, "");  
          let countCommas = (returnString.match(/,/g) || []).length;
          if(countCommas <= 0){
            //remove the condition from the string
            returnString = propFilterTextCondition==true ? returnString.replace("bcTopics:and(", "bcTopics:(") : returnString.replace("bcTopics:or(", "bcTopics:("); 
          }
          returnString += ')';        
      }
    }else{
      //refinementfilters=bcTopics:("High Five Rewards*") contains
      returnString = 'bcTopics:("' + filterText + '*")';
    }
    return returnString;
  }

  private _renderList(items: iQALists.QAList[]): void {
    jQuery('.accordion', this.domElement).empty();
    items.forEach((item: iQALists.QAList) => {
      let newDiv: string = '';
      newDiv += `      
      <h3 class="ui-accordion-header"><span class="ms-ListItem-primaryText">${item.Title}</span></h3>
      <div>        
        <span class="ms-ListItem-tertiaryText">${item.bcAnswer}</span>
        <span class="ms-ListItem-tertiaryText" style="font-size: 12px; display:block;padding-top:10px;">${item.SiteTitle}</span>
        <span class="ms-ListItem-tertiaryText" style="font-size: 12px; display:block;"><a href="${item.SiteUrl.toLowerCase()}">${item.SiteUrl.toLowerCase()}</a></span>
      </div>`;
      jQuery('.accordion', this.domElement).append(newDiv);
    });
    jQuery('.accordion', this.domElement).accordion("refresh");
  }

  private _renderListAsync(queryQuery:string, querySourceId:string, queryFilterText:string, numberOfItems:string): void {
    var qaResults = [];

    // Local environment
    if (Environment.type === EnvironmentType.Local) {
      this._getSearchResultsData('MockSearchService','','','','')
      .then((response) => {
        this._getMockListData().then((qaResponse) => {
          this._renderList(qaResponse.value);
        });    
      });
    } //Hosted Environment
    else if (
      Environment.type == EnvironmentType.SharePoint || Environment.type == EnvironmentType.ClassicSharePoint) {
      this._getSearchResultsData('SearchService',queryQuery,querySourceId,queryFilterText,numberOfItems)
        .then((response) => { 
          this._getSearchResultsListData(response).then((qaResponse) => {
            this._renderList(qaResponse.value);
          });
        });
      }
  }

  private async _getSearchResultsListData(searchResults): Promise<iQALists.QALists> {
    let searchResultItems: isearchResult.ISearchResult[] = searchResults.value;
    var _items: iQALists.QAList[]=[];
    for(let index=0; index < searchResultItems.length; index++){
      //Get the webUrl of the list item
      let SiteTitle: string = String(searchResultItems[0].SiteTitle)
      let itemPath: string = String(searchResultItems[index].Path);
      let itemPathIndex: number = itemPath.indexOf("/Lists/") ;
      let webUrl: string = itemPath.substring(0, itemPathIndex);
      //Get the List Guid and the Item Id
      let itemListGuid: string = searchResultItems[index].NormListID; 
      let itemId: string = String(searchResultItems[index].ListItemID);
      //build the full path to each list item
      let strFullPath: string = webUrl  + `/_api/web/lists('` + itemListGuid + `')/items(`+ itemId +`)`;
      let dataItem: iQALists.QAList = await this._getSharePointListData(strFullPath);
       if(!this._isEmpty(dataItem.Title) && !this._isEmpty(dataItem.bcAnswer)){
        _items.push({ Title: dataItem.Title, bcAnswer: dataItem.bcAnswer,SiteTitle:SiteTitle,SiteUrl:webUrl });
       }
    }
    return new Promise<iQALists.QAList[]>((resolve) => {
      resolve(_items);
    }).then((data: iQALists.QAList[]) => {
      var listData: iQALists.QALists = { value: data };
      return listData;
    }) as Promise<iQALists.QALists>;
  }
  /*
    This method queries the sharepoint FAQ Lists for the Question and the Answers
  */
  private async _getSharePointListData(strFullPath): Promise<iQALists.QAList> {
    try{
      let awaitResponse = await this.context.spHttpClient.get(strFullPath, SPHttpClient.configurations.v1);
      return awaitResponse.json() as Promise<iQALists.QAList>;
    }catch (error) {
      console.log(error);
    }
  }
  /*
    This method queries the sharepoint search for all matching results
  */
  private _getSearchResultsData(datasource:string, queryQuery:string,querySourceId:string,queryFilterText:string,numberOfItems:string): Promise<isearchResult.ISearchResults> {
    let _search:isearchResult.ISearchService = null;
    datasource=="MockSearchService" ? _search  = new MockSearchService() : _search = new SearchService();
    //datasource=="MockSearchService" ? _search  = new SearchService.MockSearchService() : _search = new SearchService.SearchService();
    return  _search.GetSearchResults(queryQuery, querySourceId,queryFilterText,numberOfItems)
    .then((data: isearchResult.ISearchResult[]) => {
      var listData: isearchResult.ISearchResults = { value: data };
      return listData;
    }) as Promise<isearchResult.ISearchResults>;
  }
  public render(): void {
    const element: React.ReactElement<ISearchResultsAccordionProps > = React.createElement(
      SearchResultsAccordion,
      {  
        description: this.properties.description,  
        displayMode: this.displayMode,
        updateProperty: (value: string) => {  
          this.properties.description = value;  
        }  
      }
    );

    this._getWebPartPropertyValues();
    this._renderListAsync(propKeywordQuery,propSourceId,propFilterText,propNumberOfItems); 

    ReactDom.render(element, this.domElement);
    const accordionOptions: JQueryUI.AccordionOptions = {
      animate: true,
      active: false,
      collapsible: true,
      icons: {
        header: 'ui-icon-circle-arrow-e',
        activeHeader: 'ui-icon-circle-arrow-s',
      },
      heightStyle: "content",
    };
    jQuery('.accordion', this.domElement).accordion(accordionOptions);
    jQuery('.accordion', this.domElement).accordion("refresh");
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: "Search query configuration"
          },
          groups: [
            {
              //groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Web Part Title',
                }),
                PropertyPaneTextField('searchQuery', {
                  label: 'Search query keywords',
                  multiline: true,
                  placeholder: '',
                  description: 'Example: Title: "Why am I taxed"; bcAnswer:"Recognition".',
                }),
                PropertyPaneTextField('sourceId',{
                  label: 'Result Source ID',
                  multiline: false,
                  placeholder: 'bd420bb0-b7b2-4b9d-8031-3757ad3c95ce',
                  description: 'This value is set in the SharePoint Search Center for the needed Result Source.',
                }),
                PropertyPaneTextField('filterText',{
                  label: 'Filter by Topics',
                  multiline: true,
                  description: 'Seperate each topic with a semi-colon.',
                }),
                PropertyPaneToggle('filterTextToggle',
                {
                  label: 'Filter by Topics Condition',
                  key: 'filterTextToggle',
                  onText: 'Must contain all listed terms.',
                  offText: 'May contain any listed terms',
                }),
                PropertyPaneTextField('numberOfItems',{
                  label: 'Number of items per page',
                  multiline: false,
                  placeholder: '20',
                  description: 'Default value is 20',
                }),
                PropertyPaneButton('webpartButton',  
                 {  
                  text: 'Apply', 
                  buttonType: PropertyPaneButtonType.Primary,  
                  onClick: this.ButtonClick.bind(this)  
                 })
              ]
            }
          ]
        }
      ]
    };
  }
  /*
    The following method is used to test the webpart on the local workbench only
  */
  private _getMockListData(): Promise<iQALists.QALists> {
    return MockHttpClient.get()
      .then((data: iQALists.QAList[]) => {
        var listData: iQALists.QALists = { value: data };
        return listData;
      }) as Promise<iQALists.QALists>;
  }
}
export class SearchService implements sr.ISearchService
{
    private _isEmpty(value: string): boolean {
        return value === undefined ||
          value === null ||
          value.length === 0;
      }

    public GetSearchResults(query:string, sourceId:string, filterText:string,numberOfItems:string) : Promise<sr.ISearchResult[]>{
        let numOfItems: number = Number(numberOfItems);
        const _results:sr.ISearchResult[] = [];
        sourceId = this._isEmpty(sourceId) ? undefined : sourceId;
        return new Promise<sr.ISearchResult[]>((resolve) => {
            sp.search({
                  Querytext: query,
                  RowLimit:numOfItems,
                  StartRow:0,
                  SourceId: sourceId, 
                  SelectProperties: ['Title','bcTopics','ListItemID','Path','NormListID','bcAnswer','SiteTitle'],//Title,Path,Created,Filename,SiteLogo,PreviewUrl,PictureThumbnailURL,ServerRedirectedPreviewURL,ServerRedirectedURL,HitHighlightedSummary,FileType,contentclass,ServerRedirectedEmbedURL,ParentLink,DefaultEncodingURL,owstaxidmetadataalltagsinfo,Author,AuthorOWSUSER,SPSiteUrl,SiteTitle,IsContainer,IsListItem,HtmlFileType,SiteId,WebId,UniqueID,OriginalPath,FileExtension,NormSiteID,NormListID,NormUniqueID
                  RefinementFilters:[filterText],
                  SortList: [{Property:'LastModifiedTime', Direction:SortDirection.Descending}],
                  TrimDuplicates: true,
                }).then((results) => {  
                results.PrimarySearchResults.forEach((result:esr.ExtendedSearchResult)=>{
                _results.push({
                  ListItemID:result.ListItemID,
                  Path:result.Path,
                  NormListID: result.NormListID,
                  SiteTitle: result.SiteTitle,
                });
                });
            })
            .then(
                () => { resolve(_results);}
            )
            .catch((error) =>{
              console.log("ProcessHttpClientResponseException Error occurred: ", error);
            });                  
        });
    }
}
export class MockSearchService implements sr.ISearchService
{
    public GetSearchResults(query:string, sourceId: string, filterText:string) : Promise<sr.ISearchResult[]>{
        return new Promise<sr.ISearchResult[]>((resolve) => {
          
      resolve([
            {ListItemID:4,Path:'https://<weburlhere>/Lists/FAQs/DispForm.aspx?ID=4',NormListID:'123',SiteTitle:'Temp SiteTitle'},
            {ListItemID:5,Path:'https://<weburlhere>/Lists/FAQs/DispForm.aspx?ID=5',NormListID:'345',SiteTitle: 'Temp SiteTitle'},
            ]);
        });
    }
}