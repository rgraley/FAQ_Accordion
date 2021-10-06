import { WebPartContext } from '@microsoft/sp-webpart-base';  
import { DisplayMode } from '@microsoft/sp-core-library';  

export interface ISearchResultsAccordionProps {
  description: string;  
  displayMode: DisplayMode; 
  updateProperty: (value: string) => void; 
}
