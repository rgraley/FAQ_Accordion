import * as React from 'react';
import styles from '../components/SearchResultsAccordionWebPart.module.scss';
import { ISearchResultsAccordionProps } from './ISearchResultsAccordionProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";

export default class SearchResultsAccordion extends React.Component<ISearchResultsAccordionProps, {}> {
  public render(): React.ReactElement<ISearchResultsAccordionProps> {
    return (
      <div >
          <WebPartTitle displayMode={this.props.displayMode} title={this.props.description} updateProperty={this.props.updateProperty} placeholder="Web Part Title" />
          <div className={ styles.container }>
            <div className={ styles.row }>
              <div className={ styles.column }>
                <div className="accordion">
                </div>                
              </div>
            </div>
          </div>
      </div>
    );
  }
}
