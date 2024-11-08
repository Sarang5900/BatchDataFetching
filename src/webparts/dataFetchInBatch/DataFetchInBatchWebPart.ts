import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DataFetchInBatchWebPartStrings';
import DataFetchInBatch from './components/DataFetchInBatch';
import { IDataFetchInBatchProps } from './components/IDataFetchInBatchProps';

export interface IDataFetchInBatchWebPartProps {
  description: string;
}

export default class DataFetchInBatchWebPart extends BaseClientSideWebPart<IDataFetchInBatchWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDataFetchInBatchProps> = React.createElement(
      DataFetchInBatch,
      {
        description: this.properties.description,
        context: this.context,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
