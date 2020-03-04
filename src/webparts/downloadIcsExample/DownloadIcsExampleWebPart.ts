import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'DownloadIcsExampleWebPartStrings';
import DownloadIcsExample from './components/DownloadIcsExample';
import { IDownloadIcsExampleProps } from './components/DownloadIcsExample.Types';

export interface IDownloadIcsExampleWebPartProps {
  description: string;
}

export default class DownloadIcsExampleWebPart extends BaseClientSideWebPart <IDownloadIcsExampleWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDownloadIcsExampleProps> = React.createElement(
      DownloadIcsExample,
      {
        description: this.properties.description
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
