import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'DocumentUploaderWebPartStrings';
import DocumentUploader from './components/DocumentUploader';
import { IDocumentUploaderProps } from './components/IDocumentUploaderProps';

export interface IDocumentUploaderWebPartProps {
  description: string;
}

export default class DocumentUploaderWebPart extends BaseClientSideWebPart<IDocumentUploaderWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDocumentUploaderProps > = React.createElement(
      DocumentUploader,
      {
        description: this.properties.description,
        context: this.context
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
            description: ''
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Document Library Name'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
