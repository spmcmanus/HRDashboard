import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import Dashdemo from './components/Dashdemo';
import { IDashdemoProps } from './components/IDashdemoProps';
import { IDashdemoWebPartProps } from './IDashdemoWebPartProps';

export default class DashdemoWebPart extends BaseClientSideWebPart<IDashdemoWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IDashdemoProps > = React.createElement(
      Dashdemo,
      {
        title: this.properties.title,
        defaultTag: this.properties.defaultTag
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('title', {
                  label: 'Title',
                  value: 'HR Dashboard'
                }),
                PropertyPaneTextField('defaultTag', {
                  label: 'Default Tag Name',
                  value: 'HRHome'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
