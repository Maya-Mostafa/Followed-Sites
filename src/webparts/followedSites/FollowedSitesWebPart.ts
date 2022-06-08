import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'FollowedSitesWebPartStrings';
import FollowedSites from './components/FollowedSites';
import { IFollowedSitesProps } from './components/IFollowedSitesProps';

export interface IFollowedSitesWebPartProps {
  wpTitle: string;
  editTxt: string;
  okTxt: string;
}

export default class FollowedSitesWebPart extends BaseClientSideWebPart<IFollowedSitesWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IFollowedSitesProps> = React.createElement(
      FollowedSites,
      {
        context: this.context,
        wpTitle: this.properties.wpTitle,
        editTxt: this.properties.editTxt,
        okTxt: this.properties.okTxt
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
                PropertyPaneTextField('wpTitle', {
                  label: 'Links Title',
                  value: this.properties.wpTitle,
                  description: 'e.g. Followed Sites'
                }),
                PropertyPaneTextField('editTxt', {
                  label: 'Edit Button Text',
                  value: this.properties.editTxt,
                }),
                PropertyPaneTextField('okTxt', {
                  label: 'Ok/Apply Button Text',
                  value: this.properties.okTxt,
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
