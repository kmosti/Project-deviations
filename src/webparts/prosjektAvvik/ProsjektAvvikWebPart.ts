import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';

import * as strings from 'ProsjektAvvikWebPartStrings';
import ProsjektAvvik from './components/ProsjektAvvik';
import { IProsjektAvvikProps } from './components/IProsjektAvvikProps';
import { IprosjektAvvikWebPartProps } from './IprosjektAvvikWebPartProps';

export default class ProsjektAvvikWebPart extends BaseClientSideWebPart<IprosjektAvvikWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IProsjektAvvikProps > = React.createElement(
      ProsjektAvvik,
      {
        title: this.properties.title,
        powerapplink: this.properties.powerapplink,
        maxResults: this.properties.maxResults,
        context: this.context
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
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "Innstillinger",
              groupFields: [
                PropertyPaneTextField('title', {
                  label: "Tittel"
                }),
                PropertyPaneTextField('powerapplink', {
                  label: "Link til powerapp",
                  multiline: true
                }),
                PropertyPaneSlider('maxResults', {
                  label: "Max resultater",
                  min: 1,
                  max: 100
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
