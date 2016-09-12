import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneSettings,
  IWebPartContext,
  PropertyPaneTextField
} from '@microsoft/sp-client-preview';

import * as strings from 'organisationChartStrings';
import OrganisationChart, { IOrganisationChartProps } from './components/OrganisationChart';
import { IOrganisationChartWebPartProps } from './IOrganisationChartWebPartProps';

import ModuleLoader from '@microsoft/sp-module-loader';

export default class OrganisationChartWebPart extends BaseClientSideWebPart<IOrganisationChartWebPartProps> {

  public constructor(context: IWebPartContext) {
    super(context);
  }

  public render(): void {
    //ModuleLoader.loadCss('https://appsforoffice.microsoft.com/fabric/2.2.0/fabric.min.css');
    //ModuleLoader.loadCss('https://appsforoffice.microsoft.com/fabric/2.2.0/fabric.components.min.css');

    const element: React.ReactElement<IOrganisationChartProps> = React.createElement(OrganisationChart, {
      description: this.properties.description
    });

    ReactDom.render(element, this.domElement);
  }

  protected get propertyPaneSettings(): IPropertyPaneSettings {
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
