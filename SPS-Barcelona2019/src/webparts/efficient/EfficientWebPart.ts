import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'EfficientWebPartStrings';
import Efficient from './components/Efficient';
import { IEfficientProps } from './components/IEfficientProps';

export interface IEfficientWebPartProps {
  urlApplications: string;
}

export default class EfficientWebPart extends BaseClientSideWebPart<IEfficientWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IEfficientProps> = React.createElement(
      Efficient,
      {
        urlApplications: this.properties.urlApplications
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private validateEmpty(val: string): string {
    if (val == "") {
      return strings.EmptyUrlMessage;
    }
    return "";
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
          groups: [
            {
              groupName: "Origen de datos",
              groupFields: [
                PropertyPaneTextField('urlApplications', {
                  label: strings.DataSourceLabel,
                  onGetErrorMessage: this.validateEmpty,
                  deferredValidationTime: 1000
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
