import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';

import * as strings from 'InneficientWebPartStrings';
import Inneficient from './components/Inneficient';
import { IInneficientProps } from './components/IInneficientProps';

export interface IInneficientWebPartProps {
  urlApplications: string;
}

export default class InneficientWebPart extends BaseClientSideWebPart<IInneficientWebPartProps> {

  protected onInit(): Promise<void> {
    // De esta forma bloqueamos la carga del webpart
    return new Promise((resolve: any, reject: any) => {
      setTimeout(() =>{
        console.log("Timer");
        resolve(super.onInit()) ;
      }, 3000);
    });
    // De esta forma generamos una request asincrona y no bloqueamos la carga
    // new Promise((resolve: any, reject: any) => {
    //   setTimeout(() =>{
    //     console.log("Timer");
    //     resolve(super.onInit()); 
    //   }, 3000);
    // });
    // return super.onInit();
  }

  public render(): void {
    const element: React.ReactElement<IInneficientProps > = React.createElement(
      Inneficient,
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
