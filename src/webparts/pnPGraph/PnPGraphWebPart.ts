import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { SPHttpClient } from '@microsoft/sp-http';
import * as strings from 'PnPGraphWebPartStrings';
import PnPGraph from './components/PnPGraph';
import { IPnPGraphProps } from './components/IPnPGraphProps';

import { setup as pnpSetup } from '@pnp/common';

export interface IPnPGraphWebPartProps {
  description: string;
}

export default class PnPGraphWebPart extends BaseClientSideWebPart<IPnPGraphWebPartProps> {
  
  public onInit(): Promise<void> {
    pnpSetup({
      spfxContext: this.context
    });
    return Promise.resolve();
  }

  public async render(): Promise<void> {
    let myurl= this.context.pageContext.site.absoluteUrl +"/_api/SP.UserProfiles.PeopleManager/getmyproperties";
     
    let response = await this.context.spHttpClient.get(myurl, SPHttpClient.configurations.v1);
    if (!(response.ok)) {throw new Error(await response.text()); }

    const responseProps: any = await response.json();
    let mynumber = responseProps.UserProfileProperties.filter((item: { Key: string; }) => item.Key ==  "WorkPhone").map((item: { Value: any; }) => item.Value)[0];
    let firstname =  responseProps.UserProfileProperties.filter((item: { Key: string; }) => item.Key ==  "FirstName").map((item: { Value: any; }) => item.Value)[0];

    const element: React.ReactElement<IPnPGraphProps> = React.createElement(
      PnPGraph,
      {
        description: this.properties.description,
        context: this.context,
        siteurl: this.context.pageContext.web.absoluteUrl,
        myhttp:this.context.spHttpClient,
        mysite:this.context.pageContext.site.absoluteUrl,
        me:  {email:this.context.pageContext.user.email,displayname:this.context.pageContext.user.displayName,phone:mynumber,firstname:firstname}
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
