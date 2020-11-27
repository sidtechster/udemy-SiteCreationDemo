import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SiteCreationWpWebPart.module.scss';
import * as strings from 'SiteCreationWpWebPartStrings';

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

export interface ISiteCreationWpWebPartProps {
  description: string;
}

export default class SiteCreationWpWebPart extends BaseClientSideWebPart<ISiteCreationWpWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.siteCreationWp }">
        
        <h1>Create a new subsite</h1>
        <p>Please fill the below details to create a new subsite.</p><br/>
        Subsite title: <br/><input type="text" id="txtSubsiteTitle"/><br/>
        Subsite URL: <br/><input type="text" id="txtSubsiteUrl"/><br/>
        Subsite description: <br/><textarea id="txtSubsiteDesc" rows="5" cols="30"></textarea><br/>

        <input type="button" id="btnCreateSubsite" value="Create Subsite"/><br/>

      </div>`;

      this.bindEvents();
  }

  private bindEvents(): void {
    this.domElement.querySelector('#btnCreateSubsite').addEventListener('click', () => { this.createSubsite(); });
  }

  private createSubsite(): void {

    let subsiteTitle = document.getElementById('txtSubsiteTitle')["value"];
    let subsiteUrl = document.getElementById('txtSubsiteUrl')["value"];
    let subsiteDesc = document.getElementById('txtSubsiteDesc')["value"];

    const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/webinfos/add";

    const spHttpClientOptions: ISPHttpClientOptions = {
      body: `{
        "parameters":{
          "@odata.type": "SP.WebInfoCreationInformation",
          "Title": "${subsiteTitle}",
          "Url": "${subsiteUrl}",
          "Description": "${subsiteDesc}",
          "Language": 1033,
          "WebTemplate": "STS#0",
          "UseUniquePermissions": true
        }
      }`
    };

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if(response.status === 200) {
          alert("New subsite created");
        }
        else {
          alert("Error " + response.status + " - " + response.statusText);
        }
      });

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
