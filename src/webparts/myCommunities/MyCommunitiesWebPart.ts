import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneSlider,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './MyCommunitiesWebPart.module.scss';
import * as strings from 'MyCommunitiesWebPartStrings';

import {HttpClient } from '@microsoft/sp-http';

import {app} from '@microsoft/teams-js'; 

export interface IMyCommunitiesWebPartProps {
  description: string;
  numberOfBlocks: number;
  seeAllButton: string;
}

export default class MyCommunitiesWebPart extends BaseClientSideWebPart<IMyCommunitiesWebPartProps> {
  private communityInfo: any[] = []; // Array to store community information

  private excludedTitles: string[] = ['Buy & Sell', 'Carpooling', 'Accommodation','OneRPG' /* add more titles as needed */];

  private isTeams = false;
  private isEmbedded = false;
  protected async onInit(): Promise<void> {
    try {
      await app.initialize();
      const context = await app.getContext();
      console.log("Context:", context);
      if(context.app.host.name.includes("teams") || context.app.host.name.includes("Teams")){
        console.log("The extension is running inside Microsoft Teams");
        this.isTeams = true;
      }else{
        console.log("The extension is running outside Microsoft Teams");
      }
    } catch (exp) {
        console.log("The extension is running outside Microsoft Teams");
  }
  this.isEmbedded = document.body.classList.contains('embedded');
  if (this.isEmbedded) {
    console.log('Body has the embedded class');
  } else {
    console.log('Body does not have the embedded class');
  }
    await this.getCommunityInfo();
}



public render(): void {
  const decodedSeeAllButton = decodeURIComponent(this.properties.seeAllButton);
  console.log("Url for See All button: ", decodedSeeAllButton);
  this.domElement.innerHTML = `
    <div>
      <div class="${styles.topSection}">
        <div class="${styles.MyCommunitiesHeading}">My Communities</div>
        <div><a href="${decodedSeeAllButton}" target="_blank">See all</a></div>
      </div>
      <section class="${styles.MyCommunities}">
        ${this.communityInfo.slice(0, this.properties.numberOfBlocks).map((group: any, index: number) => {
          console.log("Group WebUrl:", group.webUrl);
          let fullLink = "";
          if (group.webUrl && group.webUrl.includes("feedId=")) {
            const feedId = group.webUrl.split("feedId=")[1];
            console.log("Extracted FeedId:", feedId);
            
            // Construct the string
            const jsonString = `{"_type":"Group","id":"${feedId}"}`;
            console.log("Constructed JSON string:", jsonString);

            // Encode the string in base64
            const encodedString = btoa(jsonString);
            console.log("Encoded String:", encodedString);
            
            const teamsLink = `https://teams.microsoft.com/l/entity/db5e5970-212f-477f-a3fc-2227dc7782bf/vivaengage?context=%7B%22subEntityId%22:%22type=custom,data=group:${encodedString}%22%7D`;
            console.log("Teams Link:", teamsLink);

            fullLink = teamsLink;
          }

          let link = "";

          if(this.isTeams && this.isEmbedded){
            link = fullLink;
          }else if(!this.isTeams && this.isEmbedded){
            link = "https://aka.ms/VivaEngage/Outlook";
          }else{
            link = group.webUrl.Url;
          }

          return `<div class="${styles.col}">
            <div class="${styles.ContentBox}">
              <div class="${styles.ImgPart}">
                <a href="${link}" target="_blank">
                  <img src="${group.mugshotUrlTemplate.replace('{width}', '500').replace('{height}', '500')}" alt="${group.fullName}">
                </a>
              </div>
              <div class="${styles.Contents}">
              <h3 onclick="window.open('${link}')">${group.fullName}</h3>
                <p>${group.description}</p>
              </div>
            </div>
          </div>`;
        }).join('')}
      </section>
    </div>
  `;
}



  private async getCommunityInfo() {
    try {
      const tokenProvider = await this.context.aadTokenProviderFactory.getTokenProvider();
      const token = await tokenProvider.getToken("https://api.yammer.com");
  
      const response = await this.context.httpClient.get(
        `https://api.yammer.com/api/v1/groups.json?mine=1`,
        HttpClient.configurations.v1,
        {
          headers: {
            Authorization: `Bearer ${token}`,
            'Content-type': 'application/json',
          },
        }
      );
  
      const data = await response.json();
      console.log("api response:", data);
  
      if (data) {
        this.communityInfo = data
          .map((group: any) => ({
            fullName: group.full_name,
            description: group.description,
            webUrl: group.web_url,
            mugshotUrlTemplate: group.mugshot_url_template,
          }))
          .filter((group: any) => !this.excludedTitles.includes(group.fullName));
      } else {
        console.error('Groups not found in the Yammer API response.');
      }
    } catch (error) {
      console.error('Error fetching community information:', error);
    }
  
    this.render(); // Render the web part after fetching data
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
                }),
                PropertyPaneSlider('numberOfBlocks', {
                  label: 'Number of Blocks',
                  min: 1,
                  max: 10,
                  step: 1
                }),
                PropertyPaneTextField('seeAllButton',{

                  label: 'Url for See All button'

                })
              ]
            }
          ]
        }
      ]
    };
  }
}
