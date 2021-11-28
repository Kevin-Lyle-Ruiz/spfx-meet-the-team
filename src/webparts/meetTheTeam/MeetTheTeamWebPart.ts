import { Version } from "@microsoft/sp-core-library";
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from "@microsoft/sp-webpart-base";

import styles from "./MeetTheTeamWebPart.module.scss";
import * as strings from "MeetTheTeamWebPartStrings";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

require("../../../node_modules/@fortawesome/fontawesome-free/css/all.min.css");
require("./components/bucketStyle.css");

export interface ISPList {
  value: ISPListItem[];
}

export interface ISPListItem {
  ID: string;
  Profile_x0020_Image: {
    [Url: string]: string
  };
  Last_x0020_Name: string;
  First_x0020_Name: string;
  Job_x0020_Title: string;
  Jabber_x0020_Extension: string;
  Email: string;
  Jabber_x0020_Instant_x0020_Messa: string;
  Bucket: string;
  Additional_x0020_Information: {
    [Url: string]: string
  };
}

export interface IMeetTheTeamWebPartProps {
  description: string;
}

export default class MeetTheTeamWebPart extends BaseClientSideWebPart<IMeetTheTeamWebPartProps> {
  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.meetTheTeam}">
        <div class="${styles.webpart_container}">
          <div id="mainPage" class="${styles.webpart_grid}">
          </div>

          <button class="${styles.webpart_button} accordion" type="button">Resource Development / Operations</button>
          <div id="resourceDevelopment" class="panel ${styles.webpart_grid}">
          </div>

          <button class="${styles.webpart_button} accordion" type="button">PS&RS Consultants</button>
          <div id="consultants" class="panel ${styles.webpart_grid}">
          </div>
        </div>
      </div>
    `;

    this._renderList();

    const acc = document.getElementsByClassName("accordion");
    let i: number;

    for (i = 0; i < acc.length; i++) {
      acc[i].addEventListener("click", function () {
        this.classList.toggle("active");
        const panel = this.nextElementSibling;
        if (panel.style.maxHeight) {
          panel.style.maxHeight = null;
        } else {
          panel.style.maxHeight = panel.scrollHeight + "px";
        }
      });
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription,
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField("description", {
                  label: strings.DescriptionFieldLabel,
                }),
              ],
            },
          ],
        },
      ],
    };
  }

  private _getListData(): Promise<ISPList> {
    return this.context.spHttpClient
      .get(
        // this.context.pageContext.web.absoluteUrl +
        "https://thehub.medpro.com/sites/psrs/public/_api/lists/getbytitle('Meet%20The%20Team')/Items",
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderList(): void {
    this._getListData().then((response) => {
      let mainPageHTML: string = "";
      let resourceDevelopmentHTML: string = "";
      let consultantsHTML: string = "";

      response.value.forEach((item: ISPListItem) => {
        const bucket: string = item.Bucket;
        const attachmentHtml: string = (item.Additional_x0020_Information ? `<a href="${item.Additional_x0020_Information.Url}" target="_blank"><i class="fa fa-info-circle ${styles.social_icons} ${styles.webpart_icon}"></i></a>` : "");
        const imageSource: string = (item.Profile_x0020_Image ? item.Profile_x0020_Image.Url : `https://thehub.medpro.com/sites/psrs/public/PublishingImages/Pages/Meet%20The%20Team/Placeholder.jpg`);

        if (bucket === "Resource Development/Operations") {
          resourceDevelopmentHTML += `
            <div class="${styles.webpart_card}">
              <div>
                <img class="${styles.profile_image}" src="${imageSource}" />
              </div>
              <h1 class="${styles.webpart_header}">${item.First_x0020_Name} ${item.Last_x0020_Name}</h1>
              <div class="${styles.webpart_info}">
                <p class="${styles.job_title}">${item.Job_x0020_Title}</p>
                <div class="${styles.icon_container}">
                  <a href="CISCOTEL:${item.Jabber_x0020_Extension}"><i class="fa fa-phone-square-alt ${styles.social_icons} ${styles.webpart_icon}"></i></a>
                  <a href=mailto:${item.Email}><i class="fa fa-envelope ${styles.social_icons} ${styles.webpart_icon}"></i></a>
                  <a href="IM:${item.Jabber_x0020_Instant_x0020_Messa}"><i class="fa fa-comments ${styles.social_icons} ${styles.webpart_icon}"></i></a>
                  ${attachmentHtml}
                </div>
              </div>
            </div>
          `;
        } else if (bucket === "PS&RS Consultants") {
          consultantsHTML += `
            <div class="${styles.webpart_card}">
              <div>
                <img class="${styles.profile_image}" src="${imageSource}" />
              </div>
              <h1 class="${styles.webpart_header}">${item.First_x0020_Name} ${item.Last_x0020_Name}</h1>
              <div class="${styles.webpart_info}">
                <p class="${styles.job_title}">${item.Job_x0020_Title}</p>
                <div class="${styles.icon_container}">
                  <a href="CISCOTEL:${item.Jabber_x0020_Extension}"><i class="fa fa-phone-square-alt ${styles.social_icons} ${styles.webpart_icon}"></i></a>
                  <a href=mailto:${item.Email}><i class="fa fa-envelope ${styles.social_icons} ${styles.webpart_icon}"></i></a>
                  <a href="IM:${item.Jabber_x0020_Instant_x0020_Messa}"><i class="fa fa-comments ${styles.social_icons} ${styles.webpart_icon}"></i></a>
                  ${attachmentHtml}
                </div>
              </div>
            </div>
          `;
        } else {
          mainPageHTML += `
            <div class="${styles.webpart_card}">
              <div>
                <img class="${styles.profile_image}" src="${imageSource}" />
              </div>
              <h1 class="${styles.webpart_header}">${item.First_x0020_Name} ${item.Last_x0020_Name}</h1>
              <div class="${styles.webpart_info}">
                <p class="${styles.job_title}">${item.Job_x0020_Title}</p>
                <div class="${styles.icon_container}">
                  <a href="CISCOTEL:${item.Jabber_x0020_Extension}"><i class="fa fa-phone-square-alt ${styles.social_icons} ${styles.webpart_icon}"></i></a>
                  <a href=mailto:${item.Email}><i class="fa fa-envelope ${styles.social_icons} ${styles.webpart_icon}"></i></a>
                  <a href="IM:${item.Jabber_x0020_Instant_x0020_Messa}"><i class="fa fa-comments ${styles.social_icons} ${styles.webpart_icon}"></i></a>
                  ${attachmentHtml}
                </div>
              </div>
            </div>
          `;
        }
      });

      const mainPageContainer: Element =
        this.domElement.querySelector("#mainPage");
      const resourceDevelopmentContainer: Element =
        this.domElement.querySelector("#resourceDevelopment");
      const consultantsContainer: Element =
        this.domElement.querySelector("#consultants");

      mainPageContainer.innerHTML = mainPageHTML;
      resourceDevelopmentContainer.innerHTML = resourceDevelopmentHTML;
      consultantsContainer.innerHTML = consultantsHTML;
    });
  }

  protected get dataVersion(): Version {
    return Version.parse("1.0");
  }
}
