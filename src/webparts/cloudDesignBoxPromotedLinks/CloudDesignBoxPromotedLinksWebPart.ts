import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
//import for rest calls
import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';

import styles from './CloudDesignBoxPromotedLinksWebPart.module.scss';
import * as strings from 'CloudDesignBoxPromotedLinksWebPartStrings';

//load jquery with jquery cycle as dependancy
import * as jQuery from 'jquery';

export interface ICloudDesignBoxPromotedLinksWebPartProps {
  description: string;
  imagelibraryname: string;
  tilecolour: string;
  tileanimation: boolean;
}

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Title: string;
  BackgroundImageLocation: string;
  Description: string;
  LinkLocation: string;
  LaunchBehavior: string;
  Order: string;
}

export default class CloudDesignBoxPromotedLinksWebPartWebPart extends BaseClientSideWebPart<ICloudDesignBoxPromotedLinksWebPartProps> {

  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${escape(this.properties.imagelibraryname)}')/Items?$select=Title,BackgroundImageLocation,Description,LinkLocation,LaunchBehavior,Order&$orderby=TileOrder,Title`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
  }

  private _renderList(items: ISPList[]): void {
    //store html and colours and data
    let html: string = "";
    
    //check if items exist
    if (typeof(items) != "undefined"){
      //check if there are any items
      if (items.length > 0){
        //loop through items and add to html
        items.forEach((item: ISPList) => {
          let cdbcolour: string = this.properties.tilecolour;
          let cdblaunchbeh: string = "";
          let cdbbackgimage:string = "";
          let cdbdescription:string = "Click here";
          //validate launch beha
          if(item.LaunchBehavior != "New tab"){
            cdblaunchbeh=`window.open('${item.LinkLocation['Url']}','_blank');`;
          }else{
            cdblaunchbeh=`location.href='${item.LinkLocation['Url']}';`;
          }
          //validate bg image
          if(item.BackgroundImageLocation != null){
            cdbbackgimage=item.BackgroundImageLocation['Url'];
          }
          //validate desc
          if(item.Description != null){
            cdbdescription=item.Description;
          }
          html+=`
          <div class="${styles.tiles}">
            <div class="${styles.tilecontent} ${styles.tpmouse}" style="background-color:${cdbcolour};background-image:url('${cdbbackgimage}');position:relative;" onclick="${cdblaunchbeh}">
              <div class="${styles.cdbdescholder}">
                <div class="${styles.cdbdescholdertitle}"><span>${item.Title}</span></div>
                <div class="${styles.cdbdescholderdesc}"><span>${cdbdescription}</span></div>
              </div>
            </div>
          </div>`;
        });
      }else{
        html = `There are no links in this list.<br /><a href="${escape(this.context.pageContext.web.absoluteUrl)}/Lists/Promoted%20Links/AllItems.aspx" class="${styles.cdbbutton}">Add Links</a>`;                
      }
    }else{
        html = `This list does not exist. <br /><a href="${escape(this.context.pageContext.web.absoluteUrl)}/_layouts/15/viewlsts.aspx" class="${styles.cdbbutton}">Manage Lists</a>`;
    }

    //render results
    const listContainer: Element = this.domElement.querySelector(`#${styles.row}`);
    listContainer.innerHTML = html;

    //load jQuery features
    jQuery(document).ready(function() {
      jQuery(`.${styles.cdbsubjecttiles}`).show();
    });
    //only load animation if prop is true
    if(this.properties.tileanimation == true)
    {
      jQuery( document ).ready(function() {
        jQuery("." + styles.tilecontent)
          .mouseenter(function() {
            jQuery(this).children("." + styles.cdbdescholder).children("." + styles.cdbdescholderdesc).slideToggle("slow");
        })
          .mouseleave(function() {
            jQuery(this).children("." + styles.cdbdescholder).children("." + styles.cdbdescholderdesc).slideToggle("slow");
        });
      });
    }
    
    //end
  }

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.cloudDesignBoxPromotedLinks}">
        <div class="${styles.cdbsubjecttiles}"><div><div class="${styles.actionicons} ${styles.row}" id="${styles.row}">
          <!--jquery to insert links here-->
        </div><span class="${styles.cdbclear}"></span></div></div>
      </div>
      `;

      //render promoted links data
      this._getListData()
      .then((response) => {
        this._renderList(response.value);
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
                PropertyPaneTextField('imagelibraryname', {
                  label: strings.ImageLibraryFieldLabel
                }),
                PropertyPaneTextField('tilecolour', {
                  label: strings.TileColour
                }),
                PropertyPaneToggle('tileanimation', {
                  label: strings.TileAnimation,
                  onText: 'On',
                  offText: 'Off'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}


// To do:
// option to pick backgorund size
// test everything
//show drop down of promoted lists in web part properties
// create a new promoted list if it doesnt exist?
//restrict colour picker
