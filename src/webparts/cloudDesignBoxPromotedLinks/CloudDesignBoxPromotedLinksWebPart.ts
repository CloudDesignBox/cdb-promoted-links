import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
//for list import
import { IODataList } from '@microsoft/sp-odata-types';
//import for rest calls
import {
  SPHttpClient,
  SPHttpClientResponse   
} from '@microsoft/sp-http';

//import colour picker library - third party library https://oliviercc.github.io/sp-client-custom-fields
import { PropertyFieldColorPicker } from 'sp-client-custom-fields/lib/PropertyFieldColorPicker';

import styles from './CloudDesignBoxPromotedLinksWebPart.module.scss';
import * as strings from 'CloudDesignBoxPromotedLinksWebPartStrings';

//load jquery
import * as jQuery from 'jquery';

export interface ICloudDesignBoxPromotedLinksWebPartProps {
  description: string;
  imagelibraryname: string;
  tilecolour: string;
  tileanimation: boolean;
  Color: string;
  backgroundsize: string;
}
//promoted lists to populate properties
export interface ISPListofLists {
  value: ISPListofLists[];
  Title:string;
}
export interface ISPListofListsItem{
  Title: string;
}
//items from pomoted list
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
//array data for list of lists
export interface IPromotedListLists {
  key: string;
  text: string;
}

export default class CloudDesignBoxPromotedLinksWebPartWebPart extends BaseClientSideWebPart<ICloudDesignBoxPromotedLinksWebPartProps> {
  /*Guidance on dynamically loading options into properties - http://www.sharepointnutsandbolts.com/2016/09/sharepoint-framework-spfx-web-part-properties-dynamic-dropdown.html*/
  private promotedLinksDropdown: IPromotedListLists[];
  private promotedlistsLoaded: boolean;
  private getListOfLosts(url: string) : Promise<any> {
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
      if (response.ok) {
        return response.json();
      } else {
        console.log("WARNING - failed to hit URL " + url + ". Error = " + response.statusText);
        return null;
      }
    });
  }

  private LoadGetListOfLists(): Promise<IPromotedListLists[]> {
    var requrl = this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=BaseTemplate%20eq%20170`;
    return this.getListOfLosts(requrl).then((response) => {
        var options: Array<IPromotedListLists> = new Array<IPromotedListLists>();
        response.value.map((item: IODataList) => {
            options.push( { key: item.Title, text: item.Title });
        });
        return options;
    });
  }

  //load Promoted List Data
  private _getListData(): Promise<ISPLists> {
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${escape(this.properties.imagelibraryname)}')/Items?$select=Title,BackgroundImageLocation,Description,LinkLocation,LaunchBehavior,Order&$orderby=TileOrder,Title`, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        return response.json();
      });
    }

  //render promoted links list
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
            <div class="${styles.tilecontent} ${styles.tpmouse}" style="background-color:${cdbcolour};background-image:url('${cdbbackgimage}');background-size:${this.properties.backgroundsize};position:relative;" onclick="${cdblaunchbeh}">
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
    //load promoted lists if they haven't already been loaded
    if (!this.promotedlistsLoaded) {
      this.LoadGetListOfLists().then((response) => {
        this.promotedLinksDropdown = response;
        this.promotedlistsLoaded = true;
        // refresh now that lists are loaded
        this.context.propertyPane.refresh();
        this.onDispose();
      });
   }
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
                PropertyPaneDropdown('imagelibraryname', {
                  label: strings.ImageLibraryFieldLabel,
                  options: this.promotedLinksDropdown
                })
              ]
            },
            {
              groupName: strings.AdvancedGroupName,
              groupFields: [
                PropertyPaneToggle('tileanimation', {
                  label: strings.TileAnimation,
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyFieldColorPicker('tilecolour', {
                  label: strings.TileColour,
                  initialColor: this.properties.tilecolour,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  render: this.render.bind(this),
                  disableReactivePropertyChanges: this.disableReactivePropertyChanges,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'colorFieldId'
               }),
               PropertyPaneDropdown('backgroundsize', {
                label: strings.BackgroundSizeFieldLabel,
                options: [
                  { key: 'cover', text: 'Cover' },
                  { key: 'auto', text: 'Auto' },
                  { key: '100%', text: '100%' },
                  { key: '50%', text: '50%' },
                ]
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
// create a new promoted list if it doesnt exist?
