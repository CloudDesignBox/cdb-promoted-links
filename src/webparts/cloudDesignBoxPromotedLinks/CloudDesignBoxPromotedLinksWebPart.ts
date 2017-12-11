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

import centralstyles from '../../sharedresources/centralstyles.module.scss';
import styles from './CloudDesignBoxPromotedLinksWebPart.module.scss';
import * as strings from 'CloudDesignBoxPromotedLinksWebPartStrings';

//load jquery
import * as jQuery from 'jquery';

export interface IPromotedLinksByCloudDesignBoxWebPartProps {
  description: string;
  imagelibraryname: string;
  tilecolour: string;
  tileanimation: boolean;
  Color: string;
  backgroundsize: string;
  showtitle: boolean;
  webparttitle: string;
  setwidth: string;
  themecolour:boolean;
  showaddbutton:boolean;
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

export default class CloudDesignBoxPromotedLinksWebPartWebPart extends BaseClientSideWebPart<IPromotedLinksByCloudDesignBoxWebPartProps> {
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
    return this.context.spHttpClient.get(this.context.pageContext.web.absoluteUrl + `/_api/web/lists/getbytitle('${escape(this.properties.imagelibraryname)}')/Items?$select=Title,BackgroundImageLocation,Description,LinkLocation,LaunchBehavior,Order&$top=4999&$orderby=TileOrder,Title`, SPHttpClient.configurations.v1)
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
          let cdbcolour: string = `background-color:${this.properties.tilecolour};`;
          let cdblaunchbeh: string = "";
          let cdbbackgimage:string = "";
          let cdbdescription:string = "Click here";
          //validate launch beha
          if(item.LaunchBehavior == "New tab"){
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
          //if theming is used, don't set bg colour
          if(this.properties.themecolour==true){
            cdbcolour=``;
          }
          //if tile size isn't set - make default
          if(this.properties.setwidth==null || typeof this.properties.setwidth === 'undefined'){
            this.properties.setwidth="150px";
          }
          html+=`<div style="width:${this.properties.setwidth};display:inline-block;" class="${styles.mobiletile}">
            <div>
              <div class="${styles.tiles}">
                <div class="${styles.tilecontent} ${styles.tpmouse}" style="${cdbcolour}background-image:url('${cdbbackgimage}');background-size:${this.properties.backgroundsize};position:relative;" onclick="${cdblaunchbeh}">
                  <div class="${styles.cdbdescholder}">
                    <div class="${styles.cdbdescholdertitle}"><span>${item.Title}</span></div>
                    <div class="${styles.cdbdescholderdesc}"><span>${cdbdescription}</span></div>
                  </div>
                </div>
              </div>
            </div>
          </div>`;
        });
        //add new link button
        if(this.properties.showaddbutton == true){
          html += `<div style="width:${this.properties.setwidth};display:inline-block;" class="${styles.mobiletile}">
            <div>
              <div class="${styles.tiles}">
                <div class="${centralstyles.cdbplusbuttoncontainer}"><a href="${escape(this.context.pageContext.web.absoluteUrl)}/Lists/${this.properties.imagelibraryname}/NewForm.aspx?Source=${encodeURIComponent(window.location.href).replace(".","%2E")}" class="${centralstyles.cdbplusbuttonleft}">&#43;</a><span class="${centralstyles.cdbclear}">add</span></div>
              </div>
            </div>
          </div>`;
        }
      }else{
        html = `There are no links in this list.<br />`;                
      }
    }else{
        html = `Open the web part properties and select a list that already exists on the site or create a new promoted links list on this site using the link below.<br /><a href="${escape(this.context.pageContext.web.absoluteUrl)}/_layouts/15/viewlsts.aspx" class="${centralstyles.cdbbutton}">Manage Lists</a>`;
    }

    //render results
    const listContainer: Element = this.domElement.querySelector(`#${styles.row}`);
    listContainer.innerHTML = html;

    //load jQuery features
    jQuery(() => {
      jQuery(`.${styles.cdbsubjecttiles}`).show();
    });

    //only load animation if prop is true
    if(this.properties.tileanimation == true)
    {
      jQuery(() => {
        //remove any previous instances
        jQuery("." + styles.tilecontent).unbind();
        jQuery("." + styles.tilecontent)
          .mouseenter(function() {
            jQuery(this).children("." + styles.cdbdescholder).children("." + styles.cdbdescholderdesc).stop();
            jQuery(this).children("." + styles.cdbdescholder).children("." + styles.cdbdescholderdesc).slideToggle("fast");
        })
          .mouseleave(function() {
            jQuery(this).children("." + styles.cdbdescholder).children("." + styles.cdbdescholderdesc).stop();
            jQuery(this).children("." + styles.cdbdescholder).children("." + styles.cdbdescholderdesc).slideToggle("fast");
        });
      });
    }
    
    //end
  }

  public render(): void {
    let temphtml:string = `
      <div class="${styles.cloudDesignBoxPromotedLinks}">
        <div class="${styles.cdbsubjecttiles}">
          <div>`;
              //show header if true
    if(this.properties.showtitle == true){
      temphtml += `<div class="${centralstyles.heading}">${this.properties.webparttitle}</div>`;
    }
    temphtml += `
            <div class="${styles.actionicons} ${styles.row}" id="${styles.row}">
              <!--jquery to insert links here-->
              <span class="${styles.cdbclear}"></span>
            </div>
          </div>
        </div>
      </div>
      `;
      
      this.domElement.innerHTML = temphtml;

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
                }),
                PropertyPaneToggle('showtitle', {
                  label: "Show Title",
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneTextField('webparttitle', {
                  label: "Web Part Title"
                })
              ]
            },
            {
              groupName: "Colour and Animation",
              groupFields: [
                PropertyPaneToggle('tileanimation', {
                  label: strings.TileAnimation,
                  onText: 'On',
                  offText: 'Off'
                }),
                PropertyPaneToggle('themecolour', {
                  label: "Use Theme for Tile Background Colour (not colour below)?",
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
               })
              ]
            },
            {
              groupName: "Size",
              groupFields: [
                PropertyPaneDropdown('backgroundsize', {
                  label: strings.BackgroundSizeFieldLabel,
                  options: [
                    { key: 'cover', text: 'cover' },
                    { key: 'auto', text: 'auto' },
                    { key: '100%', text: '100%' },
                    { key: '50%', text: '50%' },
                    { key: '30%', text: '30%' },
                    { key: '27%', text: '27%' }
                  ]
                }),
                PropertyPaneDropdown('setwidth', {
                  label: "Tile Width - desktop view only",
                  options: [
                    { key: '150px', text: '150px' },
                    { key: '50%', text: '50% - 2 in a row' },
                    { key: '33%', text: '33% - 3 in a row' },
                    { key: '25%', text: '25% - 4 in a row' },
                    { key: '20%', text: '20% - 5 in a row' }
                  ]
                }),
                PropertyPaneToggle('showaddbutton', {
                  label: "Show add new link button?",
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


