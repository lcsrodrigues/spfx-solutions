import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import * as strings from 'WebdoorWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';

import * as $ from 'jquery';
require('../../../node_modules/bxslider/dist/jquery.bxslider.css');
require('../../../node_modules/bxslider/dist/jquery.bxslider.js');

export interface IWebdoorWebPartProps {
  description: string;
}

export default class WebdoorWebPart extends BaseClientSideWebPart<IWebdoorWebPartProps> {

  protected onInit():Promise<void>{
    
    SPComponentLoader.loadCss(this.context.pageContext.site.absoluteUrl+'/siteAssets/css/webpart-style.css');
    SPComponentLoader.loadScript(this.context.pageContext.site.absoluteUrl+'/siteAssets/js/webpart-script.js');
    
    return super.onInit();
  }

  public render(): void {
    var self = this;
    self.domElement.innerHTML = `
      <div class='webpart-webdoor'>
        <ul id='target-webdoor'>
        </ul>
      </div>`;

      self.chargeCarousel();
  }

  protected chargeCarousel():void{

    var self = this;
    var REST = self.context.pageContext.site.absoluteUrl+"/_api/web/lists/getbytitle('WebDoor')/items";

    $.ajax({
      url:REST,
      type:"GET",
      headers:{
        "accept":"application/json;odata=verbose"
      },
      success:function(data)
      {
        var results = data.d.results;
        var strOut = "";
        var target = $("#target-webdoor");

        if(results.length)
        {
          
          for(var I=0; I<results.length; I++)
          {
            strOut +="<li class='card-webdoor' itemID='"+results[I].ID+"'>";
            strOut +="  <img src='"+results[I].imagem.Url+"'>";
            strOut +="  <span>"+results[I].Title+"</span>";
            strOut +="</li>";
          }
          
          target.html(strOut);
          target.bxSlider(
            {
              auto: true,
              autoControls: false,
              stopAutoOnClick: true,
              pager: true
            }
          );
        }
      },
      error:function(err)
      {
        console.log(err);
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
