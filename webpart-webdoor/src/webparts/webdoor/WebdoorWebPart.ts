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
import 'slick-carousel';

export interface IWebdoorWebPartProps {
  description: string;
}

export default class WebdoorWebPart extends BaseClientSideWebPart<IWebdoorWebPartProps> {

  protected onInit():Promise<void>{
    
    SPComponentLoader.loadCss(this.context.pageContext.site.absoluteUrl+'/_catalogs/masterpage/siteAssets/js/slick-1.8.1/slick/slick.css');
    SPComponentLoader.loadCss(this.context.pageContext.site.absoluteUrl+'/_catalogs/masterpage/siteAssets/js/slick-1.8.1/slick/slick-theme.css');
    SPComponentLoader.loadCss(this.context.pageContext.site.absoluteUrl+'/siteAssets/css/webpart-style.css');
    
    SPComponentLoader.loadScript(this.context.pageContext.site.absoluteUrl+'/siteAssets/js/webpart-script.js');

//https://devrodrigues.sharepoint.com/sites/meu-portal/_catalogs/masterpage/siteAssets/js/slick-1.8.1/slick/slick.css
//    <link rel="stylesheet" type="text/css" href="siteAssets/js/slick-1.8.1/slick/slick.css"/>
//		<link rel="stylesheet" type="text/css" href="siteAssets/js/slick-1.8.1/slick/slick-theme.css"/>		
//    <script type="text/javascript" src="siteAssets/js/slick-1.8.1/slick/slick.js"></script>

    return super.onInit();
  }

  public render(): void {
    var self = this;
    self.domElement.innerHTML = `
      <div class='webpart-webdoor'>
        <section id='target-webdoor'>
        </div>
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
            strOut +="<div class='card-webdoor' itemID='"+results[I].ID+"'>";
            strOut +="  <img src='"+results[I].imagem.Url+"'>";
            strOut +="  <span>"+results[I].Title+"</span>";
            strOut +="</div>";
          }
          
          target.html(strOut);
          target.slick({
            dots: false,
            infinite: true,
            speed: 1000,
            autoplay:true,
            autoplaySpeed:3000,
            fade: true,
            cssEase: 'linear',
            pauseOnHover:true
          });
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
