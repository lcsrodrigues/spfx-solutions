import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './PropriedadesUsuarioWebPart.module.scss';
import * as strings from 'PropriedadesUsuarioWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';

export interface IPropriedadesUsuarioWebPartProps {
  description: string;
  iduser: string;
}

import * as $ from "jquery";

export default class PropriedadesUsuarioWebPart extends BaseClientSideWebPart<IPropriedadesUsuarioWebPartProps> {
  
  protected onInit():Promise<void>{
    
    SPComponentLoader.loadCss(this.context.pageContext.site.absoluteUrl+'/SiteAssets/css/style.css');
    
    return super.onInit();
  }

  public render(): void {
    var self = this;
    this.domElement.innerHTML = `
    <div>
      <h1>Propriedades do usuário</h1>
      <ul id="user-properties">
      </ul>
    </div>`;
      
    self.getUsers();
  }

  protected getUsers(): void {
    
    var self = this;
    var REST =  this.context.pageContext.site.absoluteUrl+"/_api/web/SiteUsers?"
                +"$expand=Groups,Alerts&"
                +"$orderby=ID asc";

        $.ajax({
        url: REST,
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose"
        },
        success: function(data)
        {
            var results = data.d.results;
           
            if(results.length)
            {
                for( var I=0; I<results.length; I++ )
                {
                  self.getAllPropertiesUser(results[I]);
                }
            }
        },
        error:function(error)
        {
            console.log(JSON.stringify(error));            
        }
    });
  }

  protected getAllPropertiesUser(paramUser):void {

    var self = this;
    var REST = this.context.pageContext.site.absoluteUrl+"/_api/SP.UserProfiles.PeopleManager/GetPropertiesFor(accountName=@v)?@v='i:0%23.f|membership|"+paramUser.Email+"'";
  
    $.ajax({
        url: REST,
        type: "GET",
        headers: {
            "accept": "application/json;odata=verbose",
        },
        success: function(data)
        {
          var user = data.d;
          if(typeof(user.Email) != "undefined")
          {
            var allPropertiesUser = user.UserProfileProperties.results;
        
            //Propriedades do usuário
            var nameUser = user.DisplayName;
            var pictureUser = user.PictureUrl;
            var emailUser = user.Email;
            var oneDriveUser = user.PersonalUrl;
            var birthdayUser = "-";
            var departamentUser = "-";
            var arrayGroupsUser = paramUser.Groups.results;
        
            if(pictureUser == null)
            {
              pictureUser = "https://cdn.pixabay.com/photo/2015/10/05/22/37/blank-profile-picture-973460_960_720.png";
            }
            
            for(var I=0; I<allPropertiesUser.length; I++)
            {
              var currentProperties = allPropertiesUser[I];
              switch(currentProperties.Key)
              {
                case 'SPS-Birthday':
                if(currentProperties.Value)
                {
                  birthdayUser = currentProperties.Value.split(" ")[0];
                }
                break;
                case 'Department':
                if(currentProperties.Value)
                {
                  departamentUser = currentProperties.Value;
                }
                break;
              }
            }
            
            var strOut = "";
            
            strOut+="	<li class='card-user'>";
            strOut+="		<div class='header-user'>";
            strOut+="			<div class='properties-user'>";
            strOut+="				<span>"+nameUser+"</span><br/>";
            strOut+="				<span>"+emailUser+"</span><br/>";
            strOut+="				<span>"+birthdayUser+"</span><br/>";
            strOut+="			</div>";
            strOut+="			<div class='imagem-user'>";
            strOut+="				<img src='"+pictureUser+"' />";
            strOut+="			</div>";
            strOut+="		</div>";
            strOut+="		<div class='body-user'>";
            strOut+="			<span><a href='"+oneDriveUser+"' target='_blank'>OneDrive</a></span><br/>";
            strOut+="			<span>Departamento: "+departamentUser+"</span>";
            
            if(arrayGroupsUser.length)
            {
              strOut+="<br/>";
              strOut+="<ul><a>Grupos</a><br/>";
              
              for(var I=0; I<arrayGroupsUser.length; I++)
              {
                var titleGroup = arrayGroupsUser[I].Title;
                
                if(titleGroup.indexOf("SharingLinks") == -1)
                {
                  if( titleGroup.length > 25 )
                  {
                    titleGroup = titleGroup.substring(0,25)+"...";
                  }
                  strOut+="<li title='"+arrayGroupsUser[I].Title+"' >"+ titleGroup +"</li><br/>";
                }
              }
              
              strOut+="<ul/>";
            }
            
            strOut+="		</div>";
            strOut+="	</li>";
        
            $("#user-properties").append(strOut);
            self.showMoreProperties();
          }
        },
        error:function(error)
        {
            console.log(JSON.stringify(error));
        }
    });
  }

  protected showMoreProperties() {
    $("#user-properties").on("mouseover",".card-user",function()
    {
      $(this).find(".body-user").show();
    }).on("mouseleave",".card-user",function()
    {
      $(this).find(".body-user").hide();	
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
                }),
                PropertyPaneTextField('iduser', {
                  label: strings.iduser
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
