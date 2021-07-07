import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import {
  SPHttpClient,
  SPHttpClientResponse,
  ISPHttpClientOptions
} from '@microsoft/sp-http';
import styles from './CreatenewlistWebPart.module.scss';
import * as strings from 'CreatenewlistWebPartStrings';

export interface ICreatenewlistWebPartProps {
  description: string;
}

export default class CreatenewlistWebPart extends BaseClientSideWebPart<ICreatenewlistWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${ styles.createnewlist }">
        <h3>Create a new list dynamically</h3><br/>
        <p>Listname:<br/> <input type="text" id="txtlistname"></input></p><br/>
        <p>List Description: <input type="text" id="txtlistdescr"></input></p><br/>
        <p><input type="button" id="btnlist" value="Click to create list"/></p><br/>
      </div>`;
      this.bindEvents();
  }
  private bindEvents():void{
    this.domElement.querySelector('#btnlist').addEventListener('click',()=>{this.createnewlist();});
  }
  private createnewlist():void{
      var newlistname=document.getElementById("txtlistname")["value"];
      var newlistdesc=document.getElementById("txtlistdescr")["value"];

      const listurl:string=this.context.pageContext.web.absoluteUrl+"/_api/web/lists/GetbyTitle('"+newlistname+"')";
      this.context.spHttpClient.get(listurl,SPHttpClient.configurations.v1)
      .then((response:SPHttpClientResponse)=>{
        if(response.status===200){
          alert("list already exists");
          return;
        }
        if(response.status===404){
          const url:string=this.context.pageContext.web.absoluteUrl+"/_api/web/lists";
          const listdefinition:any={
              "Title": newlistname,
              "Description":newlistdesc,
              "ContentTypeEnabled":true,
              "AllowContentType":true,
              "Basetemplate":100
          };
          const sphttpclientoptions:ISPHttpClientOptions={
            "body":JSON.stringify(listdefinition)
          };
          this.context.spHttpClient.post(url,SPHttpClient.configurations.v1,sphttpclientoptions)
          .then((response:SPHttpClientResponse)=>{
            if(response.status===201){
              alert("lust created success");
            }else{
              alert("Errormesssage"+response.status+","+response.statusText);
            }
          });
        }
        else{
          alert("Error.messsage"+response.status+","+response.statusText);
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
