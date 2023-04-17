import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import * as strings from 'MytaskscountWebPartStrings';

import {
  SPHttpClient,
  SPHttpClientResponse
} from '@microsoft/sp-http';

export interface IMytaskscountWebPartProps {
  description: string;
}

export default class MytaskscountWebPart extends BaseClientSideWebPart<IMytaskscountWebPartProps> {

  private _taskcount:number = 0;
  private _lastRowId: number = 346000; //336800; //345800; //336800; //339170 //327716; 
  private _threshold: number = 5000;

  private _renderList(items: any): void {
    let html: string = '';
    html += `
    <div style="background-color:#1F618D; color:white; width:220px; height:80px; border-radius:3px; padding-top:20px; padding-left:40px;">
        <h3>Pending Task(s): ${items} </h3>
    </div>
      `;
  
    const listContainer: any = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
  }

  private _getTasks(): Promise<any> {

    let url = `${this.context.pageContext.site.absoluteUrl}/_api/web/Lists/GetByTitle('Tasks')/items?%24skiptoken=Paged%3DTRUE%26p_ID%3D${this._lastRowId}&%24top=${this._threshold}`;

    if(this._lastRowId == 345800){
      url = `${this.context.pageContext.site.absoluteUrl}/_api/web/Lists/GetByTitle('PTasks')/items?$top=${this._threshold}`;
    }
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'odata-version': ''
          }
        })
          .then((res: SPHttpClientResponse) => {
            
            if (res.ok) {

              return res.json();
              
            } else {
              alert(`Something went wrong! Check the error in the browser console.`);
            }        
          })
          .catch(() => {console.log()});
   
  }

  private _renderListAsync(): void {
    
    this._getTasks()
      .then((response) => {      
 
        for(let i=0; i<response.value.length; i++){

          if( response.value[i].AssignedToId != null ){
            for(let j=0; j<response.value[i].AssignedToId.length; j++){
              
              if( ( response.value[i].AssignedToId[j] == this.context.pageContext.legacyPageContext["userId"])){
                
                this._taskcount = this._taskcount +1;
              }             
  
            }              
          }

          if( response.value[i].DelegateUserId != null ){
            for(let j=0; j<response.value[i].DelegateUserId.length; j++){
              
              if( ( response.value[i].DelegateUserId[j] == this.context.pageContext.legacyPageContext["userId"])){
                
                this._taskcount = this._taskcount +1;
              }             
  
            }              
          }

          if(i == response.value.length -1 ){
            this._renderList(this._taskcount);
            
            if(response.value.length == this._threshold){

              this._lastRowId = response.value[response.value.length - 1].Id;
              this._renderListAsync();
            }
          }
        }        

      })
      .catch(() => {console.log("Render list item failed!")});
  }

  public render(): void {
    this.domElement.innerHTML = `
    <section > 
      <a href="https://banglalinkdigitalcomm.sharepoint.com/sites/vloungeonline/SitePages/MyTasks.aspx"><div id="spListContainer" /></a>
    </section>`;
    this._renderListAsync();
  }

  // protected onInit(): Promise<void> {
  //   return this._getEnvironmentMessage().then(message => {
  //     this._environmentMessage = message;
  //   });
  // }  

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
