import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import SelectGabaritDialog from './SelectGabaritDialog';
import { 
  SPHttpClient, ISPHttpClientOptions, 
  SPHttpClientResponse, IHttpClientOptions,
  HttpClientResponse, 
  HttpClient } from 
  '@microsoft/sp-http';
  

import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';

import * as strings from 'GabaritsCommandSetStrings';

import pnp from "sp-pnp-js";
import { Web } from  "sp-pnp-js";
  
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGabaritsCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'GabaritsCommandSet';

export default class GabaritsCommandSet extends BaseListViewCommandSet<IGabaritsCommandSetProperties> {
  private urlTemplateLibrary: string = '';
  
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized GabaritsCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 1;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':

        
        //Dialog.prompt(`patate`).then((value2: string)=>{Dialog.alert(value2);});
        break;
      case 'COMMAND_2':

        const dialog: SelectGabaritDialog = new SelectGabaritDialog();
        dialog.message = 'Select template:';
        //dialog.gabaritName = "Patate";
        
        
        dialog.show().then(() => {
          //this._colorCode = dialog.colorCode;
          Dialog.alert(`New Document name: ${dialog.gabaritName}`);

          this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/GetStorageEntity('TemplatesRepoKey')`,SPHttpClient.configurations.v1).then((response:SPHttpClientResponse) => {
            response.json().then((responseJson:any) => {      
            console.log(LOG_SOURCE, responseJson);            
            this.urlTemplateLibrary = responseJson.Value;
            
            if ( typeof this.urlTemplateLibrary != 'undefined' && this.urlTemplateLibrary){
            
              let webTemplate = new Web(this.urlTemplateLibrary)
              
              //Retrieve document template
              
              webTemplate.getFileByServerRelativeUrl("/templates/gabaritTest2.docx").getBuffer().then((buffer:
              ArrayBuffer) => {      
                pnp.sp.web.getFolderByServerRelativeUrl("/Shared%20Documents/").files.add(`${dialog.gabaritName}.docx`,buffer,
                true).then(_=> Log.info(LOG_SOURCE,"done"));   
              
                Dialog.alert(`${this.urlTemplateLibrary}`);                
              });
  
            }else{Dialog.alert('urlTemplateLibrary is empty');}      
          });
        });


        
        });
        break;
      
      default:
        throw new Error('Unknown command');
    }
  }

  
}
