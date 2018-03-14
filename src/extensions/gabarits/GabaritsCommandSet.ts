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
  private urlWebTemplateLibrary: string = '';
  
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

        Dialog.alert("One document selected");
        
        break;
      case 'COMMAND_2':
       
        this.context.spHttpClient.get(`${this.context.pageContext.web.absoluteUrl}/_api/web/GetStorageEntity('TemplatesRepoKey')`,SPHttpClient.configurations.v1).then((response:SPHttpClientResponse) => {
          response.json().then((responseJson:any) => { 
           
            this.urlWebTemplateLibrary = responseJson.Value;
            console.log(`Template library : ${this.urlWebTemplateLibrary}`);

            if ( typeof this.urlWebTemplateLibrary != 'undefined' && this.urlWebTemplateLibrary){
              const dialog: SelectGabaritDialog = new SelectGabaritDialog();
              dialog.message = 'Select template:';
              console.log(LOG_SOURCE, responseJson);                                           
              dialog.urlTemplateLibrary =  this.urlWebTemplateLibrary;
              dialog.show().then(() => { 
                
              
                
              if(typeof dialog.fileName != 'undefined' && dialog.fileName)
              { 
                dialog.gabaritName = "papate.docx";
                  console.log("New document name is : " + dialog.fileName);      
                  if ( typeof dialog.gabaritName != 'undefined' && dialog.gabaritName)
                  {
                    let webTemplate = new Web(this.urlWebTemplateLibrary)
                    
                    //Retrieve document template                    
                    webTemplate.getFileByServerRelativeUrl(`/sites/templates/templates/${dialog.gabaritName}`).getBuffer().then((buffer:ArrayBuffer) => {      
                      pnp.sp.web.getFolderByServerRelativeUrl("/Shared Documents/").files.add(`${dialog.fileName}.docx`,buffer,true).then(_=> Log.info(LOG_SOURCE,"done"));                     
                      console.log(`${this.urlWebTemplateLibrary}`);                
                      console.log(`template size ${buffer.byteLength}`);
                    });
              
                  }
                  else{ Dialog.alert("No template was selected"); }
              }
              else{ Dialog.alert("No new document name was entered"); }
  
              });
            }
            else{Dialog.alert("Unable to find template library. Please advise your Administrator");}           
          });
        });

        break;
      default:
        throw new Error('Unknown command');
    }
  }
 
}
