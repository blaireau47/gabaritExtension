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
import ConfigurationInfo from './ConfigurationInfo';
  
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGabaritsCommandSetProperties {

}

const LOG_SOURCE: string = 'GabaritsCommandSet';

export default class GabaritsCommandSet extends BaseListViewCommandSet<IGabaritsCommandSetProperties> {
  private urlWebTemplateLibrary: string = '';
  
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized GabaritsCommandSet');
    console.log("Version 1.0.0.2");
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_GET_ENT_TEMPLATES');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      compareOneCommand.visible = event.selectedRows.length === 0;
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {

      case 'COMMAND_GET_ENT_TEMPLATES':

        //Recupere les information de configuration 
        let configInfo = new ConfigurationInfo("https://devebl.sharepoint.com/sites/Templates/Templates/configInfos.json",this.context.pageContext.web.absoluteUrl, this.context.spHttpClient);

        configInfo.getConfigurationData().then(()=> {

          //console.log("*************" + objConfig);
          this.urlWebTemplateLibrary = configInfo.templateLibraryUrl;
          
          configInfo.getTemplatesData().then((_gabarit:any[]) => {
                        
              if ( typeof this.urlWebTemplateLibrary != 'undefined' && this.urlWebTemplateLibrary){
                const dialog: SelectGabaritDialog = new SelectGabaritDialog();
                dialog.message = 'Select template:';
                dialog.gabarits = _gabarit;                                  
                dialog.urlTemplateLibrary =  this.urlWebTemplateLibrary;
                dialog.show().then(() => { 
                                                
                  if(typeof dialog.fileName != 'undefined' && dialog.fileName)
                  { 
                      if ( typeof dialog.gabaritName != 'undefined' && dialog.gabaritName)
                      {                     
                        dialog.gabaritName = "papate.docx";
                        console.log("Nogabrit was selected so using default : " + dialog.gabaritName);     
                      }
                      console.log("New document name is : " + dialog.fileName);      
                      if ( typeof dialog.gabaritName != 'undefined' && dialog.gabaritName)
                      {
                        let webTemplate = new Web(this.urlWebTemplateLibrary)
                        
                        //Retrieve document template             
                        //this.context.pageContext.list.serverRelativeUrl.       
                        webTemplate.getFileByServerRelativeUrl(`/sites/templates/templates/${dialog.gabaritName}`).getBuffer().then((buffer:ArrayBuffer) => {      
                          pnp.sp.web.getFolderByServerRelativeUrl(this.context.pageContext.list.serverRelativeUrl).files.add(`${dialog.fileName}.docx`,buffer,true).then(_=>  window.location.href = document.location.href);                     
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
