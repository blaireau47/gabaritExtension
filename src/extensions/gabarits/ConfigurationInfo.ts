import { 
    SPHttpClient, ISPHttpClientOptions, 
    SPHttpClientResponse, IHttpClientOptions,
    HttpClientResponse, 
    HttpClient } from 
    '@microsoft/sp-http';

import { Log } from '@microsoft/sp-core-library';
import pnp from "sp-pnp-js";
import { Web } from  "sp-pnp-js";

const requestHeaders: Headers = new Headers();
    
requestHeaders.append('Content-type', 'application/json');
requestHeaders.append('Cache-Control', 'no-cache');

const httpClientOptions: IHttpClientOptions = {  
  headers: requestHeaders
};


export default class ConfigurationInfo {
    public configUrl: string; //Contient le url complet de la configuration
    public templateLibraryUrl:string;
    public templateType: string;
    private httpClient: SPHttpClient;
    private absoluteUrl: string;

    constructor(_configUrl: string,_absoluteUrl:string,  _http:SPHttpClient){
        
        this.httpClient = _http;
        this.absoluteUrl = _absoluteUrl;

        if( typeof _configUrl != 'undefined' && _configUrl){
            this.configUrl = _configUrl;     
        }else{this.configUrl = "https://localhost:4321/dist/configInfos.json";}
           
    
    }


  private config_data: any;
  public getConfigurationData() : Promise<void> {
        
    return this.httpClient.get(`${this.absoluteUrl}/_api/web/GetStorageEntity('TemplatesRepoKey')`,SPHttpClient.configurations.v1).then((response:SPHttpClientResponse) => {

        return response.json().then((responseJson:any) => { 
            this.templateLibraryUrl = responseJson.Value;
            pnp.sp.site.rootWeb.select("AllProperties").expand("AllProperties").get().then(w => {
                
                this.templateType = "pont";
                return Promise.resolve();

            });            
        });

    });

    /*

        if(typeof(this.config_data) === "undefined") {
                return this.httpClient.get( this.configUrl,SPHttpClient.configurations.v1,httpClientOptions).then(res => {
                    this.config_data = res.json().then(body => 
                        this.config_data = body);
                    return this.config_data;
                }).catch(this.handleError);
        } else {
            return Promise.resolve(this.config_data);
        }
    */

  }

  
  public getTemplatesData() : Promise<any[]> {

    /*
    const xml =  "<View>"+
                    "<ViewFields><FieldRef Name='ID' /><FieldRef Name='FileLeafRef' /><FieldRef Name='LinkFilenameNoMenu' /><FieldRef Name='Title' /></ViewFields>"+
                    "<Query>"+
                      "<Where><Contains><FieldRef Name='ProjectType' /><Value Type='TaxonomyFieldTypeMulti'>Route</Value></Contains></Where>"+
                      "<OrderBy><FieldRef Name='LinkFilename' Ascending='False' /></OrderBy>"+
                    "</Query>"+
                  "</View>";
    
    const q: CamlQuery = {
      ViewXml: xml,
    };    
    */

    const gabarits: any[] = [];
    if ( typeof this.templateLibraryUrl != 'undefined' && this.templateLibraryUrl){
        
        let webTemplate = new Web(this.templateLibraryUrl)
        webTemplate.lists.getByTitle("Templates").items.select("FileLeafRef","Title","Id").filter("Title eq 'pont'").get().then((docTemplates) => { 

            if (docTemplates.length != 0) {
                console.log(JSON.stringify(docTemplates, null, 4));
                for (let docTemplate of   docTemplates) {
                    console.log("doc template name " + docTemplate.Name);
                    gabarits.push({
                    key: docTemplate.Id,
                    name: docTemplate.FileLeafRef,
                    value: docTemplate.FileLeafRef
                    });
                    console.log("_items pushd has  " + gabarits.length);
                }
            }
            return Promise.resolve(gabarits);
        }); 
    }
    else {
        console.log('urlTemplateLibrary is empty please configure tenant Key');
        return Promise.reject(gabarits);
    }

     
  }

  private handleError(error: Response){

    console.log(`Impossible de recupérer le fichier ${error} `);
  }

}




