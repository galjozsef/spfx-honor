import { sp } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult, IItemUpdateResult } from "@pnp/sp/items";
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';

import * as jQuery from 'jquery';
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  RowAccessor,
  ListViewAccessor,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import './Ext/bootoast.css';
import './Ext/bootstrap.css';
import pnp from 'sp-pnp-js';

import * as strings from 'TigFeladasCaCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ITigFeladasCaCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'TigFeladasCaCommandSet';
var selItems: string="";
var lElso: boolean=false;
var bootoast:any=require('./Ext/bootoast.js');

export default class TigFeladasCaCommandSet extends BaseListViewCommandSet<ITigFeladasCaCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized TigFeladasCaCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      if (decodeURIComponent(window.location.href).toLowerCase().indexOf("tigdokumentumok/forms/sap exportlsra vr.aspx")!=-1)
      {
        if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=")==-1)
        {
            compareOneCommand.visible = event.selectedRows.length >0;
        }
        else
        {
          if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=d169cf86-c080-4a4d-bc44-8155b0cebb2c")!=-1)
          {
            compareOneCommand.visible = event.selectedRows.length >0;
          }
          else
          {
            compareOneCommand.visible = false;
          }
        }
      }
      else
      {
        if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=d169cf86-c080-4a4d-bc44-8155b0cebb2c")!=-1)
        {
          compareOneCommand.visible = event.selectedRows.length >0;
        }
        else
        {
          compareOneCommand.visible = false;
        }
      }      
    }
  }


  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
         selItems=`[`;         
         lElso=true;
         bootoast.toast({
          message: 'A TIGek feladás SAP rendszerbe elkezdődött!',
          type: 'info',
          position: 'rightTop',
          timeout: 5,
        });
        var i:  number = 0;
        var db: number = event.selectedRows.length;

         event.selectedRows.forEach((row: RowAccessor, index: number) => {
         const itemID = event.selectedRows[index].getValueByName("ID");           
         if (lElso)
         {
            selItems=`${selItems}"${itemID}"`;
            lElso=false;
         }
         else
         {
            selItems=`${selItems}, "${itemID}"`;
         }
         i=i+1;
         if ((i % 10)==0)
         {            
            var perc : number = Math.round((i/db)*10000)/100;
            bootoast.toast
            ({
              message: `Feldolgozottság: ${perc} %`,
              type: 'info',
              position: 'rightTop',
              timeout: 5,            
            });          
         }
         index=index+1;
        });  
        selItems=`${selItems}]`;      
        this.insertListItem(); 
        break;              
      default:
        throw new Error('Unknown command');
    }
  }

  private startMSFlow(itemID: number): Promise<HttpClientResponse> {
    if (itemID > 0) {
      const postURL = "https://prod-91.westeurope.logic.azure.com/workflows/46726babf8cf42cba56dceb1fca9cda5/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=acVNKXhwPdJ2c9PImPbgBRelbH8OlShEXZrY0eHWw6A";

      const body: string = JSON.stringify({
        'itemID': itemID.toString()
      });

      const requestHeaders: Headers = new Headers();
      requestHeaders.append('Content-type', 'application/json');

      const httpClientOptions: IHttpClientOptions = {
        body: body,
        headers: requestHeaders
      };

      console.log("About to make REST API request.");

      return this.context.httpClient.post(
        postURL,
        HttpClient.configurations.v1,
        httpClientOptions)
        .then((value: HttpClientResponse): Promise<HttpClientResponse> => {
          console.log("REST API response received.");
          return value.json();
        }, (error: any) => {
          console.log("Post error:" + error.message);
        });
    }
  }

  public async insertListItem()
  {
    let dateStr = new Date().toLocaleString();
    await sp.web.lists.getByTitle("SAP Export").items.add({
      Title: dateStr,
      Allapot: "Rögzítve",
      TIGDokumentumID: selItems
    }).then((_result: IItemAddResult)=> {    
      this.startMSFlow(_result.data.ID);
          }).then(() => {
            bootoast.toast({
              message: 'A TIGek feladása a SAP rendszerbe megtörtént!',
              type: 'success',
              position: 'rightTop',
              timeout: 5,
            });      
    })
    .catch((err: Error) => {
      bootoast.toast({
        message: `Hiba a feladás során: ${Error}`,
        type: 'danger',
        position: 'rightTop',
        timeout: 5,            
      });               
    });
  }

}

