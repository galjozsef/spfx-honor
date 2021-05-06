import { sp } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemUpdateResult } from "@pnp/sp/items";

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
import * as jQuery from 'jquery';
import pnp from 'sp-pnp-js';

import * as strings from 'JovahagyCaCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IJovahagyCaCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'JovahagyCaCommandSet';
const itemArray: string[]=[];
var selItems: string="";
var lElso: boolean=false;
var bootoast:any=require('./Ext/bootoast.js');

export default class JovahagyCaCommandSet extends BaseListViewCommandSet<IJovahagyCaCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized JovahagyasCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      if (window.location.href.toLowerCase().indexOf("termekcikk/nyitott.aspx")!=-1)
      {
        if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=")==-1)
        {
            compareOneCommand.visible = event.selectedRows.length >0;
        }
        else
        {
          if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=4d493767-f7a3-49bb-905c-0ed1cfe1ed85")!=-1)
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
        if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=4d493767-f7a3-49bb-905c-0ed1cfe1ed85")!=-1)
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
          event.selectedRows.forEach((row: RowAccessor, index: number) => {
          const itemID = event.selectedRows[index].getValueByName("ID");
          //alert(`Field ID: ${row.getValueByName('ID')} - Field title: ${row.getValueByName('Title')}`);
          itemArray.push(itemID);
          /*if (this.updateListItem(itemp ID))
          {            
          }
          else
          {
            Dialog.alert(`Hibás update: Field ID: ${row.getValueByName('ID')} - Field title: ${row.getValueByName('Title')}`);  
          } */           
          index=index+1;
        });        
        this.updateListItemAll();
        /*this.UpdateStatus(event.selectedRows,"Jóváhagyott")*/
        //Dialog.alert('Jóváhagyás megtörtént!');  
        //window.location.reload(false);
         
        break;        
      default:
        throw new Error('Unknown command');
    }
  }

  /*public async updateListItem(itemID: any)  {    
    let list = sp.web.lists.getByTitle("Termékcikk");
    const i =  await list.items.getById(itemID).update({      
      Allapot: "Jóváhagyott",     
    });
    }
   */ 
  
 /* private UpdateStatus(itemcollection:any,columnValue:string)

{
  let b=pnp.sp.createBatch();
  itemcollection.forEach(item => {
    pnp.sp.web.lists.getByTitle("Termékcikk").items.getById(item.getValueByName("ID")).inBatch(b).update({Allapot:columnValue}).then(w=>{
    });

  });
  b.execute().then(w=>{
    location.reload();
    Dialog.alert('Végeztem!');    
  });
}*/

  public async updateListItemAll()
  {
    bootoast.toast({
      message: 'A jóváhagyás elkezdődött!',
      type: 'info',
      position: 'rightTop',
      timeout: 5,
    });

    let list = sp.web.lists.getByTitle("Termékcikk");
    const entityTypeFullName = await list.getListItemEntityTypeFullName();

    let batch = sp.web.createBatch();
    var i:number=0;
    var db:number=itemArray.length;
    for (var v in itemArray) 
    { 
      /*i=i+1;
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
      } */    
      var y: number = +itemArray[v];   
      list.items.getById(y).inBatch(batch).update({ Allapot: "Jóváhagyott" }, "*", entityTypeFullName).then(b => {
        console.log(b);
      })
      .catch((err: Error) => {
        bootoast.toast({
          message: `Hiba a jóváhagyás során: ${Error}`,
          type: 'danger',
          position: 'rightTop',
          timeout: 5,            
        });                                 
      });  
    }
    
    this.emptyArray();

    bootoast.toast({
      message: 'Az adatok kötegelt feldolgozása folyamatban..!',
      type: 'info',
      position: 'rightTop',
      timeout: 5,
    });

    await batch.execute().then(w=>{     
      bootoast.toast({
        message: 'A jóváhagyás befejeződött!',
        type: 'info',
        position: 'rightTop',
        timeout: 5,
      });
    })
    .catch((err: Error) => {
        bootoast.toast({
          message: `Hiba a jóváhagyás során: ${Error}`,
          type: 'danger',
          position: 'rightTop',
          timeout: 5,            
        });  
    }); 

    /*
    let list = await sp.web.lists.getByTitle("Termékcikk");
    var i:number=0;
    var db:number=itemArray.length;
    for (var v in itemArray) // for acts as a foreach  
    {  
         var y: number = +itemArray[v];         
         await sp.web.lists.getByTitle("Termékcikk").items.getById(y).update({Allapot: "Jóváhagyott"
          }).then((_result: IItemUpdateResult)=> {  
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
          })
          .catch((err: Error) => {
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
            bootoast.toast({
              message: `Hiba a jóváhagyás során: ${Error.toString()}`,
              type: 'danger',
              position: 'rightTop',
              timeout: 20,            
            });              
          });

    }    
    this.emptyArray();
    bootoast.toast({
      message: 'A jóváhagyás megtörtént!',
      type: 'success',
      position: 'rightTop',
      timeout: 60,
    }); 
    //Dialog.alert('A jóváhagyás megtörtént!');  */      
  }
  
  public emptyArray() {      
    itemArray.length = 0;
  }


  /*public async updateListItem(itemID: any): Promise<boolean> {
    let list = sp.web.lists.getByTitle("Termékcikk");
    return await sp.web.lists.getByTitle("Termékcikk").items.getById(itemID).update({
      Allapot: "Jóváhagyott",     
    }).then(async (_result: IItemUpdateResult): Promise<boolean> => {      
      return true;
    }, (_error: any): boolean => false);
  }*/

}



