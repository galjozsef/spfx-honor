import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-users/web";
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
import pnp from 'sp-pnp-js';
import './Ext/bootoast.css';
import './Ext/bootstrap.css';
import * as jQuery from 'jquery';
import * as strings from 'JovahagyVisszaCaCommandSetStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IJovahagyVisszaCaCommandSetProperties {
  // This is an example; replace with your own properties
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'JovahagyVisszaCaCommandSet';
const itemArray: string[]=[];
var selItems: string="";
var lElso: boolean=false;
var bootoast:any=require('./Ext/bootoast.js');

export default class JovahagyVisszaCaCommandSet extends BaseListViewCommandSet<IJovahagyVisszaCaCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized JovahagyasVisszavonasaCommandSet');    
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    
    /*sp.web.currentUser.get().then((user) => {
      const newLocal = [];
      let groups: any[] = newLocal;    
      sp.web.siteUsers.getById(user.Id).groups.get()
      .then((groupsData) => {
                groupsData.forEach(group => {
                  groups.push({
                   Id: group.Id,
                        Title: group.Title
                      });
                  });
      });
    });*/
    
    if (compareOneCommand) 
    {
      // This command should be hidden unless exactly one row is selected.
      if (window.location.href.toLowerCase().indexOf("termekcikk/jvhagyott.aspx")!=-1)
      {
        if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=")==-1)
        {
            compareOneCommand.visible = event.selectedRows.length >0;
        }
        else
        {
          if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=ff1bf916-89b6-4734-a5f2-0cf8475d5df0")!=-1)
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
        if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=ff1bf916-89b6-4734-a5f2-0cf8475d5df0")!=-1)
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
           Dialog.alert(`Hib??s update: Field ID: ${row.getValueByName('ID')} - Field title: ${row.getValueByName('Title')}`);  
         } */           
         index=index+1;
       });        
       this.updateListItemAll();      
       break;        
     default:
       throw new Error('Unknown command');
   }
 }

 /*public async updateListItem(itemID: any)  {    
   let list = sp.web.lists.getByTitle("Term??kcikk");
   const i =  await list.items.getById(itemID).update({      
     Allapot: "J??v??hagyott",     
   });
   }
  */ 
 
/* private UpdateStatus(itemcollection:any,columnValue:string)

{
 let b=pnp.sp.createBatch();
 itemcollection.forEach(item => {
   pnp.sp.web.lists.getByTitle("Term??kcikk").items.getById(item.getValueByName("ID")).inBatch(b).update({Allapot:columnValue}).then(w=>{
   });

 });
 b.execute().then(w=>{
   location.reload();
   Dialog.alert('V??geztem!');    
 });
}*/

 public async updateListItemAll()
 {
  bootoast.toast({
    message: 'A j??v??hagy??s visszavon??sa elkezd??d??tt!',
    type: 'info',
    position: 'rightTop',
    timeout: 5,
  });

  let list = sp.web.lists.getByTitle("Term??kcikk");
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
        message: `Feldolgozotts??g: ${perc} %`,
        type: 'info',
        position: 'rightTop',
        timeout: 5,            
      });          
    } */    
    var y: number = +itemArray[v];   
    list.items.getById(y).inBatch(batch).update({ Allapot: "Nyitott" }, "*", entityTypeFullName).then(b => {
      console.log(b);
    })
    .catch((err: Error) => {
      bootoast.toast({
        message: `Hiba a j??v??hagy??s visszavon??sa sor??n: ${Error}`,
        type: 'danger',
        position: 'rightTop',
        timeout: 5,            
      });                                 
    });  
  }
  
  this.emptyArray();

  bootoast.toast({
    message: 'Az adatok k??tegelt feldolgoz??sa folyamatban..!',
    type: 'info',
    position: 'rightTop',
    timeout: 5,
  });

  await batch.execute().then(w=>{     
    bootoast.toast({
      message: 'A j??v??hagy??s visszavon??sa befejez??d??tt!',
      type: 'info',
      position: 'rightTop',
      timeout: 5,
    });
  })
  .catch((err: Error) => {
      bootoast.toast({
        message: `Hiba a j??v??hagy??s visszavon??sa sor??n: ${Error}`,
        type: 'danger',
        position: 'rightTop',
        timeout: 5,            
      });  
  }); 
   /*bootoast.toast({
      message: 'A j??v??hagy??s visszavon??sa elkezd??d??tt!',
      type: 'info',
      position: 'rightTop',
      timeout: 60,
   });
   let list = await sp.web.lists.getByTitle("Term??kcikk");
   var i:number=0;
   var db:number=itemArray.length;
   for (var v in itemArray) // for acts as a foreach  
   {  
        var y: number = +itemArray[v];         
        await sp.web.lists.getByTitle("Term??kcikk").items.getById(y).update({Allapot: "Nyitott"
         }).then((_result: IItemUpdateResult)=> {  
          i=i+1;
          if ((i % 10)==0)
          {            
            var perc : number = Math.round((i/db)*10000)/100;
            bootoast.toast
            ({
              message: `Feldolgozotts??g: ${perc} %`,
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
              message: `Feldolgozotts??g: ${perc} %`,
              type: 'info',
              position: 'rightTop',
              timeout: 5,            
            }); 
          }  
          bootoast.toast({
            message: `Hiba a j??v??hagy??s visszavon??sa sor??n: ${Error.toString()}`,
            type: 'danger',
            position: 'rightTop',
            timeout: 20,            
          });  
         });

   } 
   this.emptyArray();
    bootoast.toast({
      message: 'A j??v??hagy??s visszavon??sa megt??rt??nt!',
      type: 'success',
      position: 'rightTop',
      timeout: 60,
    });    
   //Dialog.alert('A j??v??hagy??s visszavon??sa megt??rt??nt!');
   */        
 }
 
  public emptyArray() {      
    itemArray.length = 0;
  }

}


