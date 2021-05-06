import { sp } from "@pnp/sp";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { 
  IItemUpdateResult,
  IItemAddResult
} from "@pnp/sp/items";

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

import * as strings from 'TigvisszavoncaCommandSetStrings';

import ColorPickerDialog from './ColorPickerDialog';
import { IColor } from 'office-ui-fabric-react';
import './Ext/bootoast.css';
import './Ext/bootstrap.css';
import * as jQuery from 'jquery';


export interface ITigvisszavoncaCommandSetProperties {
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'TigvisszavoncaCommandSet';
const itemArray: string[]=[];
var selItems: string="";
var lElso: boolean=false;
var bootoast:any=require('./Ext/bootoast.js');


export default class TigvisszavoncaCommandSet extends BaseListViewCommandSet<ITigvisszavoncaCommandSetProperties> {
  private _longText: string;
  @override
  public onInit(): Promise<void> {        
    Log.info(LOG_SOURCE, 'Initialized TigvisszavoncaCommandSet');    
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
      if (compareOneCommand) {
        // This command should be hidden unless exactly one row is selected.
        /*if (window.location.href.toLowerCase().indexOf("tigdokumentumok/forms/tigviszzavonhato.aspx")!=-1)
        {
          if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=")==-1)
          {
              compareOneCommand.visible = event.selectedRows.length >0;
          }
          else
          {
            if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=8474d7f6-544d-4409-981f-b6e687637353")!=-1)
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
          if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=8474d7f6-544d-4409-981f-b6e687637353")!=-1)
          {
            compareOneCommand.visible = event.selectedRows.length >0;
          }
          else
          {
            compareOneCommand.visible = false;
          }
        } */
        
        if ((this.rightView("tigviszzavonhato","8474d7f6-544d-4409-981f-b6e687637353"))||(this.rightView("konyvelesrevar","b67a5718-d936-4935-82ad-76439779430e")))
        {
          compareOneCommand.visible = event.selectedRows.length >0;
        }
        else
        {
          compareOneCommand.visible = false;
        } 

      }
    }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
         /* Dialog.alert(`${this.properties.sampleTextOne}`);
          const dialog: ColorPickerDialog = new ColorPickerDialog();
          dialog.message = 'Add meg a visszaadás okát:';
          dialog.rows = "300";
          // Use 'FFFFFF' as the default color for first usage
                  
          dialog.show().then(() => {
            this._longText = dialog.longText;
            Dialog.alert(`A megjegyzés: ${dialog.longText}`);
          });*/

         /* this.emptyArray();
          const dialog: ColorPickerDialog = new ColorPickerDialog();  
            dialog.itemUrlFromExtension = event.selectedRows[0].getValueByName("FileRef");  
            dialog.otherParam = "This is parameter passed from Extension"  
              
            dialog.show().then(() => {                
              this._longText = dialog.paramFromDialog;
              if (this._longText!='XXXXXMégseXXXXX')
              {
                event.selectedRows.forEach((row: RowAccessor, index: number) => {
                    const itemID = event.selectedRows[index].getValueByName("ID");           
                    itemArray.push(itemID);            
                    index=index+1;
                });        
                this.updateListItemAll(); 
              }
              else
              {
                 alert('Nem történt meg a visszaadás'); 
              }              
            });                        
          break;*/
          /*event.selectedRows.forEach((row: RowAccessor, index: number) => {
            const itemID = event.selectedRows[index].getValueByName("ID");           
            itemArray.push(itemID);            
            index=index+1;
          });        
          this.updateListItemAll();          
          break;        */

          const dialog: ColorPickerDialog = new ColorPickerDialog();  
          dialog.itemUrlFromExtension = event.selectedRows[0].getValueByName("FileRef");  
          dialog.otherParam = "This is parameter passed from Extension";  
            
          dialog.show().then(() => 
          {                
            this._longText = dialog.paramFromDialog;         
            if ((this._longText!='XXXXXMégseXXXXX')&&(this._longText!=''))
            {
              bootoast.toast({
                message: 'A TIGek visszavonása elkezdődött!',
                type: 'info',
                position: 'rightTop',
                timeout: 5,
              });
              selItems=`[`;         
              lElso=true;
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
              bootoast.toast({
                message: 'A TIGek visszavonása megtörtént!',
                type: 'success',
                position: 'rightTop',
                timeout: 5,
              }); 
            }
            else
            {
              bootoast.toast
              ({
                message: 'A TIGek visszavonása nem történt meg!',
                type: 'danger',
                position: 'rightTop',
                timeout: 5, 
              });              
            }  
          });  
          break;
        default:
          throw new Error('Unknown command');
      }
    }
    
    public async insertListItem()
    {
      var sViewName:string=this.viewName("tigviszzavonhato","8474d7f6-544d-4409-981f-b6e687637353","Visszavonható");
      sViewName=sViewName+this.viewName("konyvelesrevar","b67a5718-d936-4935-82ad-76439779430e","Könyvelésre vár");
      let dateStr = new Date().toLocaleString();
      await sp.web.lists.getByTitle("TIG Jóváírás").items.add({
        Title: "Storno",
        Allapot: "Rögzítve",
        Megjegyzes: this._longText,
        TIGDokumentumID: selItems, 
        Nezet: sViewName,
      }).then((_result: IItemAddResult)=> {          
      })
      .catch((err: Error) => {
        bootoast.toast({
          message: `Hiba a visszavonás során: ${err.message}`,
          type: 'danger',
          position: 'rightTop',
          timeout: 5,            
        });               
      });
          
      /*Dialog.alert('A TIG helyesbítés beállítása elkezdődött!');        */
    }

    public viewName(viewName:string,viewID:string,retName:string): string
    {
      var ret:string="";
      if (window.location.href.toLowerCase().indexOf("tigdokumentumok/forms/"+viewName+".aspx")!=-1)
      {
        if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=")==-1)
        {
          ret=retName;
        }
        else
        {
          if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid="+viewID)!=-1)
          {
            ret=retName;
          }          
        }
      }
      else
      {
        if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid="+viewID)!=-1)
        {
          ret=retName;
        }       
      }
      return ret;
    }

    public rightView(viewName:string,viewID:string): boolean
    {
      var ret:boolean=false;
      if (window.location.href.toLowerCase().indexOf("tigdokumentumok/forms/"+viewName+".aspx")!=-1)
      {
        if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=")==-1)
        {
          ret=true;
        }
        else
        {
          if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid="+viewID)!=-1)
          {
            ret=true;
          }          
        }
      }
      else
      {
        if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid="+viewID)!=-1)
        {
          ret=true;
        }       
      }
      return ret;
    }

    public async updateListItemAll()
    {
      let list = await sp.web.lists.getByTitle("TIGDokumentumok");
      for (var v in itemArray) 
      {  
           var y: number = +itemArray[v];         
           await sp.web.lists.getByTitle("TIGDokumentumok").items.getById(y).update({Allapot: "Visszavonás alatt",
            VisszavonasIndoklas:this._longText
            }).then((_result: IItemUpdateResult)=> {          
            })
            .catch((err: Error) => {
              Dialog.alert(`Hibás update: $y')}`);              
            });
  
      } 
      this.emptyArray();

      Dialog.alert('A TIGek visszavonás allati állapotba kerültek!');        
    }          

    public emptyArray() {      
      itemArray.length = 0;
    }

  }
