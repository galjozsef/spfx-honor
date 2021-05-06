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
import './Ext/bootoast.css';
import './Ext/bootstrap.css';
import * as jQuery from 'jquery';
import * as strings from 'TighelyesbitcaCommandSetStrings';


export interface ITighelyesbitcaCommandSetProperties { 
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'TighelyesbitcaCommandSet';
const itemArray: string[]=[];
var selItems: string="";
var lElso: boolean=false;
var bootoast:any=require('./Ext/bootoast.js');

export default class TighelyesbitcaCommandSet extends BaseListViewCommandSet<ITighelyesbitcaCommandSetProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized TighelyesbitcaCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
      if (compareOneCommand) {
        // This command should be hidden unless exactly one row is selected.
       /* if (window.location.href.toLowerCase().indexOf("tigdokumentumok/forms/konyvelesrevar.aspx")!=-1)
        {
          if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=")==-1)
          {
              compareOneCommand.visible = event.selectedRows.length >0;
          }
          else
          {
            if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=b67a5718-d936-4935-82ad-76439779430e")!=-1)
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
          if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=b67a5718-d936-4935-82ad-76439779430e")!=-1)
          {
            compareOneCommand.visible = event.selectedRows.length >0;
          }
          else
          {
            compareOneCommand.visible = false;
          }
        }*/
        
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
        bootoast.toast({
          message: 'A helyesbítés elkezdődött!',
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
          message: 'A TIG helyesbítés beállítása megtörtént!',
          type: 'success',
          position: 'rightTop',
          timeout: 5,
        });  
        break;     
      default:
        throw new Error('Unknown command');
    }
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

  public async insertListItem()
  {
    var sViewName:string=this.viewName("tigviszzavonhato","8474d7f6-544d-4409-981f-b6e687637353","Visszavonható");
    sViewName=sViewName+this.viewName("konyvelesrevar","b67a5718-d936-4935-82ad-76439779430e","Könyvelésre vár");
    let dateStr = new Date().toLocaleString();
    await sp.web.lists.getByTitle("TIG Jóváírás").items.add({
      Title: "Helyesbítés",
      Allapot: "Rögzítve",
      TIGDokumentumID: selItems, 
      Nezet: sViewName,
    }).then((_result: IItemAddResult)=> {          
    })
    .catch((err: Error) => {
      bootoast.toast({
        message: `Hiba a helyesbítés során: ${Error}`,
        type: 'danger',
        position: 'rightTop',
        timeout: 5,            
      });               
    });
        
    /*Dialog.alert('A TIG helyesbítés beállítása elkezdődött!');        */
  }

}
