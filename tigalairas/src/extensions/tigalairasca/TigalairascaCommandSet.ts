import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { Web } from 'sp-pnp-js';
import { IItemAddResult, IItemUpdateResult } from "@pnp/sp/items";
import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  RowAccessor,
  IListViewCommandSetListViewUpdatedParameters,
  IListViewCommandSetExecuteEventParameters
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import CustomDialog from './CustomDialog';
import './Ext/bootoast.css';
import './Ext/bootstrap.css';
import * as jQuery from 'jquery';
import * as strings from 'TigalairascaCommandSetStrings';
import { HttpClient, IHttpClientOptions, HttpClientResponse } from '@microsoft/sp-http';


export interface ITigalairascaCommandSetProperties {
  sampleTextOne: string;
  sampleTextTwo: string;
}

const LOG_SOURCE: string = 'TigalairascaCommandSet';
var selItems: string = "";
var lElso: boolean = false;
var bootoast: any = require('./Ext/bootoast.js');

export default class TigalairascaCommandSet extends BaseListViewCommandSet<ITigalairascaCommandSetProperties> {
  private param1Text: string;
  private param2Text: string;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized TigalairascaCommandSet');
    return Promise.resolve();
  }

  @override
  public onListViewUpdated(event: IListViewCommandSetListViewUpdatedParameters): void {
    const compareOneCommand: Command = this.tryGetCommand('COMMAND_1');
    if (compareOneCommand) {
      // This command should be hidden unless exactly one row is selected.
      if (window.location.href.toLowerCase().indexOf("tigdokumentumok/forms/generalva.aspx") != -1) {
        if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=") == -1) {
          compareOneCommand.visible = event.selectedRows.length > 0;
        }
        else {
          if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=fda88453-f01d-4153-b83f-f9e04f57e2a2") != -1) {
            compareOneCommand.visible = event.selectedRows.length > 0;
          }
          else {
            compareOneCommand.visible = false;
          }
        }
      }
      else {
        if (decodeURIComponent(window.location.href).toLowerCase().indexOf("viewid=fda88453-f01d-4153-b83f-f9e04f57e2a2") != -1) {
          compareOneCommand.visible = event.selectedRows.length > 0;
        }
        else {
          compareOneCommand.visible = false;
        }
      }
    }
  }

  @override
  public onExecute(event: IListViewCommandSetExecuteEventParameters): void {
    switch (event.itemId) {
      case 'COMMAND_1':
        selItems = `[`;
        lElso = true;
        var i: number = 0;
        var db: number = event.selectedRows.length;

        const dialog: CustomDialog = new CustomDialog();
        dialog.itemUrlFromExtension = event.selectedRows[0].getValueByName("FileRef");
        dialog.otherParam = "This is parameter passed from Extension";

        dialog.show().then(() => {
          this.param1Text = dialog.param1FromDialog;
          this.param2Text = dialog.param2FromDialog;
          if (this.param1Text != 'XXXXXMégseXXXXX') {
            bootoast.toast({
              message: 'A TIGek aláírása elkezdődött!',
              type: 'info',
              position: 'rightTop',
              timeout: 5,
            });
            event.selectedRows.forEach((row: RowAccessor, index: number) => {
              const itemID = event.selectedRows[index].getValueByName("ID");
              if (lElso) {
                selItems = `${selItems}"${itemID}"`;
                lElso = false;
              }
              else {
                selItems = `${selItems}, "${itemID}"`;
              }
              i = i + 1;
              if ((i % 10) == 0) {
                var perc: number = Math.round((i / db) * 10000) / 100;
                bootoast.toast
                  ({
                    message: `Feldolgozottság: ${perc} %`,
                    type: 'info',
                    position: 'rightTop',
                    timeout: 5,
                  });
              }
              index = index + 1;
            });
            selItems = `${selItems}]`;
            this.insertListItem();
          }
          else {
            bootoast.toast
              ({
                message: `Aláírás megszakítva!`,
                type: 'warning',
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

  private startMSFlow(itemID: number): Promise<HttpClientResponse> {
    if (itemID > 0) {
      const postURL = "https://prod-118.westeurope.logic.azure.com/workflows/c522e2df8e104d5e83c469eef9f2ce3b/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=UUBSKeWFqWbikXhdrEvkNvYKAflUh3sPi0_E15tbav4";

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


  private async insertListItem() {
    let dateStr = new Date().toLocaleString();
    let web = new Web(this.context.pageContext.site.absoluteUrl);
    await
      web.currentUser.get().then(result => {
        sp.web.lists.getByTitle("TIG Honor aláírás").items.add({
          Title: dateStr,
          Allapot: "Feldolgozásra vár",
          AlairasHonapja: this.param1Text,
          FizetesiMod: this.param2Text,
          TIGDokumentumID: selItems,
          ExportIdopontja: dateStr,
          InditottaId: result.Id
        }).then((_result: IItemAddResult) => {
          this.startMSFlow(_result.data.ID);
          }).then(() => {
            bootoast.toast({
              message: 'A TIGek aláírása megtörtént!',
              type: 'success',
              position: 'rightTop',
              timeout: 5,
            });
        })
          .catch((err: Error) => {
            bootoast.toast({
              message: `Hiba az aláírás során: ${Error}`,
              type: 'danger',
              position: 'rightTop',
              timeout: 5,
            });
          });
      }).catch((err: Error) => {
        bootoast.toast({
          message: `Hiba az aláírás során: ${Error}`,
          type: 'danger',
          position: 'rightTop',
          timeout: 5,
        });
      });

  }
}