import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';  
import './Ext/bootoast.css';
import './Ext/bootstrap.css';
var bootoast:any=require('./Ext/bootoast.js');

export default class CustomDialog extends BaseDialog 
{  
    public itemUrlFromExtension: string;  
    public otherParam: string;  
    public paramFromDialog:string=`XXXXXMégseXXXXX`;  
    
    public render(): void 
    {  
        var html:string = "";  
          
        html +=  `<div style="padding: 10px;">`;  
        html +=  `<h3>Add meg a visszavonás indoklását!</h3>`;  
        html +=  `Visszavonás indoklása:`;  
        html +=  `<br>`;          
        html +=  `<input type="text" id="inputParam" size="100">` + `</input>`;  
        html +=  `<br>`; 
        html +=  `<br>`; 
        html +=  `<table>`; 
        html +=  `<tr>`; 
        html +=  `<td style="padding: 10px;" >`; 
        html +=  `<input style="padding: 5px;" type="button" id="OkButton"  value="Visszavonás indítása">`;  
        html +=  `</td>`; 
        html +=  `<td style="padding: 10px;" >`; 
        html +=  `<input style="padding: 5px;" type="button" id="Cancel"    value="Mégse">`;  
        html +=  `</td>`; 
        html +=  `</tr>`; 
        html +=  `</table>`; 
        html +=  `</div>`;  
        this.paramFromDialog = `XXXXXMégseXXXXX`;     
        this.domElement.innerHTML += html;  
        this._setButtonEventHandlers();    
    }  
    
    private _setButtonEventHandlers(): void 
    {    
        const webPart: CustomDialog = this;    
        this.domElement.querySelector('#OkButton').addEventListener('click', () => 
        {    
            this.paramFromDialog =  document.getElementById("inputParam")["value"] ;   
            if (this.paramFromDialog=='')
            {
              bootoast.toast({
                message: 'A megjegyzés nem lehet üres!',
                type: 'danger',
                position: 'rightTop',
                timeout: 10,
              });
            }
            else
            {  
              this.close();  
            }
        });   
        this.domElement.querySelector('#Cancel').addEventListener('click', () => 
        {    
            this.paramFromDialog = `XXXXXMégseXXXXX`;             
            this.close();  
        });
    } 
    public getConfig(): IDialogConfiguration 
    {  
      return ;       
    }  
      
    protected onAfterClose(): void 
    {  
      super.onAfterClose();       
    }  
    
} 
