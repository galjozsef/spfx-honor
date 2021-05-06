import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';  

export default class CustomDialog extends BaseDialog 
{  
    public itemUrlFromExtension: string;  
    public otherParam: string;  
    public param1FromDialog:string=`XXXXXMégseXXXXX`;  
    public param2FromDialog:string=`XXXXXMégseXXXXX`;    
    
    public render(): void 
    {  
        var html:string = "";  
          
        html +=  `<div style="padding: 10px;">`;  
        html +=  `<h3>Töltsd ki az adatokat!</h3>`;  
        html +=  `<table>`;
        html +=  `<tr>`; 
        html +=  `<td style="padding: 10px;" >`; 
        html +=  `Aláírás hónapja:`;             
        html +=  `</td>`;   
        html +=  `<td style="padding: 10px;" >`;     
        html +=  `<select name="AlairasHonapja" id="inputParam1">`;
        html +=  `<option value="tárgyhavi">tárgyhavi</option>`;
        html +=  `<option value="elhatárolt">elhatárolt</option>`;            
        html +=  `</select>`;
        html +=  `</td>`;   
        html +=  `</tr>`;        
        html +=  `<tr>`; 
        html +=  `<td style="padding: 8px;" >`; 
        html +=  `Fizetési mód:`;          
        html +=  `</td>`;   
        html +=  `<td style="padding: 8px;" >`; 
        html +=  `<select name="FizetesiMod" id="inputParam2">`;
        html +=  `<option value="banki átutalás">banki átutalás</option>`;           
        html +=  `</select>`;
        html +=  `</td>`;   
        html +=  `</tr>`; 
        html +=  `</table>`;
        html +=  `<br>`; 
        html +=  `<table>`; 
        html +=  `<tr>`; 
        html +=  `<td style="padding: 10px;" >`; 
        html +=  `<input style="padding: 5px;" type="button" id="OkButton"  value="TIG aláírás indítása">`;  
        html +=  `</td>`; 
        html +=  `<td style="padding: 10px;" >`; 
        html +=  `<input style="padding: 5px;" type="button" id="Cancel"    value="Mégse">`;  
        html +=  `</td>`; 
        html +=  `</tr>`; 
        html +=  `</table>`; 
        html +=  `</div>`;  
        this.domElement.innerHTML += html;  
        this._setButtonEventHandlers();    
    }  
    
    private _setButtonEventHandlers(): void 
    {    
        const webPart: CustomDialog = this;    
        this.domElement.querySelector('#OkButton').addEventListener('click', () => 
        {    
            this.param1FromDialog =  document.getElementById("inputParam1")["value"] ;   
            this.param2FromDialog =  document.getElementById("inputParam2")["value"] ;   
            this.close();  
        });   
        this.domElement.querySelector('#Cancel').addEventListener('click', () => 
        {    
            this.param1FromDialog = `XXXXXMégseXXXXX`;             
            this.param2FromDialog = `XXXXXMégseXXXXX`;             
            this.close();  
        });
    } 
    public getConfig(): IDialogConfiguration 
    {
      return {
      isBlocking: false
      };
    }
      
    protected onAfterClose(): void 
    {  
      super.onAfterClose();       
    }  
    
} 
