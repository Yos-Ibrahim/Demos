import { WebPartContext } from "@microsoft/sp-webpart-base";
import { ListName } from "../Enum/ListNames";

export default class ChoiceService{

    private context:WebPartContext;
    constructor(context:WebPartContext){
        this.context=context
    }
//reading the choice
public async getChoiceValues(siteurl:string,fieldvalue:any):Promise<any>{
    try{

        const response=await fetch(`${siteurl}/_api/web/lists/getbytitle('${ListName.SharePointFormList}')/fields/?$filter=EntityPropertyName eq '${fieldvalue}'`,{
            method:'GET',
            headers:{
                'Accept':'application/json;odata=nometadata'

            }
                
            
        });

    if(!response.ok){
        throw new Error(`Error while fetching the choice values ${response.statusText}`);
    }
    const data=await response.json();
    const choice=data.value[0].Choices;
    return choice.map((item:any)=>({
        key:item,
        text:item
    }));
    }
    catch(err){
console.log(err);
throw err;
    }
}

public async getLookupValues():Promise<any>{
    try{
const response=await fetch(`${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${ListName.LookupList}')/items?$select=Title,ID`,{
    method:'GET',
    headers:{
       'Accept':'application/json;odata=nometadata' 
    }
});
 if(!response.ok){
        throw new Error(`Error while fetching the choice values ${response.statusText}`);
    }
    const data=await response.json();
    return data.value.map((city:{Title:string,ID:string})=>({
        key:city.ID,
        text:city.Title
    }));
    }
    catch(err){
console.log(err);
throw err;
    }
}
}