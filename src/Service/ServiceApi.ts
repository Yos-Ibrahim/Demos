import {Web} from "@pnp/sp/presets/all";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { ListName } from "../Enum/ListNames";
import { ISharePointFormColumns } from "../Commonmethods/SharePointColumns";

export class ServiceApiClass{
    private web;
    constructor(siteurl:string){
      this.web=Web(siteurl);  
    }

public async addListItems(formData:ISharePointFormColumns):Promise<any>{
try{
const listName=this.web.lists.getByTitle(ListName.SharePointFormList); //here the list name is coming from the enum file
const result=await listName.items.add({
    Title:formData.Name,
    EmailAddress:formData.Email,
    Age:parseInt(formData.Age),
    Salary:parseFloat(formData.Compensation),
    Score:formData.Experience,
    FullAddress:formData.Address,
    Permission:formData.Permission,
    AdminId:formData.AdminId,
    ManagerId:{results:formData.ManagerId},
    Department:formData.Department,
    Gender:formData.Gender,
    CityId:formData.City,
    Skills:{results:formData.Skills}

});
return result
}
catch(err){
console.log("Error while adding list item",err);
}
}

}