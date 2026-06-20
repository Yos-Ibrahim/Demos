import {sp}  from "@pnp/sp/presets/all";
import { IPaginatedTableState } from "../models/ISharePointColumnState";
import { ListNames } from "../Enum/ListNames";
export default class PaginatedServiceApiClass{

public static async getPaginatedItems():Promise<IPaginatedTableState[]>{
    
try{
let allItems:any[]=[];
let paged=await sp.web.lists.getByTitle(ListNames.PaginatedList).items.select("Id","Title","EmailAddress","Age","Admin/Title","City/Title")
.expand("Admin","City").top(4999).getPaged();

//first batch of items
allItems.push(...paged.results);
console.log(`Fetched ${paged.results.length} items`);

//contine fetching next batches
while(paged.hasNext){
    paged=await paged.getNext();
    allItems.push(...paged.results);
    console.log(`Fetched ${paged.results.length} items`);
}
console.log(`Total items fetched: ${allItems.length}`);

//return the same structure as IPaginatedTableState
return allItems.map((e:any)=>({
    key:e.Id,
    Title:e.Title,
    EmailAddress:e.EmailAddress,
    Age:e.Age,
    City:e.City?.Title || "",
    Admin:e.Admin?.Title || ""
}))
}
catch(err){
    console.error("Error fetching paginated items:", err);
    return [];
}

    }
}