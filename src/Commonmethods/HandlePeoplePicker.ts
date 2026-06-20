import * as React from 'react';
import { ISharePointFormColumns } from '../Commonmethods/SharePointColumns';

export const HandleSingleSelectedPeoplePicker=(items:any[], 
    setFormdata:React.Dispatch<React.SetStateAction<ISharePointFormColumns>>)=>{
if(items.length>0){
    setFormdata(prev=>({...prev,Admin:items[0].text,AdminId:items[0].id}));
}
else{
    setFormdata(prev=>({...prev,Admin:"",AdminId:0}));
}

}

    export const HandleMultiSelectedPeoplePicker=(items:any[], 
    setFormdata:React.Dispatch<React.SetStateAction<ISharePointFormColumns>>)=>{

    setFormdata(prev=>({...prev,Manager:items.map(i=>i.text),ManagerId:items.map(i=>i.id)}));

    }