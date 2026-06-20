import * as React from 'react';
import { ISharePointFormColumns } from './SharePointColumns';
import { IDropdownOption } from '@fluentui/react';
export const handleSkillsChange=(options:IDropdownOption,formData:ISharePointFormColumns,setFormData:React.Dispatch<React.SetStateAction<ISharePointFormColumns>>)=>{
const selectedkey=options.selected?[...formData.Skills,options?.key as string]:
formData.Skills.filter((key:any)=>key!==options);
setFormData(prev=>({...prev,Skills:selectedkey}));

}

// [a,b,c,d]=>[c,d]