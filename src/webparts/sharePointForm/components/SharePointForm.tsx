import * as React from 'react';
import styles from './SharePointForm.module.scss';
import type { ISharePointFormProps } from './ISharePointFormProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { ServiceApiClass } from '../../../Service/ServiceApi';
import { handleSkillsChange } from '../../../Commonmethods/IMultiselectdropdown';
import { ISharePointFormColumns } from '../../../Commonmethods/SharePointColumns';

import { useCallback,useState } from 'react';
//import {sp} from "@pnp/sp/presets/all";
import { Dialog } from '@microsoft/sp-dialog';
import { ChoiceGroup, Dropdown, PrimaryButton, Slider, TextField, Toggle } from '@fluentui/react';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { HandleMultiSelectedPeoplePicker, HandleSingleSelectedPeoplePicker } from '../../../Commonmethods/HandlePeoplePicker';
const SharePointForm:React.FC<ISharePointFormProps>=(props)=>{
  const [formdata,setFormData]=useState<ISharePointFormColumns>({
      //initialization of the form 
    Name:"",
      Email:"",
      Age:"",
      Compensation:"",
       Experience:0,
       Address:"",
       Permission:false,
       Admin:"",
      AdminId:0,
      Manager:[],
      ManagerId:[],
      Department:"",
      Gender:"",
      Skills:[],
      City:""

      });

 /*  useEffect(()=>{
    sp.setup({
      spfxContext:props.context
    })
  },[props.context]); */

  const createItems=async()=>{
    try{
     const _service=new ServiceApiClass(props.siteurl);
     const result=await _service.addListItems(formdata);
      Dialog.alert("Item created successfully with ID: "+result.data.Id);
      console.log("Result after creating item",result);
      setFormData({
        Name:"",
        Email:"",
        Age:"",
        Compensation:"",
         Experience:0,
         Address:"",
         Permission:false,
         Admin:"",
        AdminId:0,
        Manager:[],
        ManagerId:[],
        Department:"",
        Gender:"",
        Skills:[],
        City:""
})
    }
    catch(err){
      console.log("Error while creating item",err);
    }
  }
  //form event 
  const handleChange=useCallback((field:keyof ISharePointFormColumns,value:string|number|boolean)=>{
setFormData(prev=>({...prev,[field]:value}))
  },[]);
  return(
    <>
    <TextField
    label='Name'
    value={formdata.Name}
    onChange={(_,val)=>handleChange("Name",val||"")}
    />
    <TextField
    label='Email Address'
    value={formdata.Email}
    onChange={(_,val)=>handleChange("Email",val||"")}
    iconProps={{iconName:"mail"}}
    />
    <TextField
    label='Age'
    value={formdata.Age}
    onChange={(_,val)=>handleChange("Age",val||"")}
    />
     <TextField
    label='Compensation'
    value={formdata.Compensation}
    onChange={(_,val)=>handleChange("Compensation",val||"")}
    prefix='$'
    suffix='USD'
    />
    <Slider
    label='Experience'
    value={formdata.Experience}
    onChange={(val)=>handleChange("Experience",val)}
    max={25}
    step={0.1}
    />
    
<Dropdown
    label='Department'
   options={props.departmentoptions}
   placeholder='--select--'
   selectedKey={formdata.Department}
  onChange={(_,option)=>handleChange("Department",option?.key as string)}
    />
    <Dropdown
    label='City'
   options={props.cityoptions}
   placeholder='--select--'
   selectedKey={formdata.City}
  onChange={(_,option)=>handleChange("City",option?.key as string)}
    />
    <Dropdown
    label='Skills'
   options={props.skillsoptions}
   placeholder='--select--'
   defaultSelectedKey={formdata.Skills}
  onChange={(_,opt)=>handleSkillsChange(opt!,formdata,setFormData)}
    multiSelect
  
    />
<ChoiceGroup
    label='Gender'
   options={props.genderoptions}
   selectedKey={formdata.Gender}
  onChange={(_,option)=>handleChange("Gender",option?.key as string)}
    />
    <TextField
    label='Full Address'
    value={formdata.Address}
    onChange={(_,val)=>handleChange("Address",val||"")}
   rows={5}
   multiline
    />
    <Toggle
    label='Permission'
    checked={formdata.Permission}
    onChange={(_,checked)=>handleChange("Permission",!!checked)}
    />

{/* single people picker*/}

    <PeoplePicker
    context={props.context as any}
    titleText="Admin"
    personSelectionLimit={1}
    //  groupName={"Team Site Owners"} : Leave this blank in case you want to filter from all users
    showtooltip={true}
    onChange={(items)=>HandleSingleSelectedPeoplePicker(items,setFormData)}
    principalTypes={[PrincipalType.User]}
    ensureUser={true}
    webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
    defaultSelectedUsers={[formdata.Admin?formdata.Admin:""]}
    resolveDelay={1000} 
    />
 <br/>
 <PeoplePicker
    context={props.context as any}
    titleText="Manager"
    personSelectionLimit={2}
    //  groupName={"Team Site Owners"} : Leave this blank in case you want to filter from all users
    showtooltip={true}
    onChange={(items)=>HandleMultiSelectedPeoplePicker(items,setFormData)}
    principalTypes={[PrincipalType.User]}
    ensureUser={true}
    webAbsoluteUrl={props.context.pageContext.web.absoluteUrl}
    defaultSelectedUsers={formdata.Manager}
    resolveDelay={1000} />
 <br/>
    <PrimaryButton
    text='Save'
    onClick={()=>createItems()}
    iconProps={{iconName:"save"}}
    />
    </>
  )
}

export default SharePointForm;