import { ChoiceGroup, ComboBox, Dropdown, PrimaryButton, Slider, TextField, Toggle } from '@fluentui/react';
import * as React from 'react';
import styles from '../RegistrationForm.module.scss';

//functional component
const BasicForm:React.FC<{}>=()=>{
    
return(
    <>
   {/*  <h3 style={{color:"blue",textAlign:"center",textDecoration:"underline"}}>
        Basic Form</h3> */}

        
    <p className={styles.h4}> Basic Form </p> 
    <hr/>
   <form>
    <TextField label='Name'
    placeholder='Enter your name here..'
    required
    iconProps={{iconName:'people'}} 
    />

    <TextField label='Email Address'
    placeholder='Enter your Email here..'
    iconProps={{iconName:'mail'}} 
    />

    <TextField label='compensation'
   prefix='$'
   suffix='USD'
    />

    <TextField label='Password'
    type='password'
    canRevealPassword={true}
    />

    <TextField
    label='File'
    type='file'
    />

    <Dropdown 
    label='Department'
    options={[
        {key:"IT",text:"IT"},
        {key:"HR",text:"HR"}
    ]}
    multiSelect
    />

    <ChoiceGroup
    label='Gender'
    options={[
        {key:"Male",text:"Male"},
         {key:"Female",text:"Female"}
    ]}
    />
    <ComboBox
    label='combobox Department'
    options={[
        {key:"IT",text:"IT"},
        {key:"HR",text:"HR"}
    ]}
    multiSelect
    autoComplete='on'
    allowFreeform
    />

    <Slider
    label='Experience'
    min={0}
    max={25}
    step={0.1}
  />
    <Toggle
    label="Permission"
    />
  <TextField 
  label='address'
    multiline
    rows={5}

/>
<br/>


<PrimaryButton
text="Save"
iconProps={{iconName:"Save"}}
onClick={()=>alert("I am save button")}
/>
&nbsp;
&nbsp;
&nbsp;

<PrimaryButton
text="Cancel"
iconProps={{iconName:"Cancel"}}
onClick={()=>alert("I am cancel button")}
/>
   </form>
    </>
)
}
export default BasicForm