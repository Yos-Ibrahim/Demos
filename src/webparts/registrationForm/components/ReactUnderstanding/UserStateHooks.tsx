import { PrimaryButton } from '@fluentui/react';
import * as React from 'react';
import { useState } from 'react';

const UseStateHooks:React.FC<{}>=(props)=>{

    const [count,setCount]=useState<number>(0);
    return(
        <> 
        <p>Count of the Value:{count}</p>  
        <PrimaryButton
        label='Count'
        onClick={()=>setCount(count+1)}
        iconProps={{iconName:'add'}} 
        />
        
        </>
    )
    }
    export default UseStateHooks

    

