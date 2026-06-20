import { PrimaryButton } from '@fluentui/react';
import * as React from 'react';
import { useState, useEffect } from 'react';

const UseEffectHooks:React.FC<{}>=(props)=>{

    const [count,setCount]=useState<number>(0);

    useEffect(()=>{
        console.log("Use effects is called");
    },[count])
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
    export default UseEffectHooks

    

