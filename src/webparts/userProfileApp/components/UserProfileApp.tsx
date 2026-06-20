import * as React from 'react';
import styles from './UserProfileApp.module.scss';
import type { IUserProfileAppProps } from './IUserProfileAppProps';
import { escape } from '@microsoft/sp-lodash-subset';
import {IUserProfile} from '../../../models/IUserProfile';
import{ AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
// WE NEDD DOO
const GRAPH_RESOURCE = "https://graph.microsoft.com";
//we need to provide ms grap url
const graphUrl = "https://graph.microsoft.com/v1.0/me";
const UserProfileApp: React.FC<IUserProfileAppProps> = (props) => {
  const [profile,setProfile]=React.useState<IUserProfile|any>(null);
  const [loading,setLoading]=React.useState<boolean>(false);
  const [error,setError]=React.useState<string|any>("");
  const [fetchProfile,setFetchProfile]=React.useState<boolean>(false);


  const fetchuserProfile=React.useCallback(async():Promise<void>=>{
setLoading(true);
setError(null);

try{
  //step 1 get the toke
  const client:AadHttpClient = await props.aadHttpClientFactory.getClient(GRAPH_RESOURCE);
  //end point url
  const response:HttpClientResponse = await client.get(graphUrl, AadHttpClient.configurations.v1);
//step 3 get the response
if(!response.ok){
  throw new Error(`Error fetching user profile: ${response.statusText}`);
}
//step 4 conver the response to json
const userProfileData:IUserProfile = await response.json();
//step 5 set the profile data to state
setProfile(userProfileData);
setFetchProfile(true);
}
catch(err:unknown){
  const mssg=err instanceof Error ? err.message : "An unknown error occurred";
  setError(mssg);
  
}
finally{
  setLoading(false);
}
  },[props.aadHttpClientFactory]);

  React.useEffect(()=>{
fetchuserProfile();

  },[fetchProfile])
  return(
    <>


  <div className={styles.container}>
<h2 className={styles.title}>👤 My Profle<span className={styles.badge}>AadHttpclient</span></h2>

{/* Loading state */}
{loading&&(
  <div className={styles.loadingBox}>
    <div className={styles.spinner}></div>
    <span >We are fetching data from Microsoft Graph Api</span>
    </div>
)}
{/* Error state */}
{error&&(
  <div className={styles.errorBox}>
    <strong>❌Error</strong>{error}
    <button className={styles.retryBtn} onClick={fetchuserProfile}>Retry</button>
    </div>
)}

{/* Profile data */}
{fetchProfile&&profile&&!loading&&(
  <div className={styles.profileCard}>
    {/* Avatar */}
    <div className={styles.avatar}>
      {profile.displayName?.charAt(0).toUpperCase()??'?'}
      </div>

  </div>
)}

{/* Information rows */}
<div className={styles.infoGrid}>
  <ProfileRow icon="👤" label="Name" value={profile?.displayName} />
  <ProfileRow icon="✉️" label="Email" value={profile?.mail} />
  <ProfileRow icon="💼" label="Job Title" value={profile?.jobTitle} />
  <ProfileRow icon="🏢" label="Department" value={profile?.department} />
  <ProfileRow icon="📍" label="Office Location" value={profile?.officeLocation} />
  <ProfileRow icon="📱" label="Mobile Phone" value={profile?.mobilePhone} />
</div>
{/*  */}

<details className={styles.rawJson}>
  <summary> 🔎Raw JSON</summary>
  <pre>{JSON.stringify(profile, null, 2)}</pre>
</details>
{/* manual */}
{fetchProfile&&!loading&&(
  <button className={styles.refreshBtn} onClick={fetchuserProfile}>
    🔄️ Refresh
  </button>
)}
</div>

    </>
  )
}

//small helpers 
interface IProfileRowProps{
  icon:string;
  label:string;
  value:string|null|undefined;
}
const ProfileRow:React.FC<IProfileRowProps> =({icon,label,value})=>{
  return <div className={styles.infoRow}>
    <div className={styles.infoIcon}>{icon}</div>
    <div className={styles.infoLabel}>{label}</div>
    <div className={styles.infoValue}>{value}</div>
  </div>;
}
export default UserProfileApp;