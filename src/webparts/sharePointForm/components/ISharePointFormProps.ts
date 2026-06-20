import { WebPartContext } from "@microsoft/sp-webpart-base";


export interface ISharePointFormProps {
  context:WebPartContext;
  siteurl:string;
  departmentoptions:any;
  genderoptions:any;
  skillsoptions:any;
  cityoptions:any
}
