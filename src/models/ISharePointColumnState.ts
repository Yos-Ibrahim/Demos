export interface ISharePointFormColumns{

Name:string;
Email:string;
Age:any;
Compensation:any;
Experience:number;
Address:string;

Permission:boolean;
Admin: string;
AdminId:any;
Manager:string[];
ManagerId:any[]; 
Department:string;
Skills:any[];
Gender:string;
City:string;
}

export interface IPaginatedTableState{
    key:number;
    Title:string;
    EmailAddress:string;
    Age:number;
    City:string;
    Admin:string;
}