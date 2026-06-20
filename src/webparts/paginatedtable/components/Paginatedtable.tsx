import * as React from 'react';
// import styles from './Paginatedtable.module.scss';
import type { IPaginatedtableProps } from './IPaginatedtableProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { Dropdown, IDropdownOption, PrimaryButton } from '@fluentui/react';
import { useState,useEffect } from 'react';
import PaginatedServiceApiClass from '../../../Service/PaginatedServiceApi';
import jsPDF from 'jspdf';
import 'jspdf-autotable';
import autoTable from 'jspdf-autotable';
import * as XLSX from 'xlsx';
import {  Input,Table } from 'antd';
import type { ColumnType, FilterDropdownProps } from 'antd/es/table/interface';


const pagintedItems:IDropdownOption[]=[
  {key:5,text:"5 items"},
  {key:10,text:"10 items"},
  {key:15,text:"15 items"},
  {key:20,text:"20 items"},
]

const  Paginatedtable:React.FC<IPaginatedtableProps>=(props)=>{
  const [allItems,setAllItems]=useState<any[]>([]);
  const [searchText,setSearchText]=useState<string>("");
  const [loading,setLoading]=useState<boolean>(false);
  const [pageSize,setPageSize]=useState<number>(5);
  const [page,SetPage]=useState<number>(1);

  useEffect(()=>{
    
    const loadList=async()=>{
      setLoading(true);
      const response =await  PaginatedServiceApiClass.getPaginatedItems();
      setAllItems(response);
      setLoading(false);
    }
    loadList();
  },[]);

//search function for each column

const getColumnSearchProps = (dataIndex: string): ColumnType<any> => ({
  filterDropdown: (props: FilterDropdownProps) => {
    const { setSelectedKeys, selectedKeys, confirm, clearFilters } = props;

    return (
      <div style={{ padding: 8 }}>
        <Input
          placeholder={`Search ${dataIndex}`}
          value={selectedKeys[0]}
          onChange={(e) =>
            setSelectedKeys(e.target.value ? [e.target.value] : [])
          }
          onPressEnter={() => confirm()}
          style={{ marginBottom: 8, display: 'block' }}
        />

        <button
          onClick={() => confirm()}
          style={{ marginRight: 8 }}
        >
          Search
        </button>

        <button
          onClick={() => clearFilters && clearFilters()}
        >
          Reset
        </button>
      </div>
    );
  },

  filterIcon: (filtered: boolean) => (
    <span style={{ color: filtered ? '#1677ff' : undefined }}>🔍</span>
  ),

  onFilter: (value, record) =>
    (record[dataIndex] ?? '')
      .toString()
      .toLowerCase()
      .includes((value as string).toLowerCase()),
});




//end of search function

  //to check any undefined or empty values
const filteredItems=allItems.filter((items:any)=>items?.Title?.toLowerCase().includes(searchText.toLowerCase())
||items?.EmailAddress?.toLowerCase().includes(searchText.toLowerCase())||
items?.Admin?.toLowerCase().includes(searchText.toLowerCase())||
items?.City?.toLowerCase().includes(searchText.toLowerCase())||
items?.Age?.toString().includes(searchText)

);

//pagination logic
const paginatedItems=filteredItems.slice((page-1)*pageSize ,page*pageSize);

//page =2
//pageszie=10
//start index=(2-1)*10=10
//end index=2*10=20
//filert(10,20)

const ageFilters = Array.from(new Set(allItems.map(i => i.Age))).map(v => ({
  text: v,
  value: v
}));

const cityFilters = Array.from(new Set(allItems.map(i => i.City))).map(v => ({
  text: v,
  value: v
}));

const columns=[
  {
    title:"Name",
    dataIndex:"Title",
    key:"Title",
    sorter:(a:any,b:any)=>(a.Title||"").localeCompare(b.Title||""),
    ...getColumnSearchProps("Title"),
  },
  {
     title:"Email Address",
    dataIndex:"EmailAddress",
    key:"EmailAddress",
    sorter:(a:any,b:any)=>(a.EmailAddress||"").localeCompare(b.EmailAddress||""),
    ...getColumnSearchProps("EmailAddress"),
  },
  {
title:"Age",
dataIndex:"Age",
key:"Age",
sorter:(a:any,b:any)=>(a.Age||0)-(b.Age||0),
    filters: ageFilters,
    onFilter: (value:any, record:any) => record.Age === value,
  },
  {
    title:"Admin",
    dataIndex:"Admin",
    key:"Admin",
    sorter:(a:any,b:any)=>(a.Admin||"").localeCompare(b.Admin||""),
    ...getColumnSearchProps("Admin"),
  },
  {
    title:"City",
    dataIndex:"City",
    key:"City",
    sorter:(a:any,b:any)=>(a.City||"").localeCompare(b.City||""),
    filters: cityFilters,
    onFilter: (value:any, record:any) =>
      (record.City ?? "").toString().toLowerCase() ===
      (value as string).toLowerCase(),
  },
];

//search box
const handleSearch=(e:any)=>setSearchText(e.target.value);

//export to excel
const exportExcel=()=>{
  const worksheet=XLSX.utils.json_to_sheet(filteredItems);
  const workbook=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook,worksheet,"SPListData");
  XLSX.writeFile(workbook,"SPListdata.xlsx");
}

//export to pdf
const exporttopdf=()=>{
  const doc=new jsPDF();
  doc.text("SharePoint list data",10,10);
  const tableRows:any[]=[];
  filteredItems.forEach((item)=>{
    tableRows.push([item.Title,item.EmailAddress,item.Age,item.Admin,item.City]);
  });
  autoTable(doc,{
    head:[["Name","Email Address","Age","Admin","City"]],
    body:tableRows,
    startY:20
  });
  doc.save("splistdata.pdf");
}
  return(
    <>
    {/* search box */}
    <Input
    placeholder='search here...'
    style={{marginBottom:20,width:"300px"}}
    value={searchText}
    onChange={handleSearch}
    />
    {/* Page Size */}
    <Dropdown
    label='--select page--'
    options={pagintedItems}
    onChange={(_,options)=>{
      setPageSize(options?.key as any);
      SetPage(1)
    }}
     style={{marginBottom:20,width:"300px"}}
    />
    {/* Buttons */}
    <div style={{marginBottom:20,display:'flex',gap:'10px'}}>

<PrimaryButton

text="Export to Excel"
iconProps={{iconName:"FileExcel"}}
onClick={exportExcel}
styles={{root:{backgroundColor:"green",border:"green"}}}

/>
<PrimaryButton

text="Export to PDF"
iconProps={{iconName:"pdf"}}
onClick={exporttopdf}
styles={{root:{backgroundColor:"red",border:"red"}}}

/>

    </div>
    {/* Table */}
    <Table
    columns={columns}
    dataSource={paginatedItems}
    loading={loading}
    pagination={{
      current:page,
      pageSize:pageSize,
      total:filteredItems.length,
      onChange:(p)=>SetPage(p)
    }}
    rowKey="key"
    />
    </>
  )
}
export default  Paginatedtable;