using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Xrm;
using System.Data;


namespace CRM.Models
{
    class Leads : SortableBindingList<Lead>
    {
        public Leads() : base() { }

        public Leads(string FilePath, string RangeName)
            : base()
        {
            Lead objLead;
            ExcelInfo objExcelInfo;


            objExcelInfo = new ExcelInfo(FilePath);
            objExcelInfo.LoadDataFromExcel(RangeName);

            var objFilteredList = objExcelInfo.ExcelDataSet.Tables[0].AsEnumerable();

            foreach (var objItem in objFilteredList)
            {
                objLead = new Lead()
                {
                    Subject = objItem.Field<string>("Topic"), //this is the Topic field
                    FirstName = objItem.Field<string>("First Name"),
                    LastName = objItem.Field<string>("Last Name"),
                    CompanyName = objItem.Field<string>("Company"),
                    Telephone1 = objItem.Field<string>("Phone Number"), //this is the business phone number
                    Description = objItem.Field<string>("Description"),
                    LeadSourceCode = 100000002
                };
                this.Add(objLead);
            }
        }
    }
}
