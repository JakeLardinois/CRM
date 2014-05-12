using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Xrm;
using System.Data;


namespace CRM.Models
{
    class Leads : SortableBindingList<MyLead>
    {
        
        public Leads() : base() { }

        public Leads(string FilePath, string RangeName)
            : base()
        {
            MyLead objLead;
            ExcelInfo objExcelInfo;
            int intCounter;
            StringBuilder objStrBldr;
            List<string> Questions;
            

            objExcelInfo = new ExcelInfo(FilePath);
            objExcelInfo.LoadDataFromExcel(RangeName);

            var objFilteredList = objExcelInfo.ExcelDataSet.Tables[0].AsEnumerable();

            intCounter = 0;
            objStrBldr = new StringBuilder();
            Questions = new List<string>();

            foreach (var objItem in objFilteredList)
            {
                //The first row contains the questions that I need to get... So I put those questions into a collection here...
                if (intCounter == 0) 
                {
                    intCounter++;

                    Questions.Add(objItem.Field<string>("Q4"));
                    Questions.Add(objItem.Field<string>("Q5"));
                    Questions.Add(objItem.Field<string>("Q6"));
                    Questions.Add(objItem.Field<string>("Q7"));
                    Questions.Add(objItem.Field<string>("Q8"));
                    Questions.Add(objItem.Field<string>("Q9"));
                    Questions.Add(objItem.Field<string>("Q12"));
                    Questions.Add(objItem.Field<string>("Q13"));
                    Questions.Add(objItem.Field<string>("Q14"));
                    Questions.Add(objItem.Field<string>("Q15"));
                    Questions.Add(objItem.Field<string>("Q16"));
                    Questions.Add(objItem.Field<string>("Q17_1"));
                    Questions.Add(objItem.Field<string>("Q18_1"));
                    Questions.Add(objItem.Field<string>("Q18_2"));
                    Questions.Add(objItem.Field<string>("Q19_1"));
                    Questions.Add(objItem.Field<string>("Q19_2"));
                    Questions.Add(objItem.Field<string>("Q20_1"));
                    Questions.Add(objItem.Field<string>("Q21_1"));

                    continue; //don't create a lead object for this iteration
                }

                //builds the questions and answers for the Description text...
                objStrBldr.Append(Questions[0] + Environment.NewLine + "\t" + objItem.Field<string>("Q4") + Environment.NewLine);
                objStrBldr.Append(Questions[1] + Environment.NewLine + "\t" + objItem.Field<string>("Q5") + Environment.NewLine);
                objStrBldr.Append(Questions[2] + Environment.NewLine + "\t" + objItem.Field<string>("Q6") + Environment.NewLine);
                objStrBldr.Append(Questions[3] + Environment.NewLine + "\t" + objItem.Field<string>("Q7") + Environment.NewLine);
                objStrBldr.Append(Questions[4] + Environment.NewLine + "\t" + objItem.Field<string>("Q8") + Environment.NewLine);
                objStrBldr.Append(Questions[5] + Environment.NewLine + "\t" + objItem.Field<string>("Q9") + Environment.NewLine);
                objStrBldr.Append(Questions[6] + Environment.NewLine + "\t" + objItem.Field<string>("Q12") + Environment.NewLine);
                objStrBldr.Append(Questions[7] + Environment.NewLine + "\t" + objItem.Field<string>("Q13") + Environment.NewLine);
                objStrBldr.Append(Questions[8] + Environment.NewLine + "\t" + objItem.Field<string>("Q14") + Environment.NewLine);
                objStrBldr.Append(Questions[9] + Environment.NewLine + "\t" + objItem.Field<string>("Q15") + Environment.NewLine);
                objStrBldr.Append(Questions[10] + Environment.NewLine + "\t" + objItem.Field<string>("Q16") + Environment.NewLine);
                objStrBldr.Append(Questions[11] + Environment.NewLine + "\t" + objItem.Field<string>("Q17_1") + Environment.NewLine);
                objStrBldr.Append(Questions[12] + Environment.NewLine + "\t" + objItem.Field<string>("Q18_1") + Environment.NewLine);
                objStrBldr.Append(Questions[13] + Environment.NewLine + "\t" + objItem.Field<string>("Q19_1") + Environment.NewLine);
                objStrBldr.Append(Questions[14] + Environment.NewLine + "\t" + objItem.Field<string>("Q19_2") + Environment.NewLine);
                objStrBldr.Append(Questions[15] + Environment.NewLine + "\t" + objItem.Field<string>("Q20_1") + Environment.NewLine);
                objStrBldr.Append(Questions[16] + Environment.NewLine + "\t" + objItem.Field<string>("Q21_1") + Environment.NewLine);

                //Builds my annotation (notes) object and assigns the above created next to it.
                Annotation objAnnotation = new Annotation();
                objAnnotation.Subject = "Survey Results";
                objAnnotation.NoteText = objStrBldr.ToString();

                objLead = new MyLead()
                {
                    Subject = objItem.Field<string>("Q1") + " " + objItem.Field<string>("Q2") + " Qualtrics Lead " + intCounter, //this is the Topic field
                    FirstName = objItem.Field<string>("Q1"),
                    LastName = objItem.Field<string>("Q2"),
                    EMailAddress1 = objItem.Field<string>("Q3"),
                    //Description = "A description Here",
                    Note = objAnnotation,
                    LeadSourceCode = 100000003
                };

                this.Add(objLead);
                intCounter++;
            }
        }
    }
}
