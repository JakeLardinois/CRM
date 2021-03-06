﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Xrm;
using CRM.Models;
using System.IO;


namespace CRM
{
    public partial class frmAddLeads : Form
    {
        private static BackgroundWorker bw;

        private Leads Leads { get; set; }


        public frmAddLeads()
        {
            InitializeComponent();

            bw = new BackgroundWorker();
            bw.DoWork += new DoWorkEventHandler(bw_DoWork);
            bw.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bw_RunWorkerCompleted);

            Leads = new Leads();
        }

        private void bw_DoWork(object sender, DoWorkEventArgs e)
        {
            XrmServiceContext xrm;


            xrm = new XrmServiceContext("Xrm");

            foreach (MyLead objLead in Leads)
            {
                //xrm.AddObject(objLead);
                Guid objLeadGuid = xrm.Create(objLead);//Using this method as opposed to xrm.AddObject results in getting the GUID required to add a note to the lead.
                                                        //This method also adds the Lead Object to the Db.
                objLead.Note.ObjectId = new Microsoft.Xrm.Client.CrmEntityReference("lead", objLeadGuid); //Creates the xref between the Annotation object (the note) and the lead
                xrm.AddObject(objLead.Note); //adds the note to the CRM Database..
            }
                
            xrm.SaveChanges();//this saves all the notes to the database.  The leads are automatically saved by calling the Create() method...
        }

        private void bw_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!(e.Error == null))
            {
                MessageBox.Show(e.Error.Message, "Upload Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {
                MessageBox.Show("The Leads Were Successfully Uploaded!", "Upload Succeeded", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void frmAddLeads_Load(object sender, EventArgs e)
        {
            dgvLeads.AutoGenerateColumns = false;
            dgvLeads.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dgvLeads.DataSource = Leads;
        }

        private void getRangeList()
        {
            ExcelInfo objExcelInfo;


            objExcelInfo = new ExcelInfo(txtFilePath.Text);

            if (objExcelInfo.GetInformation())
                cmbRangeList.DataSource = objExcelInfo.NameRanges;
            else
                MessageBox.Show((String.Format("Failed to get back information{0}{1}", Environment.NewLine, objExcelInfo.LastException.Message)));
        }

        private void cmdGetFilePath_Click(object sender, EventArgs e)
        {
            if (ofdExcel.ShowDialog() == DialogResult.OK)
            {
                txtFilePath.Text = ofdExcel.FileName;
                getRangeList();
            }
        }

        private void btnViewData_Click(object sender, EventArgs e)
        {
            if (File.Exists(txtFilePath.Text) && !string.IsNullOrEmpty(txtFilePath.Text))
            {
                Leads = new Leads(txtFilePath.Text, cmbRangeList.Text);
                dgvLeads.DataSource = Leads;
            }
            else
                MessageBox.Show("The file does not exist...", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {
            if (!bw.IsBusy)
                bw.RunWorkerAsync();
            else
                showWaitMessage();
        }

        private void showWaitMessage()
        {
            MessageBox.Show("Wait for background process to finish.  Try again later.", "Wait...", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
