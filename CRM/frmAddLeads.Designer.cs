namespace CRM
{
    partial class frmAddLeads
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmAddLeads));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.dgvLeads = new System.Windows.Forms.DataGridView();
            this.txtTopic = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtFirstName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtLastName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtCompany = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtPhoneNo = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.txtDescription = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnUpload = new System.Windows.Forms.Button();
            this.cmbRangeList = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.cmdGetFilePath = new System.Windows.Forms.Button();
            this.label6 = new System.Windows.Forms.Label();
            this.txtFilePath = new System.Windows.Forms.TextBox();
            this.btnViewData = new System.Windows.Forms.Button();
            this.ofdExcel = new System.Windows.Forms.OpenFileDialog();
            this.groupBox1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvLeads)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.groupBox1.Controls.Add(this.dgvLeads);
            this.groupBox1.Location = new System.Drawing.Point(12, 70);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(969, 274);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Leads";
            // 
            // dgvLeads
            // 
            this.dgvLeads.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvLeads.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvLeads.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.txtTopic,
            this.txtFirstName,
            this.txtLastName,
            this.txtCompany,
            this.txtPhoneNo,
            this.txtDescription});
            this.dgvLeads.Location = new System.Drawing.Point(6, 19);
            this.dgvLeads.Name = "dgvLeads";
            this.dgvLeads.Size = new System.Drawing.Size(957, 249);
            this.dgvLeads.TabIndex = 0;
            // 
            // txtTopic
            // 
            this.txtTopic.DataPropertyName = "Subject";
            this.txtTopic.HeaderText = "Topic";
            this.txtTopic.Name = "txtTopic";
            // 
            // txtFirstName
            // 
            this.txtFirstName.DataPropertyName = "FirstName";
            this.txtFirstName.FillWeight = 30F;
            this.txtFirstName.HeaderText = "First Name";
            this.txtFirstName.Name = "txtFirstName";
            // 
            // txtLastName
            // 
            this.txtLastName.DataPropertyName = "LastName";
            this.txtLastName.FillWeight = 40F;
            this.txtLastName.HeaderText = "Last Name";
            this.txtLastName.Name = "txtLastName";
            // 
            // txtCompany
            // 
            this.txtCompany.DataPropertyName = "CompanyName";
            this.txtCompany.FillWeight = 40F;
            this.txtCompany.HeaderText = "Company";
            this.txtCompany.Name = "txtCompany";
            // 
            // txtPhoneNo
            // 
            this.txtPhoneNo.DataPropertyName = "Telephone1";
            this.txtPhoneNo.FillWeight = 30F;
            this.txtPhoneNo.HeaderText = "Phone Number";
            this.txtPhoneNo.Name = "txtPhoneNo";
            // 
            // txtDescription
            // 
            this.txtDescription.DataPropertyName = "Description";
            this.txtDescription.HeaderText = "Description";
            this.txtDescription.Name = "txtDescription";
            // 
            // btnUpload
            // 
            this.btnUpload.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnUpload.Location = new System.Drawing.Point(860, 41);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(115, 23);
            this.btnUpload.TabIndex = 92;
            this.btnUpload.Text = "Upload To CRM";
            this.btnUpload.UseVisualStyleBackColor = true;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // cmbRangeList
            // 
            this.cmbRangeList.FormattingEnabled = true;
            this.cmbRangeList.Location = new System.Drawing.Point(93, 38);
            this.cmbRangeList.Name = "cmbRangeList";
            this.cmbRangeList.Size = new System.Drawing.Size(206, 21);
            this.cmbRangeList.TabIndex = 97;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(16, 41);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(76, 13);
            this.label1.TabIndex = 96;
            this.label1.Text = "Range Name :";
            // 
            // cmdGetFilePath
            // 
            this.cmdGetFilePath.Location = new System.Drawing.Point(608, 12);
            this.cmdGetFilePath.Name = "cmdGetFilePath";
            this.cmdGetFilePath.Size = new System.Drawing.Size(35, 20);
            this.cmdGetFilePath.TabIndex = 95;
            this.cmdGetFilePath.Text = "...";
            this.cmdGetFilePath.UseVisualStyleBackColor = true;
            this.cmdGetFilePath.Click += new System.EventHandler(this.cmdGetFilePath_Click);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(16, 15);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(58, 13);
            this.label6.TabIndex = 94;
            this.label6.Text = "Excel File :";
            // 
            // txtFilePath
            // 
            this.txtFilePath.Location = new System.Drawing.Point(93, 12);
            this.txtFilePath.Name = "txtFilePath";
            this.txtFilePath.Size = new System.Drawing.Size(518, 20);
            this.txtFilePath.TabIndex = 93;
            // 
            // btnViewData
            // 
            this.btnViewData.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnViewData.Location = new System.Drawing.Point(900, 9);
            this.btnViewData.Name = "btnViewData";
            this.btnViewData.Size = new System.Drawing.Size(75, 23);
            this.btnViewData.TabIndex = 98;
            this.btnViewData.Text = "View";
            this.btnViewData.UseVisualStyleBackColor = true;
            this.btnViewData.Click += new System.EventHandler(this.btnViewData_Click);
            // 
            // ofdExcel
            // 
            this.ofdExcel.FileName = "excel.xlsx";
            this.ofdExcel.Filter = "Excel Files|*.xls;*.xlsx|All Files|*.*";
            // 
            // frmAddLeads
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(993, 356);
            this.Controls.Add(this.btnViewData);
            this.Controls.Add(this.cmbRangeList);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cmdGetFilePath);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.txtFilePath);
            this.Controls.Add(this.btnUpload);
            this.Controls.Add(this.groupBox1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmAddLeads";
            this.Text = "Add Leads to Microsoft CRM";
            this.Load += new System.EventHandler(this.frmAddLeads_Load);
            this.groupBox1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvLeads)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DataGridView dgvLeads;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.ComboBox cmbRangeList;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button cmdGetFilePath;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.TextBox txtFilePath;
        private System.Windows.Forms.Button btnViewData;
        private System.Windows.Forms.OpenFileDialog ofdExcel;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtTopic;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtFirstName;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtLastName;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtCompany;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtPhoneNo;
        private System.Windows.Forms.DataGridViewTextBoxColumn txtDescription;
    }
}