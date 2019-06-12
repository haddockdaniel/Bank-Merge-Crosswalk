﻿namespace JurisUtilityBase
{
    partial class UtilityBaseMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(UtilityBaseMain));
            this.JurisLogoImageBox = new System.Windows.Forms.PictureBox();
            this.LexisNexisLogoPictureBox = new System.Windows.Forms.PictureBox();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.listBoxCompanies = new System.Windows.Forms.ListBox();
            this.labelDescription = new System.Windows.Forms.Label();
            this.statusGroupBox = new System.Windows.Forms.GroupBox();
            this.labelCurrentStatus = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.labelPercentComplete = new System.Windows.Forms.Label();
            this.OpenFileDialogOpen = new System.Windows.Forms.OpenFileDialog();
            this.button1 = new System.Windows.Forms.Button();
            this.buttonReport = new System.Windows.Forms.Button();
            this.cbBank = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.cbNew = new System.Windows.Forms.ComboBox();
            this.lblBank = new System.Windows.Forms.Label();
            this.txtBnkCode = new System.Windows.Forms.TextBox();
            this.txtBankDesc = new System.Windows.Forms.TextBox();
            ((System.ComponentModel.ISupportInitialize)(this.JurisLogoImageBox)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LexisNexisLogoPictureBox)).BeginInit();
            this.statusStrip.SuspendLayout();
            this.statusGroupBox.SuspendLayout();
            this.SuspendLayout();
            // 
            // JurisLogoImageBox
            // 
            this.JurisLogoImageBox.Image = ((System.Drawing.Image)(resources.GetObject("JurisLogoImageBox.Image")));
            this.JurisLogoImageBox.InitialImage = ((System.Drawing.Image)(resources.GetObject("JurisLogoImageBox.InitialImage")));
            this.JurisLogoImageBox.Location = new System.Drawing.Point(0, 1);
            this.JurisLogoImageBox.Name = "JurisLogoImageBox";
            this.JurisLogoImageBox.Size = new System.Drawing.Size(104, 336);
            this.JurisLogoImageBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.JurisLogoImageBox.TabIndex = 0;
            this.JurisLogoImageBox.TabStop = false;
            // 
            // LexisNexisLogoPictureBox
            // 
            this.LexisNexisLogoPictureBox.Image = ((System.Drawing.Image)(resources.GetObject("LexisNexisLogoPictureBox.Image")));
            this.LexisNexisLogoPictureBox.Location = new System.Drawing.Point(8, 343);
            this.LexisNexisLogoPictureBox.Name = "LexisNexisLogoPictureBox";
            this.LexisNexisLogoPictureBox.Size = new System.Drawing.Size(96, 28);
            this.LexisNexisLogoPictureBox.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.LexisNexisLogoPictureBox.TabIndex = 1;
            this.LexisNexisLogoPictureBox.TabStop = false;
            // 
            // statusStrip
            // 
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel});
            this.statusStrip.Location = new System.Drawing.Point(0, 394);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(662, 22);
            this.statusStrip.TabIndex = 2;
            // 
            // toolStripStatusLabel
            // 
            this.toolStripStatusLabel.BackColor = System.Drawing.SystemColors.Control;
            this.toolStripStatusLabel.ForeColor = System.Drawing.SystemColors.ControlText;
            this.toolStripStatusLabel.Name = "toolStripStatusLabel";
            this.toolStripStatusLabel.Size = new System.Drawing.Size(134, 17);
            this.toolStripStatusLabel.Text = "Status: Ready to Execute";
            // 
            // listBoxCompanies
            // 
            this.listBoxCompanies.FormattingEnabled = true;
            this.listBoxCompanies.Location = new System.Drawing.Point(111, 1);
            this.listBoxCompanies.Name = "listBoxCompanies";
            this.listBoxCompanies.Size = new System.Drawing.Size(262, 69);
            this.listBoxCompanies.TabIndex = 0;
            this.listBoxCompanies.SelectedIndexChanged += new System.EventHandler(this.listBoxCompanies_SelectedIndexChanged);
            // 
            // labelDescription
            // 
            this.labelDescription.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            this.labelDescription.Location = new System.Drawing.Point(380, 1);
            this.labelDescription.Name = "labelDescription";
            this.labelDescription.Size = new System.Drawing.Size(268, 69);
            this.labelDescription.TabIndex = 1;
            this.labelDescription.Text = "Merge Banks";
            // 
            // statusGroupBox
            // 
            this.statusGroupBox.Controls.Add(this.labelCurrentStatus);
            this.statusGroupBox.Controls.Add(this.progressBar);
            this.statusGroupBox.Controls.Add(this.labelPercentComplete);
            this.statusGroupBox.Location = new System.Drawing.Point(111, 77);
            this.statusGroupBox.Name = "statusGroupBox";
            this.statusGroupBox.Size = new System.Drawing.Size(537, 73);
            this.statusGroupBox.TabIndex = 5;
            this.statusGroupBox.TabStop = false;
            this.statusGroupBox.Text = "Utility Status:";
            // 
            // labelCurrentStatus
            // 
            this.labelCurrentStatus.AutoSize = true;
            this.labelCurrentStatus.Location = new System.Drawing.Point(7, 50);
            this.labelCurrentStatus.Name = "labelCurrentStatus";
            this.labelCurrentStatus.Size = new System.Drawing.Size(77, 13);
            this.labelCurrentStatus.TabIndex = 2;
            this.labelCurrentStatus.Text = "Current Status:";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(7, 27);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(519, 20);
            this.progressBar.TabIndex = 0;
            // 
            // labelPercentComplete
            // 
            this.labelPercentComplete.AutoSize = true;
            this.labelPercentComplete.Location = new System.Drawing.Point(413, 11);
            this.labelPercentComplete.Name = "labelPercentComplete";
            this.labelPercentComplete.Size = new System.Drawing.Size(62, 13);
            this.labelPercentComplete.TabIndex = 0;
            this.labelPercentComplete.Text = "% Complete";
            // 
            // OpenFileDialogOpen
            // 
            this.OpenFileDialogOpen.FileName = "openFileDialog1";
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(142, 298);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(425, 62);
            this.button1.TabIndex = 6;
            this.button1.Text = "Execute custom fix";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // buttonReport
            // 
            this.buttonReport.BackColor = System.Drawing.Color.LightGray;
            this.buttonReport.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonReport.ForeColor = System.Drawing.Color.MidnightBlue;
            this.buttonReport.Location = new System.Drawing.Point(573, 310);
            this.buttonReport.Name = "buttonReport";
            this.buttonReport.Size = new System.Drawing.Size(50, 38);
            this.buttonReport.TabIndex = 16;
            this.buttonReport.Text = "Pre-Report";
            this.buttonReport.UseVisualStyleBackColor = false;
            this.buttonReport.Visible = false;
            this.buttonReport.Click += new System.EventHandler(this.buttonReport_Click);
            // 
            // cbBank
            // 
            this.cbBank.FormattingEnabled = true;
            this.cbBank.Location = new System.Drawing.Point(311, 167);
            this.cbBank.Name = "cbBank";
            this.cbBank.Size = new System.Drawing.Size(312, 21);
            this.cbBank.TabIndex = 17;
            this.cbBank.SelectedIndexChanged += new System.EventHandler(this.cbBank_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(132, 170);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(173, 13);
            this.label1.TabIndex = 18;
            this.label1.Text = "Select Bank Account to Move from";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(132, 207);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(165, 13);
            this.label2.TabIndex = 19;
            this.label2.Text = "Select Bank Account to Merge to";
            // 
            // cbNew
            // 
            this.cbNew.FormattingEnabled = true;
            this.cbNew.Location = new System.Drawing.Point(311, 199);
            this.cbNew.Name = "cbNew";
            this.cbNew.Size = new System.Drawing.Size(312, 21);
            this.cbNew.TabIndex = 20;
            this.cbNew.SelectedIndexChanged += new System.EventHandler(this.cbNew_SelectedIndexChanged);
            // 
            // lblBank
            // 
            this.lblBank.AutoSize = true;
            this.lblBank.Location = new System.Drawing.Point(132, 238);
            this.lblBank.Name = "lblBank";
            this.lblBank.Size = new System.Drawing.Size(190, 13);
            this.lblBank.TabIndex = 21;
            this.lblBank.Text = "Enter New Bank Code and Description";
            this.lblBank.Visible = false;
            // 
            // txtBnkCode
            // 
            this.txtBnkCode.Location = new System.Drawing.Point(329, 238);
            this.txtBnkCode.MaxLength = 4;
            this.txtBnkCode.Name = "txtBnkCode";
            this.txtBnkCode.Size = new System.Drawing.Size(100, 20);
            this.txtBnkCode.TabIndex = 22;
            this.txtBnkCode.Visible = false;
            // 
            // txtBankDesc
            // 
            this.txtBankDesc.Location = new System.Drawing.Point(329, 264);
            this.txtBankDesc.MaxLength = 30;
            this.txtBankDesc.Name = "txtBankDesc";
            this.txtBankDesc.Size = new System.Drawing.Size(308, 20);
            this.txtBankDesc.TabIndex = 23;
            this.txtBankDesc.Visible = false;
            // 
            // UtilityBaseMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Window;
            this.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.ClientSize = new System.Drawing.Size(662, 416);
            this.Controls.Add(this.txtBankDesc);
            this.Controls.Add(this.txtBnkCode);
            this.Controls.Add(this.lblBank);
            this.Controls.Add(this.cbNew);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cbBank);
            this.Controls.Add(this.buttonReport);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.statusGroupBox);
            this.Controls.Add(this.labelDescription);
            this.Controls.Add(this.listBoxCompanies);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.LexisNexisLogoPictureBox);
            this.Controls.Add(this.JurisLogoImageBox);
            this.ForeColor = System.Drawing.SystemColors.WindowText;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "UtilityBaseMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Juris Utility - Bank Merge";
            this.Load += new System.EventHandler(this.Form1_Load);
            ((System.ComponentModel.ISupportInitialize)(this.JurisLogoImageBox)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LexisNexisLogoPictureBox)).EndInit();
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.statusGroupBox.ResumeLayout(false);
            this.statusGroupBox.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.PictureBox JurisLogoImageBox;
        private System.Windows.Forms.PictureBox LexisNexisLogoPictureBox;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
        private System.Windows.Forms.ListBox listBoxCompanies;
        private System.Windows.Forms.Label labelDescription;
        private System.Windows.Forms.GroupBox statusGroupBox;
        private System.Windows.Forms.Label labelCurrentStatus;
        private System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Label labelPercentComplete;
        private System.Windows.Forms.OpenFileDialog OpenFileDialogOpen;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button buttonReport;
        private System.Windows.Forms.ComboBox cbBank;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox cbNew;
        private System.Windows.Forms.Label lblBank;
        private System.Windows.Forms.TextBox txtBnkCode;
        private System.Windows.Forms.TextBox txtBankDesc;
    }
}

