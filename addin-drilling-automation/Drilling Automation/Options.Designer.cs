// -----------------------------------------------------------------------
// Copyright 2018 Autodesk, Inc. All rights reserved.
//
// Use of this software is subject to the terms of the Autodesk license
// agreement provided at the time of installation or download, or which
// otherwise accompanies this software in either electronic or hard copy form.
// -----------------------------------------------------------------------


namespace DrillingAutomation
{
    partial class Options
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Options));
            this.OK = new System.Windows.Forms.Button();
            this.TemplateLBL = new System.Windows.Forms.Label();
            this.Database = new System.Windows.Forms.TextBox();
            this.Browse_DB = new System.Windows.Forms.Button();
            this.BrowseTemplate = new System.Windows.Forms.OpenFileDialog();
            this.Cancel = new System.Windows.Forms.Button();
            this.Toolpath_Folder = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.Browse_TP = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.Tolerance = new System.Windows.Forms.TextBox();
            this.UseMethods = new System.Windows.Forms.RadioButton();
            this.UseToolpaths = new System.Windows.Forms.RadioButton();
            this.BrowseTemplateFolder = new System.Windows.Forms.FolderBrowserDialog();
            this.label3 = new System.Windows.Forms.Label();
            this.MachineCellCombo = new System.Windows.Forms.ComboBox();
            this.TopComponent = new System.Windows.Forms.RadioButton();
            this.TopHole = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.AllowTagDuplicate = new System.Windows.Forms.CheckBox();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // OK
            // 
            this.OK.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.OK.Location = new System.Drawing.Point(617, 255);
            this.OK.Name = "OK";
            this.OK.Size = new System.Drawing.Size(81, 26);
            this.OK.TabIndex = 120;
            this.OK.Text = "OK";
            this.OK.UseVisualStyleBackColor = true;
            this.OK.Click += new System.EventHandler(this.OK_Click);
            // 
            // TemplateLBL
            // 
            this.TemplateLBL.AutoSize = true;
            this.TemplateLBL.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TemplateLBL.Location = new System.Drawing.Point(12, 55);
            this.TemplateLBL.Name = "TemplateLBL";
            this.TemplateLBL.Size = new System.Drawing.Size(167, 20);
            this.TemplateLBL.TabIndex = 10;
            this.TemplateLBL.Text = "Hole Database (CSV):";
            // 
            // Database
            // 
            this.Database.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Database.Location = new System.Drawing.Point(185, 52);
            this.Database.Name = "Database";
            this.Database.ReadOnly = true;
            this.Database.Size = new System.Drawing.Size(473, 26);
            this.Database.TabIndex = 20;
            // 
            // Browse_DB
            // 
            this.Browse_DB.BackColor = System.Drawing.SystemColors.Control;
            this.Browse_DB.BackgroundImage = global::DrillingAutomation.Properties.Resources.Open___File_Browser_32x32;
            this.Browse_DB.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Browse_DB.Location = new System.Drawing.Point(664, 49);
            this.Browse_DB.Name = "Browse_DB";
            this.Browse_DB.Size = new System.Drawing.Size(34, 34);
            this.Browse_DB.TabIndex = 30;
            this.Browse_DB.UseVisualStyleBackColor = false;
            this.Browse_DB.Click += new System.EventHandler(this.BrowseDB_Click);
            // 
            // BrowseTemplate
            // 
            this.BrowseTemplate.FileName = "BrowseDatabase";
            // 
            // Cancel
            // 
            this.Cancel.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.Cancel.Location = new System.Drawing.Point(515, 255);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(81, 26);
            this.Cancel.TabIndex = 110;
            this.Cancel.Text = "Cancel";
            this.Cancel.UseVisualStyleBackColor = true;
            this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // Toolpath_Folder
            // 
            this.Toolpath_Folder.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Toolpath_Folder.Location = new System.Drawing.Point(185, 97);
            this.Toolpath_Folder.Name = "Toolpath_Folder";
            this.Toolpath_Folder.ReadOnly = true;
            this.Toolpath_Folder.Size = new System.Drawing.Size(473, 26);
            this.Toolpath_Folder.TabIndex = 50;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(12, 100);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(124, 20);
            this.label1.TabIndex = 40;
            this.label1.Text = "Toolpath Folder:";
            // 
            // Browse_TP
            // 
            this.Browse_TP.BackColor = System.Drawing.SystemColors.Control;
            this.Browse_TP.BackgroundImage = global::DrillingAutomation.Properties.Resources.Open___File_Browser_32x32;
            this.Browse_TP.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Browse_TP.Location = new System.Drawing.Point(664, 94);
            this.Browse_TP.Name = "Browse_TP";
            this.Browse_TP.Size = new System.Drawing.Size(34, 34);
            this.Browse_TP.TabIndex = 60;
            this.Browse_TP.UseVisualStyleBackColor = false;
            this.Browse_TP.Click += new System.EventHandler(this.Browse_TP_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(12, 140);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(158, 20);
            this.label2.TabIndex = 70;
            this.label2.Text = "Extraction Tolerance:";
            // 
            // Tolerance
            // 
            this.Tolerance.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.Tolerance.Location = new System.Drawing.Point(185, 137);
            this.Tolerance.Name = "Tolerance";
            this.Tolerance.Size = new System.Drawing.Size(57, 26);
            this.Tolerance.TabIndex = 80;
            this.Tolerance.Text = "0.002";
            // 
            // UseMethods
            // 
            this.UseMethods.AutoSize = true;
            this.UseMethods.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.UseMethods.Location = new System.Drawing.Point(6, 16);
            this.UseMethods.Name = "UseMethods";
            this.UseMethods.Size = new System.Drawing.Size(132, 24);
            this.UseMethods.TabIndex = 90;
            this.UseMethods.Text = "Drilling Method";
            this.UseMethods.UseVisualStyleBackColor = true;
            // 
            // UseToolpaths
            // 
            this.UseToolpaths.AutoSize = true;
            this.UseToolpaths.Checked = true;
            this.UseToolpaths.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.UseToolpaths.Location = new System.Drawing.Point(175, 16);
            this.UseToolpaths.Name = "UseToolpaths";
            this.UseToolpaths.Size = new System.Drawing.Size(97, 24);
            this.UseToolpaths.TabIndex = 100;
            this.UseToolpaths.TabStop = true;
            this.UseToolpaths.Text = "Toolpaths";
            this.UseToolpaths.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(14, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(103, 20);
            this.label3.TabIndex = 121;
            this.label3.Text = "Machine Cell:";
            // 
            // MachineCellCombo
            // 
            this.MachineCellCombo.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.MachineCellCombo.FormattingEnabled = true;
            this.MachineCellCombo.Location = new System.Drawing.Point(185, 9);
            this.MachineCellCombo.Name = "MachineCellCombo";
            this.MachineCellCombo.Size = new System.Drawing.Size(473, 28);
            this.MachineCellCombo.TabIndex = 122;
            this.MachineCellCombo.DropDownClosed += new System.EventHandler(this.MachineCellCombo_DropDownClosed);
            // 
            // TopComponent
            // 
            this.TopComponent.AutoSize = true;
            this.TopComponent.Checked = true;
            this.TopComponent.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.TopComponent.Location = new System.Drawing.Point(6, 19);
            this.TopComponent.Name = "TopComponent";
            this.TopComponent.Size = new System.Drawing.Size(159, 24);
            this.TopComponent.TabIndex = 127;
            this.TopComponent.TabStop = true;
            this.TopComponent.Text = "Top of Component";
            this.TopComponent.UseVisualStyleBackColor = true;
            // 
            // TopHole
            // 
            this.TopHole.AutoSize = true;
            this.TopHole.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.TopHole.Location = new System.Drawing.Point(177, 19);
            this.TopHole.Name = "TopHole";
            this.TopHole.Size = new System.Drawing.Size(109, 24);
            this.TopHole.TabIndex = 128;
            this.TopHole.Text = "Top of Hole";
            this.TopHole.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.UseToolpaths);
            this.groupBox1.Controls.Add(this.UseMethods);
            this.groupBox1.Location = new System.Drawing.Point(18, 178);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(353, 50);
            this.groupBox1.TabIndex = 129;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Toolpath Type";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.TopComponent);
            this.groupBox2.Controls.Add(this.TopHole);
            this.groupBox2.Location = new System.Drawing.Point(16, 230);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(355, 50);
            this.groupBox2.TabIndex = 130;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Depth Reference";
            // 
            // AllowTagDuplicate
            // 
            this.AllowTagDuplicate.AutoSize = true;
            this.AllowTagDuplicate.Checked = true;
            this.AllowTagDuplicate.CheckState = System.Windows.Forms.CheckState.Checked;
            this.AllowTagDuplicate.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F);
            this.AllowTagDuplicate.Location = new System.Drawing.Point(322, 139);
            this.AllowTagDuplicate.Name = "AllowTagDuplicate";
            this.AllowTagDuplicate.Size = new System.Drawing.Size(336, 24);
            this.AllowTagDuplicate.TabIndex = 131;
            this.AllowTagDuplicate.Text = "Allow same tag on multiple hole in database";
            this.AllowTagDuplicate.UseVisualStyleBackColor = true;
            // 
            // Options
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(710, 293);
            this.Controls.Add(this.AllowTagDuplicate);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.MachineCellCombo);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Tolerance);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.Browse_TP);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Toolpath_Folder);
            this.Controls.Add(this.Cancel);
            this.Controls.Add(this.Browse_DB);
            this.Controls.Add(this.Database);
            this.Controls.Add(this.TemplateLBL);
            this.Controls.Add(this.OK);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Options";
            this.Text = "Options";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button OK;
        private System.Windows.Forms.Label TemplateLBL;
        private System.Windows.Forms.TextBox Database;
        private System.Windows.Forms.Button Browse_DB;
        private System.Windows.Forms.OpenFileDialog BrowseTemplate;
        private System.Windows.Forms.Button Cancel;
        private System.Windows.Forms.TextBox Toolpath_Folder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button Browse_TP;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox Tolerance;
        private System.Windows.Forms.RadioButton UseMethods;
        private System.Windows.Forms.RadioButton UseToolpaths;
        private System.Windows.Forms.FolderBrowserDialog BrowseTemplateFolder;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox MachineCellCombo;
        private System.Windows.Forms.RadioButton TopComponent;
        private System.Windows.Forms.RadioButton TopHole;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.CheckBox AllowTagDuplicate;
    }
}