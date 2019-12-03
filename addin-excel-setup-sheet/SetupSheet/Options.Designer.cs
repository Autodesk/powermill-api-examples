// -----------------------------------------------------------------------
// Copyright 2019 Autodesk, Inc. All rights reserved.
//
// Use of this software is subject to the terms of the Autodesk license
// agreement provided at the time of installation or download, or which
// otherwise accompanies this software in either electronic or hard copy form.
// -----------------------------------------------------------------------
namespace SetupSheet
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
            this.Template = new System.Windows.Forms.TextBox();
            this.Browse = new System.Windows.Forms.Button();
            this.BrowseTemplate = new System.Windows.Forms.OpenFileDialog();
            this.Cancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // OK
            // 
            this.OK.Location = new System.Drawing.Point(525, 54);
            this.OK.Name = "OK";
            this.OK.Size = new System.Drawing.Size(62, 25);
            this.OK.TabIndex = 0;
            this.OK.Text = "OK";
            this.OK.UseVisualStyleBackColor = true;
            this.OK.Click += new System.EventHandler(this.OK_Click);
            // 
            // TemplateLBL
            // 
            this.TemplateLBL.AutoSize = true;
            this.TemplateLBL.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TemplateLBL.Location = new System.Drawing.Point(12, 13);
            this.TemplateLBL.Name = "TemplateLBL";
            this.TemplateLBL.Size = new System.Drawing.Size(79, 20);
            this.TemplateLBL.TabIndex = 1;
            this.TemplateLBL.Text = "Template:";
            // 
            // Template
            // 
            this.Template.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.Template.Location = new System.Drawing.Point(94, 10);
            this.Template.Name = "Template";
            this.Template.ReadOnly = true;
            this.Template.Size = new System.Drawing.Size(451, 26);
            this.Template.TabIndex = 2;
            // 
            // Browse
            // 
            this.Browse.BackColor = System.Drawing.SystemColors.Control;
            this.Browse.BackgroundImage = global::SetupSheet.Properties.Resources.Open___File_Browser_32x32;
            this.Browse.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch;
            this.Browse.Location = new System.Drawing.Point(551, 6);
            this.Browse.Name = "Browse";
            this.Browse.Size = new System.Drawing.Size(34, 34);
            this.Browse.TabIndex = 3;
            this.Browse.UseVisualStyleBackColor = false;
            this.Browse.Click += new System.EventHandler(this.Browse_Click);
            // 
            // BrowseTemplate
            // 
            this.BrowseTemplate.FileName = "BrowseTemplate";
            // 
            // Cancel
            // 
            this.Cancel.Location = new System.Drawing.Point(454, 55);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(56, 23);
            this.Cancel.TabIndex = 4;
            this.Cancel.Text = "Cancel";
            this.Cancel.UseVisualStyleBackColor = true;
            this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // Options
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(599, 89);
            this.Controls.Add(this.Cancel);
            this.Controls.Add(this.Browse);
            this.Controls.Add(this.Template);
            this.Controls.Add(this.TemplateLBL);
            this.Controls.Add(this.OK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Options";
            this.Text = "Options";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button OK;
        private System.Windows.Forms.Label TemplateLBL;
        private System.Windows.Forms.TextBox Template;
        private System.Windows.Forms.Button Browse;
        private System.Windows.Forms.OpenFileDialog BrowseTemplate;
        private System.Windows.Forms.Button Cancel;
    }
}