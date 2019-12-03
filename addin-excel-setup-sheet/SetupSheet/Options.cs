// -----------------------------------------------------------------------
// Copyright 2019 Autodesk, Inc. All rights reserved.
//
// Use of this software is subject to the terms of the Autodesk license
// agreement provided at the time of installation or download, or which
// otherwise accompanies this software in either electronic or hard copy form.
// -----------------------------------------------------------------------

using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SetupSheet
{
    public partial class Options : Form
    {
        public string FolderPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\ExcellSetupSheet";
        public Options()
        {
            InitializeComponent();
            string Path = FolderPath + "\\Template.ini";

            if (File.Exists(Path))
            {
                const Int32 BufferSize = 128;
                using (var fileStream = File.OpenRead(Path))
                using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
                {
                    String line;
                    while ((line = streamReader.ReadLine()) != null)
                    {
                        Template.Text = line;
                    }
                }
            }
            else
            {
                Path = System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Substring(8, System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Length-8);
                string directory = System.IO.Path.GetDirectoryName(Path);
                Path = directory + "\\Tool_List.xlsx";
                Template.Text = Path;
            }
        }

        private void OK_Click(object sender, EventArgs e)
        {
            string Path = FolderPath + "\\Template.ini";
            if (Template.Text == "")
            {
                MessageBox.Show("Please select a template");
            }
            else
            {
                if (File.Exists(Path))
                {
                    File.Delete(Path);
                }
                if (!Directory.Exists(FolderPath))
                {
                    Directory.CreateDirectory(FolderPath);
                }
                using (var TemplateINI = new StreamWriter(Path, true))
                {
                    TemplateINI.WriteLine(Template.Text);
                    TemplateINI.Close();
                }

                this.Close();
            }

        }

        private void Browse_Click(object sender, EventArgs e)
        {
            string Path = FolderPath + "\\Template.ini";
            OpenFileDialog BrowseDialog = new OpenFileDialog();
            BrowseDialog.Title = "Select your Excel template";
            BrowseDialog.InitialDirectory = Path;
            BrowseDialog.Filter = "Excel files (*.xlsx)|*.xlsx|Excel files (*.xls)|*.xls|All Files (*.*)|*.*";
            BrowseDialog.FilterIndex = 1;
            BrowseDialog.RestoreDirectory = true;
            if (BrowseDialog.ShowDialog() == DialogResult.OK)
            {
                Template.Text = BrowseDialog.FileName;
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
