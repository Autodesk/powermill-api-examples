// -----------------------------------------------------------------------
// Copyright 2018 Autodesk, Inc. All rights reserved.
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
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Text.RegularExpressions;

namespace DrillingAutomation
{
    public partial class Options : Form
    {
        public string INIPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\DrillingAutomation";
        public List<MachineCells> MachineCells = null;
        public Options()
        {
            InitializeComponent();

            string Path = INIPath + "\\Settings.ini";
            if (File.Exists(Path))
            {
                TagHoles.ExtractINIData(out MachineCells);

                MachineCellCombo.Items.Clear();
                foreach (MachineCells Machine in MachineCells)
                {
                    MachineCellCombo.Items.Add(Machine.Name);
                }
                MachineCellCombo.SelectedIndex = 0;
                UpdateFields(MachineCells);
            }
            else
            {
                MachineCellCombo.Text = "Default";
            }


        }

        private void OK_Click(object sender, EventArgs e)
        {
            bool AllParametersOK = false;
            string Path = INIPath + "\\Settings.ini";
            if (Database.Text !="" && Toolpath_Folder.Text != "" && Tolerance.Text != "")
            {
                AllParametersOK = true;
            }
            if (!AllParametersOK)
            {
                MessageBox.Show("Make sure you select a database, a toolpath folder and set a tolerance");
            }
            else
            {
                if (File.Exists(Path))
                {
                    string Line = null;
                    string TempFileName = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\DrillingAutomation\\Temp.ini";
                    bool FoundMachine = false;
                    bool NoMachine = true;

                    //Get the template from the option page
                    TagHoles.ExtractINIData(out MachineCells);

                    if (File.Exists(TempFileName))
                    {
                        File.Delete(TempFileName);
                    }
                    File.Copy(Path, TempFileName);
                    File.Delete(Path);

                    using (StreamReader reader = new StreamReader(TempFileName))
                    {
                        using (StreamWriter writer = new StreamWriter(Path))
                        {
                            while ((Line = reader.ReadLine()) != null)
                            {
                                if (Line != "")
                                {
                                    if (Line.IndexOf(MachineCellCombo.Text) >= 0)
                                    {
                                        FoundMachine = true;
                                        NoMachine = false;
                                        writer.WriteLine(Line);
                                        writer.WriteLine("Database=" + Database.Text);
                                        writer.WriteLine("ToolpathFolder=" + Toolpath_Folder.Text);
                                        writer.WriteLine("Tolerance=" + Tolerance.Text);
                                        if (AllowTagDuplicate.Checked)
                                        {
                                            writer.WriteLine("AllowTagDuplicate=True" );
                                        }
                                        else
                                        {
                                            writer.WriteLine("AllowTagDuplicate=False" );
                                        }
                                        if (UseMethods.Checked)
                                        {
                                            writer.WriteLine("ToolpathORMethods=Methods");
                                        }
                                        else
                                        {
                                            writer.WriteLine("ToolpathORMethods=Toolpaths");
                                        }
                                        if (TopComponent.Checked)
                                        {
                                            writer.WriteLine("DepthReference=TopComponent");
                                        }
                                        else
                                        {
                                            writer.WriteLine("DepthReference=TopHole");
                                        }
                                    }
                                    else if (Line.IndexOf("[") >= 0)
                                    {
                                        FoundMachine = false;
                                        writer.WriteLine(Line);
                                    }
                                    else
                                    {
                                        if (!FoundMachine)
                                        {
                                            writer.WriteLine(Line);
                                        }
                                    }
                                }
                            }
                            if (NoMachine)
                            {
                                if (MessageBox.Show("No machine cell with the same name was found, do you want to add a new one?", "Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                                {
                                    writer.WriteLine("[" + MachineCellCombo.Text + "]");
                                    writer.WriteLine("Database=" + Database.Text);
                                    writer.WriteLine("ToolpathFolder=" + Toolpath_Folder.Text);
                                    writer.WriteLine("Tolerance=" + Tolerance.Text);
                                    if (AllowTagDuplicate.Checked)
                                    {
                                        writer.WriteLine("AllowTagDuplicate=True");
                                    }
                                    else
                                    {
                                        writer.WriteLine("AllowTagDuplicate=False");
                                    }
                                    if (UseMethods.Checked)
                                    {
                                        writer.WriteLine("ToolpathORMethods=Methods");
                                    }
                                    else
                                    {
                                        writer.WriteLine("ToolpathORMethods=Toolpaths");
                                    }
                                    if (TopComponent.Checked)
                                    {
                                        writer.WriteLine("DepthReference=TopComponent");
                                    }
                                    else
                                    {
                                        writer.WriteLine("DepthReference=TopHole");
                                    }
                                }
                            }
                            writer.Close();
                        }
                        reader.Close();
                    }
                    File.Delete(TempFileName);
                }
                else
                {
                    using (StreamWriter writer = new StreamWriter(Path))
                    {
                        writer.WriteLine("[" + MachineCellCombo.Text + "]");
                        writer.WriteLine("Database=" + Database.Text);
                        writer.WriteLine("ToolpathFolder=" + Toolpath_Folder.Text);
                        writer.WriteLine("Tolerance=" + Tolerance.Text);
                        if (AllowTagDuplicate.Checked)
                        {
                            writer.WriteLine("AllowTagDuplicate=True");
                        }
                        else
                        {
                            writer.WriteLine("AllowTagDuplicate=False");
                        }
                        if (UseMethods.Checked)
                        {
                            writer.WriteLine("ToolpathORMethods=Methods");
                        }
                        else
                        {
                            writer.WriteLine("ToolpathORMethods=Toolpaths");
                        }
                        if (TopComponent.Checked)
                        {
                            writer.WriteLine("DepthReference=TopComponent");
                        }
                        else
                        {
                            writer.WriteLine("DepthReference=TopHole");
                        }
                        writer.Close();
                    }
                }
                
            }
            this.Close();
        }

        private void BrowseDB_Click(object sender, EventArgs e)
        {
            string Path = INIPath + "\\Settings.ini";
            OpenFileDialog BrowseDialog = new OpenFileDialog();
            BrowseDialog.Title = "Select your Excel template";
            BrowseDialog.InitialDirectory = Path;
            BrowseDialog.Filter = "CSV files (*.CSV)|*.csv|All Files (*.*)|*.*";
            BrowseDialog.FilterIndex = 1;
            BrowseDialog.RestoreDirectory = true;
            if (BrowseDialog.ShowDialog() == DialogResult.OK)
            {
                Database.Text = BrowseDialog.FileName;
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Browse_TP_Click(object sender, EventArgs e)
        {
            string FolderPath = "";
            CommonOpenFileDialog BrowseTemplate = new CommonOpenFileDialog();
            BrowseTemplate.InitialDirectory = "C:\\Users";
            BrowseTemplate.IsFolderPicker = true;
            if (BrowseTemplate.ShowDialog() == CommonFileDialogResult.Ok)
            {
                FolderPath = BrowseTemplate.FileName;
            }
            Toolpath_Folder.Text = FolderPath;
        }

        private void UpdateFields(List<MachineCells> MachineCells)
        {
            string CurrentMachineCell = "";
            if (MachineCellCombo.SelectedItem == null)
            {
                CurrentMachineCell = "Default";
            }
            else
            {
                CurrentMachineCell = MachineCellCombo.SelectedItem.ToString();
            }

            TagHoles.ExtractINIData(out MachineCells);
            foreach (MachineCells Machine in MachineCells)
            {
                if (Machine.Name == CurrentMachineCell)
                {
                    Database.Text = Machine.CSVPath;
                    Toolpath_Folder.Text = Machine.ToolpathsPath;
                    Tolerance.Text = Machine.Tolerance.ToString();
                    if (Machine.AllowTagDuplicate)
                    {
                        AllowTagDuplicate.Checked = true;
                    }
                    else
                    {
                        AllowTagDuplicate.Checked = false;
                    }
                    if (Machine.UseMethod)
                    {
                        UseMethods.Checked = true;
                        UseToolpaths.Checked = false;
                    }
                    else
                    {
                        UseMethods.Checked = false;
                        UseToolpaths.Checked = true;
                    }
                    if (Machine.DepthFromTop)
                    {
                        TopComponent.Checked = false;
                        TopHole.Checked = true;
                    }
                    else
                    {
                        TopComponent.Checked = true;
                        TopHole.Checked = false;
                    }
                }
            }
        }

        private void MachineCellCombo_DropDownClosed(object sender, EventArgs e)
        {
            UpdateFields(MachineCells);
        }
    }
}
