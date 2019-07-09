// -----------------------------------------------------------------------
// Copyright 2018 Autodesk, Inc. All rights reserved.
//
// Use of this software is subject to the terms of the Autodesk license
// agreement provided at the time of installation or download, or which
// otherwise accompanies this software in either electronic or hard copy form.
// -----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Runtime.InteropServices;
using PowerMILL;
using Delcam.Plugins.Framework;
using Delcam.Plugins.Events;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;

namespace DrillingAutomation
{
    /// <summary>
    /// Interaction logic for ToolSheetPaneWPF.xaml
    /// </summary>
    public partial class DrillingAutomationPaneWPF : UserControl
    {

        private IPluginCommunicationsInterface oComm;

        //public ToolSheetPaneWPF()
        //{
        //    InitializeComponent();

        //}


        public DrillingAutomationPaneWPF(IPluginCommunicationsInterface comms)
        {
            InitializeComponent();

            oComm = comms;
            PowerMILLAutomation.SetVariables(oComm);

            // subscribe to some events
            comms.EventUtils.Subscribe(new EventSubscription("EntityCreated", EntityCreated));
            comms.EventUtils.Subscribe(new EventSubscription("EntityDeleted", EntityDeleted));
            comms.EventUtils.Subscribe(new EventSubscription("ProjectClosed", ProjectClosed));
            comms.EventUtils.Subscribe(new EventSubscription("ProjectOpened", ProjectOpened));
            comms.EventUtils.Subscribe(new EventSubscription("EntityRenamed", EntityRenamed));

            //LoadProjectComponents();


        }


        public void PreInitialise(string locale) { }
        public void ProcessCommand(string Commmand) { }
        public void ProcessEvent(string EventData) { }
        public void SerializeProjectData(string Path, bool Saving) { }
        public void Uninitialise()
        {
            m_services = null;
            GC.Collect();
        }


        private string m_token;
        private PowerMILL.PluginServices m_services;
        private int m_parent_window;
        private DrillingAutomationPaneWPF drillingAutomationPaneWPF;
        public List<TaggedHoles> UnRecognizedHolesList = null;
        public List<NewHoleInfos> HoleData = null;
        public List<DBHoleExport> DBHoleExport = null;
        public List<DBHoleInfos> DBHoleData = null;
        public List<MachineCells> MachineCells = null;
        public string CSVPath = "";
        public static string CurrentMachineCell;
        private void LoadProjectComponents()
        {
            double Tolerance = 0;
            bool UseMethod = false;
            string TemplateFolder = "";
            string CSVPath = "";
            bool DepthFromTop = false;
            bool AllowTagDuplicate = true;

            listHoleFeatureSet.Items.Clear();

            List<string> ncprogs = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.FeatureSets);
            foreach (string ncprog in ncprogs)
                listHoleFeatureSet.Items.Add(ncprog);

            //Get the template from the option page
            
            TagHoles.ExtractINIData(out MachineCells);
            TagHoles.GetActiveMachineCellInfos(MachineCell_Drop.Text, MachineCells, out CSVPath, out TemplateFolder, out Tolerance, out UseMethod, out DepthFromTop, out AllowTagDuplicate);
            UpdateMachineCellsDropDown(MachineCells);
            CurrentMachineCell = MachineCell_Drop.Text;
        }

        void ProjectOpened(string event_name, Dictionary<string, string> event_arguments)
        {
            LoadProjectComponents();
        }

        void EntityCreated(string event_name, Dictionary<string, string> event_arguments)
        {
            if (MachineCell_Drop.Items.Count == 0)
            {
                UpdateMachineCellsDropDown(MachineCells);
                MachineCell_Drop.SelectedIndex = 0;
            }
            string entity_type = event_arguments["EntityType"];
            string entity_name = event_arguments["Name"];

            switch (entity_type)
            {
                case "Model":
                    break;

                case "Ncprogram":
                    break;

                case "Stockmodel":
                    break;

                case "Toolpath":
                    break;

                case "Workplane":
                    break;

                case "Featureset":
                    listHoleFeatureSet.Items.Add(entity_name);
                    NCProgramListDB.Items.Add(entity_name);
                    break;

            }
        }

        void EntityRenamed(string event_name, Dictionary<string, string> event_arguments)
        {
            string entity_type = event_arguments["EntityType"];
            string entity_orig_name = event_arguments["PreviousName"];
            string entity_new_name = event_arguments["Name"];

            switch (entity_type)
            {
                case "Model":
                    break;

                case "Ncprogram":
                    break;

                case "Stockmodel":
                    break;

                case "Toolpath":
                    break;

                case "Workplane":
                    break;
                case "Featureset":
                    if (listHoleFeatureSet.Items.Contains(entity_orig_name))
                    {
                        listHoleFeatureSet.Items.Remove(entity_orig_name);
                        listHoleFeatureSet.Items.Add(entity_new_name);
                    }
                    if (listHoleFeatureSetSelected.Items.Contains(entity_orig_name))
                    {
                        listHoleFeatureSetSelected.Items.Remove(entity_orig_name);
                        listHoleFeatureSetSelected.Items.Add(entity_new_name);
                    }
                    if (NCProgramListDB.Items.Contains(entity_orig_name))
                    {
                        NCProgramListDB.Items.Remove(entity_orig_name);
                        NCProgramListDB.Items.Add(entity_new_name);
                    }
                    break;
            }
        }

        void EntityDeleted(string event_name, Dictionary<string, string> event_arguments)
        {
            string entity_type = event_arguments["EntityType"];
            string entity_name = event_arguments["Name"];

            switch (entity_type)
            {
                case "Model":
                    break;

                case "Ncprogram":
                    break;

                case "Stockmodel":
                    break;

                case "Toolpath":
                    break;

                case "Workplane":
                    break;

                case "Featureset":
                    if (listHoleFeatureSet.Items.Contains(entity_name))
                        listHoleFeatureSet.Items.Remove(entity_name);
                    else if (listHoleFeatureSetSelected.Items.Contains(entity_name))
                        listHoleFeatureSetSelected.Items.Remove(entity_name);
                    if (NCProgramListDB.Items.Contains(entity_name))
                    {
                        NCProgramListDB.Items.Remove(entity_name);
                    }
                    break;

            }
        }

        void ProjectClosed(string event_name, Dictionary<string, string> event_arguments)
        {
            listHoleFeatureSet.Items.Clear();
        }

        private void AddFeatureSet_Click(object sender, RoutedEventArgs e)
        {
            if (listHoleFeatureSet.SelectedItems.Count > 0)
            {
                for (int i = listHoleFeatureSet.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object SelectedItem = listHoleFeatureSet.SelectedItems[i];
                    listHoleFeatureSetSelected.Items.Insert(0, SelectedItem); // to keep the same order
                    listHoleFeatureSet.Items.Remove(SelectedItem);
                }
            }
        }

        private void RemoveFeatureSet_Click(object sender, RoutedEventArgs e)
        {
            if (listHoleFeatureSetSelected.SelectedItems.Count > 0)
            {
                for (int i = listHoleFeatureSetSelected.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object SelectedItem = listHoleFeatureSetSelected.SelectedItems[i];
                    listHoleFeatureSet.Items.Insert(0, SelectedItem); // to keep the same order...
                    listHoleFeatureSetSelected.Items.Remove(SelectedItem);
                }
            }
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            string TemplateFolder = "";
            string Path = System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Substring(8, System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Length - 8);
            string directory = System.IO.Path.GetDirectoryName(Path);
            string RenameMacro = directory + "\\Rename.mac";
            string EndOfTagging = directory + "\\EndOfTagging.mac";
            double Tolerance = 0;
            bool UseMethod = false;
            bool AllowTagDuplicate = true;
            List<RecognizedHoles> RecognizedHolesList = null;


            //Get the template from the option page
            TagHoles.ExtractINIData(out MachineCells);
            TagHoles.GetActiveMachineCellInfos(MachineCell_Drop.Text, MachineCells, out CSVPath, out TemplateFolder, out Tolerance, out UseMethod, out bool DepthFromTop, out AllowTagDuplicate);


            foreach (string FeatureSet in listHoleFeatureSetSelected.Items)
            {
                PowerMILLAutomation.ExecuteEx("ACTIVATE Featureset '" + FeatureSet + "'");

                if (TagOnly.IsChecked == true || TagAndTP.IsChecked == true)
                {
                    PowerMILLAutomation.ExecuteEx("MACRO '" + RenameMacro + "'");
                    TagHoles.ExtractFeatureSetData(CSVPath, DepthFromTop, out HoleData, out List<DBHoleInfos> DBHoleData);
                    TagHoles.MatchAndRenameHolesNew(FeatureSet, HoleData, DBHoleData, Tolerance, out RecognizedHolesList, out UnRecognizedHolesList);
                    PowerMILLAutomation.ExecuteEx("MACRO '" + EndOfTagging + "'");
                }
                if (TPOnly.IsChecked == true || TagAndTP.IsChecked == true)
                {
                    TagHoles.GenerateToolpaths(directory, TemplateFolder, UseMethod);
                }

                MessageBox.Show("Process completed","Hole Tagging Plugin");
            }
        }

        public static FileInfo GetNewestFile(DirectoryInfo directory)
        {
            return directory.GetFiles()
                .Union(directory.GetDirectories().Select(d => GetNewestFile(d)))
                .OrderByDescending(f => (f == null ? DateTime.MinValue : f.LastWriteTime))
                .FirstOrDefault();
        }

        public void UpdateMachineCellsDropDown(List<MachineCells> MachineCells)
        {
            TagHoles.ExtractINIData(out MachineCells);
            if (MachineCell_Drop.Items.Count == 1)
            {
                foreach (MachineCells Machine in MachineCells)
                {
                    if (Machine.Name != "Default")
                    {
                        MachineCell_Drop.Items.Add(Machine.Name);
                    }
                }
            }
            else if (MachineCells.Count != MachineCell_Drop.Items.Count)
            {
                MachineCell_Drop.Items.Clear();
                foreach (MachineCells Machine in MachineCells)
                {
                    MachineCell_Drop.Items.Add(Machine.Name);
                }
            }
        }

        private void DatabaseTag_Click(object sender, RoutedEventArgs e)
        {
            string TemplateFolder = "";
            string Path = System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Substring(8, System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Length - 8);
            string directory = System.IO.Path.GetDirectoryName(Path);
            string RenameMacro = directory + "\\Rename.mac";
            string EndOfTagging = directory + "\\EndOfTagging.mac";
            double Tolerance = 0;
            bool UseMethod = false;
            bool DepthFromTop = false;
            bool AllowTagDuplicate = true;
            List<RecognizedHoles> RecognizedHolesList = null;

            if (NCProgramListDB.Text == "")
            {
                MessageBox.Show("Please select a Featureset", "Error");
                return;
            }

            //Get the template from the option page

            TagHoles.ExtractINIData(out MachineCells);
            TagHoles.GetActiveMachineCellInfos(MachineCell_Drop.Text, MachineCells, out CSVPath, out TemplateFolder, out Tolerance, out UseMethod, out DepthFromTop, out AllowTagDuplicate);

            PowerMILLAutomation.ExecuteEx("ACTIVATE Featureset '" + NCProgramListDB.Text + "'");
            PowerMILLAutomation.ExecuteEx("MACRO '" + RenameMacro + "'");
            TagHoles.ExtractFeatureSetData(CSVPath, DepthFromTop, out HoleData, out List<DBHoleInfos> DBHoleData);
            TagHoles.MatchAndRenameHolesNew(NCProgramListDB.Text, HoleData, DBHoleData, Tolerance, out RecognizedHolesList, out UnRecognizedHolesList);
            FillUnrecognizedHolePullDown();
            PowerMILLAutomation.ExecuteEx("MACRO '" + EndOfTagging + "'");

        }
        private void FillUnrecognizedHolePullDown()
        {
            UnrecognizedHolesCombo.Items.Clear();
            foreach (TaggedHoles Hole in UnRecognizedHolesList)
            {
                UnrecognizedHolesCombo.Items.Add(Hole.HoleName);
            }
        }
        private void UnrecognizedHolesCombo_DropDownClosed(object sender, EventArgs e)
        {
            DBHoleExport = new List<DBHoleExport>();

            if (UnRecognizedHolesList != null)
            {
                if (UnrecognizedHolesCombo.Items.Count > 0 && UnrecognizedHolesCombo.Text != "")
                {
                    HoleComponentCombo.Items.Clear();
                    foreach (TaggedHoles UnknownHole in UnRecognizedHolesList)
                    {
                        foreach (NewHoleInfos Hole in HoleData)
                        {
                            if (UnknownHole.HoleTag == Hole.HoleNumber.ToString() && UnknownHole.HoleName == UnrecognizedHolesCombo.Text)
                            {
                                if (Hole.ComponentID != -1)
                                {
                                    HoleComponentCombo.Items.Add(Hole.ComponentID);
                                }
                                DBHoleExport.Add(new DBHoleExport
                                {
                                    UpperDiameter = Hole.UpperDiameter,
                                    LowerDiameter = Hole.LowerDiameter,
                                    Depth = Hole.Depth,
                                    RColor = Hole.RColor,
                                    GColor = Hole.GColor,
                                    BColor = Hole.BColor,
                                    ComponentID = Hole.ComponentID
                                });
                            }
                        }
                    }
                    HoleComponentCombo.SelectedIndex = 0;
                    UpdateFeatureSetForm(UnrecognizedHolesCombo.Text, 0);
                }
             }
        }

        private void HoleComponentCombo_DropDownClosed(object sender, EventArgs e)
        {
            if (HoleComponentCombo != null && HoleComponentCombo.Text != "")
            {
                UpdateFeatureSetForm(UnrecognizedHolesCombo.Text, int.Parse(HoleComponentCombo.Text));
                LoadComponentData(int.Parse(HoleComponentCombo.Text),false);
            }
        }

        public void SaveComponentData(int ComponentID, bool Modify)
        {
            double CurrentColor = 0;
            if (HoleTagADD.Text != "" || Modify)
            {
                foreach (DBHoleExport HoleExport in DBHoleExport)
                {
                    if (HoleExport.ComponentID == ComponentID)
                    {
                        if (Modify)
                        {
                            if (double.Parse(MinDepthMOD.Text) > double.Parse(MaxDepthMOD.Text))
                            {
                                MessageBox.Show("Minimum value must be smaller than Maximum", "Error");
                                return;
                            }
                            if (double.Parse(MinDepthMOD.Text) < 0 || double.Parse(MaxDepthMOD.Text) < 0)
                            {
                                MessageBox.Show("Min. and Max. Depth can't be negative", "Error");
                                return;
                            }
                            if (HoleDescriptionMOD.Text == "")
                            {
                                MessageBox.Show("Description must be set", "Error");
                                return;
                            }
                            HoleExport.Tag = HoleTagMOD.Text;
                            HoleExport.Description = HoleDescriptionMOD.Text;
                            HoleExport.MinDepth = double.Parse(MinDepthMOD.Text);
                            HoleExport.MaxDepth = double.Parse(MaxDepthMOD.Text);
                            HoleExport.UpperDiameter = double.Parse(UpperDiameterMOD.Text);
                            HoleExport.LowerDiameter = double.Parse(LowerDiameterMOD.Text);
                            HoleExport.Depth = double.Parse(HoleDepthMOD.Text);
                            HoleExport.Family = HoleFamilyMOD.Text;
                            if (UseColors_MOD.IsChecked == true)
                            {
                                double.TryParse(R_Color.Text, out CurrentColor);
                                HoleExport.RColor = CurrentColor;
                                double.TryParse(G_Color.Text, out CurrentColor);
                                HoleExport.GColor = CurrentColor;
                                double.TryParse(B_Color.Text, out CurrentColor);
                                HoleExport.BColor = CurrentColor;
                            }
                            else
                            {
                                HoleExport.RColor = -1;
                                HoleExport.GColor = -1;
                                HoleExport.BColor = -1;
                            }
                        }
                        else
                        {
                            if (double.Parse(MinDepthADD.Text) > double.Parse(MaxDepthADD.Text))
                            {
                                MessageBox.Show("Minimum value must be smaller than Maximum", "Error");
                                return;
                            }
                            if (double.Parse(MinDepthADD.Text) < 0 || double.Parse(MaxDepthADD.Text) < 0)
                            {
                                MessageBox.Show("Min. and Max. Depth can't be negative", "Error");
                                return;
                            }
                            if (HoleDescriptionADD.Text == "")
                            {
                                MessageBox.Show("Description must be set", "Error");
                                return;
                            }
                            HoleExport.Tag = HoleTagADD.Text;
                            HoleExport.Description = HoleDescriptionADD.Text;
                            HoleExport.MinDepth = double.Parse(MinDepthADD.Text);
                            HoleExport.MaxDepth = double.Parse(MaxDepthADD.Text);
                            HoleExport.UpperDiameter = double.Parse(UpperDiameterADD.Text);
                            HoleExport.LowerDiameter = double.Parse(LowerDiameterADD.Text);
                            HoleExport.Depth = double.Parse(HoleDepthADD.Text);
                            HoleExport.Family = HoleFamilyADD.Text;
                            if (UseColors.IsChecked == true)
                            {
                                double.TryParse(R_Color.Text, out CurrentColor);
                                HoleExport.RColor = CurrentColor;
                                double.TryParse(G_Color.Text, out CurrentColor);
                                HoleExport.GColor = CurrentColor;
                                double.TryParse(B_Color.Text, out CurrentColor);
                                HoleExport.BColor = CurrentColor;
                            }
                            else
                            {
                                HoleExport.RColor = -1;
                                HoleExport.GColor = -1;
                                HoleExport.BColor = -1;
                            }
                        }

                    }
                }
            }
        }
        public void LoadComponentData(int ComponentID, bool Modify)
        {
            if (HoleTagADD.Text != "" || Modify)
            {
                foreach (DBHoleExport HoleExport in DBHoleExport)
                {
                    if (HoleExport.ComponentID == ComponentID)
                    {
                        if (Modify)
                        {
                            MinDepthMOD.Text = HoleExport.MinDepth.ToString();
                            MaxDepthMOD.Text = HoleExport.MaxDepth.ToString();
                            UpperDiameterMOD.Text = HoleExport.UpperDiameter.ToString();
                            LowerDiameterMOD.Text = HoleExport.LowerDiameter.ToString();
                            HoleDepthMOD.Text = HoleExport.Depth.ToString();
                            HoleTagMOD.Text = HoleExport.Tag;
                            HoleDescriptionMOD.Text = HoleExport.Description;
                            HoleFamilyMOD.Text = HoleExport.Family;
                            if (HoleExport.RColor == -1)
                            {
                                RGB_Field_MOD.IsEnabled = false;
                                R_Color_MOD.IsEnabled = false;
                                G_Color_MOD.IsEnabled = false;
                                B_Color_MOD.IsEnabled = false;
                                UseColors_MOD.IsChecked = false;
                                R_Color_MOD.Text = "-1";
                                G_Color_MOD.Text = "-1";
                                B_Color_MOD.Text = "-1";
                            }
                            else
                            {
                                RGB_Field_MOD.IsEnabled = true;
                                R_Color_MOD.IsEnabled = true;
                                G_Color_MOD.IsEnabled = true;
                                B_Color_MOD.IsEnabled = true;
                                UseColors_MOD.IsChecked = true;
                                R_Color_MOD.Text = HoleExport.RColor.ToString();
                                G_Color_MOD.Text = HoleExport.GColor.ToString();
                                B_Color_MOD.Text = HoleExport.BColor.ToString();
                            }


                        }
                        else
                        {
                            MinDepthADD.Text = HoleExport.MinDepth.ToString();
                            MaxDepthADD.Text = HoleExport.MaxDepth.ToString();
                            UpperDiameterADD.Text = HoleExport.UpperDiameter.ToString();
                            LowerDiameterADD.Text = HoleExport.LowerDiameter.ToString();
                            HoleDepthADD.Text = HoleExport.Depth.ToString();
                            R_Color.Text = HoleExport.RColor.ToString();
                            G_Color.Text = HoleExport.GColor.ToString();
                            B_Color.Text = HoleExport.BColor.ToString();
                        }
                    }
                }
            }
        }
        private void UpdateFeatureSetForm(string ComponentName, int ComponentID)
        {
            int ComponentList = 0;
            foreach (TaggedHoles UnknownHole in UnRecognizedHolesList)
            {
                ComponentList = 0;
                foreach (NewHoleInfos Hole in HoleData)
                {
                    if (UnknownHole.HoleTag == Hole.HoleNumber.ToString() && UnknownHole.HoleName == UnrecognizedHolesCombo.Text && ComponentID == Hole.ComponentID)
                    {
                        UpperDiameterADD.Text = Math.Round(Hole.UpperDiameter, 4).ToString();
                        HoleDepthADD.Text = Math.Round(Hole.Depth, 4).ToString();
                        LowerDiameterADD.Text = Math.Round(Hole.LowerDiameter, 4).ToString();
                        R_Color.Text = Hole.RColor.ToString();
                        G_Color.Text = Hole.GColor.ToString();
                        B_Color.Text = Hole.BColor.ToString();
                    }
                    ComponentList = ComponentList + 1;
                }
            }
            PowerMILLAutomation.ExecuteEx("EDIT FEATURESET ; DESELECT ALL");
            PowerMILLAutomation.ExecuteEx("EDIT FEATURESET ; SELECT '" + ComponentName + "'");
            PowerMILLAutomation.ExecuteEx("ACTIVATE FEATURESET ; FORM EDITHOLE");
            PowerMILLAutomation.ExecuteEx("EDIT FEATURESET ; FEATURE SELECTION " + ComponentID);

        }

        private void HoleComponentCombo_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Down || e.Key == Key.Up)
            {
                UpdateFeatureSetForm(UnrecognizedHolesCombo.Text, int.Parse(HoleComponentCombo.Text));
            }

        }
        private void AddToDatabase_Click(object sender, RoutedEventArgs e)
        {
            string TemplateFolder = "";
            string Path = System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Substring(8, System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Length - 8);
            string directory = System.IO.Path.GetDirectoryName(Path);
            double Tolerance = 0;
            bool UseMethod = false;
            bool DepthFromTop = false;
            bool AllowTagDuplicate = false;
            string FullLine = "";

            SaveComponentData(int.Parse(HoleComponentCombo.Text),false);

            //Get the template from the option page
            TagHoles.ExtractINIData(out MachineCells);
            TagHoles.GetActiveMachineCellInfos(MachineCell_Drop.Text, MachineCells, out CSVPath, out TemplateFolder, out Tolerance, out UseMethod, out DepthFromTop, out AllowTagDuplicate);

            if (NCProgramListDB.Text == "")
            {
                MessageBox.Show("Please select a Featureset", "Error");
                return;
            }

            if (UnrecognizedHolesCombo.Text == "")
            {
                MessageBox.Show("Please select an unrecognized hole", "Error");
                return;
            }

            if (HoleTagADD.Text == "")
            {
                MessageBox.Show("Please set a tag name", "Error");
                return;
            }

            if (HoleDescriptionADD.Text == "")
            {
                MessageBox.Show("Please set a description", "Error");
                return;
            }

            PowerMILLAutomation.ExecuteEx("FORM CANCEL EDITHOLE");

            SaveComponentData(int.Parse(HoleComponentCombo.Text),false);

            HoleFullSignature(out FullLine, false);
            SearchForExistingHole(CSVPath, FullLine, false, AllowTagDuplicate, out bool HoleExist);
            if (HoleExist)
            {
                if (AllowTagDuplicate)
                {
                    MessageBox.Show("Hole already exist in the database", "Error");
                }
                else
                {
                    MessageBox.Show("Hole tag already exist in the database", "Error");
                }
            }
            else
            {
                File.AppendAllText(CSVPath, Environment.NewLine + FullLine);
                CleanBlankSpaces(CSVPath);
                MessageBox.Show("Hole Successfully added to the database", "Information");
            }
        }

        public void CleanBlankSpaces(string CSVPath)
        {
            string TempFileName = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\DrillingAutomation\\Temp.csv";
            string Line = null;
            if (File.Exists(TempFileName))
            {
                File.Delete(TempFileName);
            }
            File.Copy(CSVPath, TempFileName);
            File.Delete(CSVPath);

            using (StreamReader reader = new StreamReader(TempFileName))
            {
                using (StreamWriter writer = new StreamWriter(CSVPath))
                {
                    while ((Line = reader.ReadLine()) != null)
                    {
                        if (Line != "")
                        {
                             writer.WriteLine(Line);
                        }
                    }
                    writer.Close();
                }
                reader.Close();
            }
            File.Delete(TempFileName);
        }
        public void SearchForExistingHole(string CSVPath, string FullLine, bool Modify, bool AllowTagDuplicate, out bool HoleExist)
        {
            HoleExist = false;

            if (File.Exists(CSVPath))
            {
                string TempFileName = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\DrillingAutomation\\Tempcsv.csv";
                if (File.Exists(TempFileName))
                {
                    File.Delete(TempFileName);
                }
                File.Copy(CSVPath, TempFileName);

                const Int32 BufferSize = 128;
                using (var fileStream = File.OpenRead(TempFileName))
                using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
                {
                    String Line;
                    while ((Line = streamReader.ReadLine()) != null)
                    {
                        if (AllowTagDuplicate)
                        {
                            if (Line == FullLine)
                            {
                                HoleExist = true;
                                break;
                            }
                        }
                        else
                        {
                            if (Modify)
                            {
                                if (Line.IndexOf(HoleTagMOD.Text) >= 0)
                                {
                                    HoleExist = true;
                                    break;
                                }
                            }
                            else
                            {
                                if (Line.IndexOf(HoleTagADD.Text) >= 0)
                                {
                                    HoleExist = true;
                                    break;
                                }
                            }
                        }
                    }
                }
                File.Delete(TempFileName);
            }
        }
        public void HoleFullSignature(out string FullLine, bool Modify)
        {
            bool FirstParameter = true;
            FullLine = "";

            foreach (DBHoleExport HoleExport in DBHoleExport)
            {
                if (FirstParameter)
                {
                    FirstParameter = false;
                }
                else
                {
                    FullLine = FullLine + ",";
                }

                FullLine = FullLine + "UD=" + Math.Round(HoleExport.UpperDiameter, 4) + " LD=" + Math.Round(HoleExport.LowerDiameter, 4) + " Depth=" + Math.Round(HoleExport.Depth, 4);

                if (HoleExport.MinDepth != 0 || HoleExport.MaxDepth != 0)
                {
                    FullLine = FullLine + " MinDepth=" + Math.Round(HoleExport.MinDepth, 4) + " MaxDepth=" + Math.Round(HoleExport.MaxDepth, 4);
                }
                if (Modify)
                {
                    if (UseColors_MOD.IsChecked == true)
                    {
                        FullLine = FullLine + " Color=R" + HoleExport.RColor.ToString() + "G" + HoleExport.GColor.ToString() + "B" + HoleExport.BColor.ToString();
                    }
                }
                else
                {
                    if (UseColors.IsChecked == true)
                    {
                        FullLine = FullLine + " Color=R" + HoleExport.RColor.ToString() + "G" + HoleExport.GColor.ToString() + "B" + HoleExport.BColor.ToString();
                    }
                }

            }

            if (Modify)
            {
                FullLine = FullLine + ";" + HoleTagMOD.Text + ";" + HoleDescriptionMOD.Text + ";" + HoleFamilyMOD.Text;
            }
            else
            {
                FullLine = FullLine + ";" + HoleTagADD.Text + ";" + HoleDescriptionADD.Text + ";" + HoleFamilyADD.Text;
            }
            
            
        }

        private void HoleComponentCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (HoleComponentCombo.Text != "")
            {
                SaveComponentData(int.Parse(HoleComponentCombo.Text),false);
            }
        }

        public void UpdateHoleTypeList(string HoleTag, string HoleFamily, string FamilyFilter)
        {
            if (!HoleTypes.Items.Contains(HoleTag))
            {
                if (FamilyFilter == "All")
                {
                    HoleTypes.Items.Add(HoleTag);
                }
                else
                {
                    if (HoleFamily == FamilyFilter && !HoleTypes.Items.Contains(HoleFamily))
                    {
                        HoleTypes.Items.Add(HoleTag);
                    }
                }
            }
        }

        private void HoleFamilies_DropDownClosed(object sender, EventArgs e)
        {
            if (HoleFamilies.Text != "")
            {
                HoleTypes.Items.Clear();
                foreach (DBHoleInfos Hole in DBHoleData)
                {
                    UpdateHoleTypeList(Hole.HoleTag, Hole.Family, HoleFamilies.Text);
                }
            }
            ClearForm();
        }

        private void DelModify_MouseLeftButtonUp(object sender, MouseButtonEventArgs e)
        {
            DBHoleData = new List<DBHoleInfos>();
            string TemplateFolder = "";
            string Path = System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Substring(8, System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Length - 8);
            string directory = System.IO.Path.GetDirectoryName(Path);
            double Tolerance = 0;
            bool UseMethod = false;
            bool DepthFromTop = false;
            bool AllowTagDuplicate = true;


            //Get the template from the option page
            TagHoles.ExtractINIData(out MachineCells);
            TagHoles.GetActiveMachineCellInfos(MachineCell_Drop.Text, MachineCells, out CSVPath, out TemplateFolder, out Tolerance, out UseMethod, out DepthFromTop, out AllowTagDuplicate);
            TagHoles.ExtractNewDatabaseData(CSVPath, out DBHoleData);

            LoadFamilyTypes();
        }

        private void HoleTypes_DropDownClosed(object sender, EventArgs e)
        {
            LoadHoleTypes();
        }

        private void HoleComponentComboMod_DropDownClosed(object sender, EventArgs e)
        {
            if (HoleComponentComboMod.Text != "")
            {
                LoadComponentData(int.Parse(HoleComponentComboMod.Text), true);
            }

        }
        private void HoleComponentComboMod_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (HoleComponentComboMod.Items.Count > 0 && HoleTagMOD.Text != "")
            {
                if (HoleComponentComboMod.Text != "")
                {
                    SaveComponentData(int.Parse(HoleComponentComboMod.Text), true);
                }
                else
                {
                    SaveComponentData(0, true);
                }
                
            }
        }
        private void ModifyInDatabase_Click(object sender, RoutedEventArgs e)
        {
            int HoleComponent = 0;
            int.TryParse(HoleComponentComboMod.Text,out HoleComponent);
            SaveComponentData(HoleComponent,true);
            UpdateCSVFile(true);
            LoadFamilyTypes();
            ClearForm();
        }

        private void DeleteFromDatabase_Click(object sender, RoutedEventArgs e)
        {
            UpdateCSVFile(false);
            LoadFamilyTypes();
            ClearForm();

        }

        public void LoadHoleTypes()
        {
            DBHoleExport = new List<DBHoleExport>();
            int ComponentID = 0;
            HoleComponentComboMod.Items.Clear();
            foreach (DBHoleInfos Hole in DBHoleData)
            {
                if (Hole.HoleTag == HoleTypes.Text)
                {
                    HoleComponentComboMod.Items.Add(ComponentID);

                    DBHoleExport.Add(new DBHoleExport
                    {
                        UpperDiameter = Hole.UpperDiameter,
                        LowerDiameter = Hole.LowerDiameter,
                        Depth = Hole.HoleDepth,
                        MinDepth = Hole.MinimumDepth,
                        MaxDepth = Hole.MaximumDepth,
                        Description = Hole.HoleDescription,
                        Tag = Hole.HoleTag,
                        Family = Hole.Family,
                        RColor = Hole.RColor,
                        GColor = Hole.GColor,
                        BColor = Hole.BColor,
                        ComponentID = ComponentID
                    });
                    ComponentID = ComponentID + 1;
                }
            }
            //HoleComponentComboMod.SelectedIndex = 0;
            LoadComponentData(0, true);
        }

        public void LoadFamilyTypes()
        {
            HoleFamilies.Items.Clear();
            HoleTypes.Items.Clear();
            HoleFamilies.Items.Add("All");
            HoleFamilies.SelectedIndex = 0;
            foreach (DBHoleInfos Hole in DBHoleData)
            {
                if (!HoleFamilies.Items.Contains(Hole.Family))
                {
                    HoleFamilies.Items.Add(Hole.Family);
                }
                UpdateHoleTypeList(Hole.HoleTag, Hole.Family, "All");
            }
        }
        public void UpdateCSVFile(bool Modify)
        {
            string Line = null;
            string TempFileName = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\DrillingAutomation\\Temp.csv";
            string TemplateFolder = "";
            double Tolerance = 0;
            bool UseMethod = false;
            string FullLine = "";
            bool DepthFromTop = false;
            bool AllowTagDuplicate = true;

            //Get the template from the option page
            TagHoles.ExtractINIData(out MachineCells);
            TagHoles.GetActiveMachineCellInfos(MachineCell_Drop.Text, MachineCells, out CSVPath, out TemplateFolder, out Tolerance, out UseMethod, out DepthFromTop, out AllowTagDuplicate);

            if (File.Exists(TempFileName))
            {
                File.Delete(TempFileName);
            }
            File.Copy(CSVPath, TempFileName);
            File.Delete(CSVPath);

            using (StreamReader reader = new StreamReader(TempFileName))
            {
                using (StreamWriter writer = new StreamWriter(CSVPath))
                {
                    while ((Line = reader.ReadLine()) != null)
                    {
                        if (Line != "")
                        {
                            if (Line.Substring(0, 1) != "*")
                            {
                                // Split the Line
                                string[] CurrentHole = Regex.Split(Line, ";");

                                if (CurrentHole.Count() >= 3)
                                {
                                    if (CurrentHole[1] == HoleTypes.Text)
                                    {
                                        if (Modify)
                                        {
                                            HoleFullSignature(out FullLine, true);
                                            SearchForExistingHole(CSVPath, FullLine, Modify, AllowTagDuplicate, out bool HoleExist);
                                            if (HoleExist)
                                            {
                                                if (AllowTagDuplicate)
                                                {
                                                    MessageBox.Show("Hole already exist in the database", "Error");
                                                }
                                                else
                                                {
                                                    MessageBox.Show("Hole tag already exist in the database", "Error");
                                                }
                                            }
                                            else
                                            {
                                                writer.WriteLine(FullLine);
                                            }
                                        }
                                        continue;
                                    }
                                }
                            }
                            writer.WriteLine(Line);
                        }
                    }
                    writer.Close();
                }
                reader.Close();
            }
            File.Delete(TempFileName);
            TagHoles.ExtractNewDatabaseData(CSVPath, out DBHoleData);
        }

        public void ClearForm()
        {
            HoleTagMOD.Text = "";
            HoleDescriptionMOD.Text = "";
            UpperDiameterMOD.Text = "";
            LowerDiameterMOD.Text = "";
            HoleDepthMOD.Text = "";
            MinDepthMOD.Text = "";
            MaxDepthMOD.Text = "";
            HoleComponentComboMod.Items.Clear();
            HoleFamilyMOD.Text = "";
            UseColors_MOD.IsChecked = false;
            R_Color_MOD.Text = "";
            G_Color_MOD.Text = "";
            B_Color_MOD.Text = "";
            R_Color_MOD.IsEnabled = false;
            G_Color_MOD.IsEnabled = false;
            B_Color_MOD.IsEnabled = false;
        }

        private void UseColors_Checked(object sender, RoutedEventArgs e)
        {
            RGB_Field.IsEnabled = true;
            R_Color.IsEnabled = true;
            G_Color.IsEnabled = true;
            B_Color.IsEnabled = true;
        }

        private void UseColors_Unchecked(object sender, RoutedEventArgs e)
        {
            RGB_Field.IsEnabled = false;
            R_Color.IsEnabled = false;
            G_Color.IsEnabled = false;
            B_Color.IsEnabled = false;
        }

        private void UseColors_MOD_Checked(object sender, RoutedEventArgs e)
        {
            RGB_Field_MOD.IsEnabled = true;
            R_Color_MOD.IsEnabled = true;
            G_Color_MOD.IsEnabled = true;
            B_Color_MOD.IsEnabled = true;
        }

        private void UseColors_MOD_Unchecked(object sender, RoutedEventArgs e)
        {
            RGB_Field_MOD.IsEnabled = false;
            R_Color_MOD.IsEnabled = false;
            G_Color_MOD.IsEnabled = false;
            B_Color_MOD.IsEnabled = false;
        }

        private void HoleTagADD_LostFocus(object sender, RoutedEventArgs e)
        {
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                HoleTagADD.Text = HoleTagADD.Text.Replace(c, '_');
            }
            HoleTagADD.Text = HoleTagADD.Text.Trim();
        }

        private void HoleTagMOD_LostFocus(object sender, RoutedEventArgs e)
        {
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                HoleTagMOD.Text = HoleTagMOD.Text.Replace(c, '_');
            }
        }

        private void MachineCell_Drop_DropDownClosed(object sender, EventArgs e)
        {
            CurrentMachineCell = MachineCell_Drop.Text;
        }

        private void MachineCell_Drop_MouseEnter(object sender, MouseEventArgs e)
        {
            UpdateMachineCellsDropDown(MachineCells);
        }

    }
}
