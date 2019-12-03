// -----------------------------------------------------------------------
// Copyright 2019 Autodesk, Inc. All rights reserved.
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

namespace SetupSheet
{
    /// <summary>
    /// Interaction logic for ToolSheetPaneWPF.xaml
    /// </summary>
    public partial class ToolSheetPaneWPF : UserControl
    {

        private IPluginCommunicationsInterface oComm;

        //public ToolSheetPaneWPF()
        //{
        //    InitializeComponent();

        //}


        public ToolSheetPaneWPF(IPluginCommunicationsInterface comms)
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

            LoadProjectComponents();


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
        private ToolSheetPaneWPF toolSheetPaneWPF;

        private void LoadProjectComponents()
        {
            listNCProgs.Items.Clear();



            List<string> ncprogs = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.NCPrograms);
            foreach (string ncprog in ncprogs)
                listNCProgs.Items.Add(ncprog);
        }

        void ProjectOpened(string event_name, Dictionary<string, string> event_arguments)
        {
            LoadProjectComponents();
        }

        void EntityCreated(string event_name, Dictionary<string, string> event_arguments)
        {
            string entity_type = event_arguments["EntityType"];
            string entity_name = event_arguments["Name"];

            switch (entity_type)
            {
                case "Model":
                    break;

                case "Ncprogram":
                    listNCProgs.Items.Add(entity_name);
                    break;

                case "Stockmodel":
                    break;

                case "Toolpath":
                    break;

                case "Workplane":
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

                    if (listNCProgs.Items.Contains(entity_orig_name))
                    {
                        listNCProgs.Items.Remove(entity_orig_name);
                        listNCProgs.Items.Add(entity_new_name);
                    }
                    if (listNCProgsSelected.Items.Contains(entity_orig_name))
                     {
                         listNCProgsSelected.Items.Remove(entity_orig_name);
                         listNCProgsSelected.Items.Add(entity_new_name);
                     }
                    break;

                case "Stockmodel":
                    break;

                case "Toolpath":
                    break;

                case "Workplane":
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
                    if (listNCProgs.Items.Contains(entity_name))
                        listNCProgs.Items.Remove(entity_name);
                    else if (listNCProgsSelected.Items.Contains(entity_name))
                        listNCProgsSelected.Items.Remove(entity_name);
                    break;

                case "Stockmodel":
                    break;

                case "Toolpath":
                    break;

                case "Workplane":
                    break;

            }
        }

        void ProjectClosed(string event_name, Dictionary<string, string> event_arguments)
        {
            listNCProgs.Items.Clear();
        }

        private void Load_List_Click(object sender, RoutedEventArgs e)
        {
            listNCProgs.Items.Clear();

            List<string> ncprogs = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.NCPrograms);
            foreach (string ncprog in ncprogs)
                listNCProgs.Items.Add(ncprog);
        }

        private void AddNCProg_Click(object sender, RoutedEventArgs e)
        {
            if (listNCProgs.SelectedItems.Count > 0)
            {
                for (int i = listNCProgs.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object SelectedItem = listNCProgs.SelectedItems[i];
                    listNCProgsSelected.Items.Insert(0, SelectedItem); // to keep the same order
                    listNCProgs.Items.Remove(SelectedItem);
                }
            }
        }

        private void RemoveNCProg_Click(object sender, RoutedEventArgs e)
        {
            if (listNCProgsSelected.SelectedItems.Count > 0)
            {
                for (int i = listNCProgsSelected.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object SelectedItem = listNCProgsSelected.SelectedItems[i];
                    listNCProgs.Items.Insert(0, SelectedItem); // to keep the same order...
                    listNCProgsSelected.Items.Remove(SelectedItem);
                }
            }
        }

        private void Generate_Click(object sender, RoutedEventArgs e)
        {
            List<string> NCProg_Tools = new List<string>();
            Microsoft.Office.Interop.Excel.Application oXL;
            _Workbook oWB;
            _Worksheet oSheetProjectSummary = null;
            _Worksheet oSheetProjectSummaryFull = null;
            _Worksheet oSheetNCProgSummary = null;
            _Worksheet oSheetNewNCProgSummary = null;
            _Worksheet oSheetToolpathDetails = null;
            _Worksheet oSheetToolList = null;
            _Worksheet oSheetNewToolList = null;
            Dictionary<string, string> VarsListProjectSummary = new Dictionary<string, string>();
            Dictionary<string, string> VarsListProjectSummaryFull = new Dictionary<string, string>();
            Dictionary<string, string> VarsListNCProgSummary = new Dictionary<string, string>();
            Dictionary<string, string> VarsListToolpathDetails = new Dictionary<string, string>();
            Dictionary<string, string> VarsListToolList = new Dictionary<string, string>();
            bool Has_ProjectSummary = false;
            bool Has_ProjectSummaryFull = false;
            bool Has_NCProgSummary = false;
            bool Has_ToolList = false;
            bool Has_ToolpathDetails = false;
            string Project_Path = PowerMILLAutomation.ExecuteEx("print $project_pathname(0)");
            string NCProgDetails = "";
            string MergedModelList = "";
            int FileQty = 0;

            if (listNCProgsSelected.Items.Count == 0)
            {
                //No NC Program selected, do nothing...
                MessageBox.Show("Please select at least one NC Program");

            } else {

                List<ToolInfo> ToolData = new List<ToolInfo>();
                List<ToolpathInfo> ToolpathData = new List<ToolpathInfo>();
                List<ToolInfo> ProjectToolData = new List<ToolInfo>();
                List<ToolpathInfo> ProjectToolpathData = new List<ToolpathInfo>();
                ProjectInfo Project = new ProjectInfo();

                //Start Excel and get Application object.

                oXL = new Microsoft.Office.Interop.Excel.Application();
                oXL.Visible = true;
                oXL.DisplayAlerts = false;

                //Get the template from the option page
                string Path = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\ExcellSetupSheet\\Template.ini";
                string TemplateFile = "";
                if (File.Exists(Path))
                {
                    const Int32 BufferSize = 128;
                    using (var fileStream = File.OpenRead(Path))
                    using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
                    {
                        String line;
                        while ((line = streamReader.ReadLine()) != null)
                        {
                            TemplateFile = line;
                        }
                    }
                } else
                {
                    Path = System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Substring(8, System.Reflection.Assembly.GetExecutingAssembly().CodeBase.Length - 8);
                    string directory = System.IO.Path.GetDirectoryName(Path);
                    TemplateFile = directory + "\\Tool_List.xlsx";
                }

                if (!Directory.Exists(Project_Path + "\\Excel_Setupsheet\\"))
                {
                    Directory.CreateDirectory(Project_Path + "\\Excel_Setupsheet\\");
                }

                string[] fileArrayXLSX = Directory.GetFiles(Project_Path + "\\Excel_Setupsheet\\", "*.xlsx");
                string[] fileArrayXLS = Directory.GetFiles(Project_Path + "\\Excel_Setupsheet\\", "*.xls");
                foreach (string File in fileArrayXLSX)
                {
                    if (File.IndexOf("~$") < 0 )
                    {
                        FileQty = FileQty + 1;
                    }
                }

                foreach (string File in fileArrayXLS)
                {
                    if (File.IndexOf("~$") < 0)
                    {
                        FileQty = FileQty + 1;
                    }
                }

                if (FileQty > 0)
                {
                    File.Copy(TemplateFile, Project_Path + "\\Excel_Setupsheet\\SetupSheet_" + FileQty + ".xlsx");
                    TemplateFile = Project_Path + "\\Excel_Setupsheet\\SetupSheet_" + FileQty + ".xlsx";
                }
                else
                {
                    File.Copy(TemplateFile, Project_Path + "\\Excel_Setupsheet\\SetupSheet.xlsx");
                    TemplateFile = Project_Path + "\\Excel_Setupsheet\\SetupSheet.xlsx";
                }


                //Get a new workbook.
                oWB = (_Workbook)(oXL.Workbooks.Open(TemplateFile));

                List<string> ModelList = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.MachinableModels);
                PowerMILLAutomation.GetModelsLimits(ModelList, out double ModelsMinX, out double ModelsMinY, out double ModelsMinZ, out double ModelsMaxX, out double ModelsMaxY, out double ModelsMaxZ);

                Project = new ProjectInfo();
                Project.MachModelsMaxX = ModelsMaxX.ToString();
                Project.MachModelsMaxY = ModelsMaxY.ToString();
                Project.MachModelsMaxZ = ModelsMaxZ.ToString();
                Project.MachModelsMinX = ModelsMinX.ToString();
                Project.MachModelsMinY = ModelsMinY.ToString();
                Project.MachModelsMinZ = ModelsMinZ.ToString();
                Project.Name = PowerMILLAutomation.ExecuteEx("print $project_pathname(1)");
                Project.Path = PowerMILLAutomation.ExecuteEx("print $project_pathname(0)");
                Project.OrderNumber = PowerMILLAutomation.ExecuteEx("print $project.orderNumber");
                Project.Programmer = PowerMILLAutomation.ExecuteEx("print $project.programmer");
                Project.PartName = PowerMILLAutomation.ExecuteEx("print $project.partname");
                Project.Customer = PowerMILLAutomation.ExecuteEx("print $project.customer");
                Project.Date = DateTime.Now.ToString();
                Project.Notes = PowerMILLAutomation.ExecuteEx("print $project.notes");
                Project.ExcelTemplate = TemplateFile;
                foreach (string Model in ModelList)
                {
                    MergedModelList = MergedModelList + Environment.NewLine + Model;
                }
                Project.ModelsList = MergedModelList;
                Project.CombinedNCTPList = "";

                foreach (Worksheet worksheet in oWB.Worksheets)
                {
                    if (worksheet.Name == "Project_Summary")
                    {
                        oSheetProjectSummary = worksheet;
                        //Extract template keywords and cell adresses
                        VarsListProjectSummary = WriteFiles.ExtractTemplateData(oSheetProjectSummary);
                        Has_ProjectSummary = true;
                    }
                    else if (worksheet.Name == "Project_Summary_Full")
                    {
                        oSheetProjectSummaryFull = worksheet;
                        //Extract template keywords and cell adresses
                        VarsListProjectSummaryFull = WriteFiles.ExtractTemplateData(oSheetProjectSummaryFull);
                        Has_ProjectSummaryFull = true;
                    }
                    else if (worksheet.Name == "NCProg_Summary")
                    {
                        oSheetNCProgSummary = worksheet;
                        //Extract template keywords and cell adresses
                        VarsListNCProgSummary = WriteFiles.ExtractTemplateData(oSheetNCProgSummary);
                        Has_NCProgSummary = true;
                    }
                    else if (worksheet.Name == "Toolpath_Details")
                    {
                        oSheetToolpathDetails = worksheet;
                        //Extract template keywords and cell adresses
                        VarsListToolpathDetails = WriteFiles.ExtractTemplateData(oSheetToolpathDetails);
                        Has_ToolpathDetails = true;
                    }
                    else if (worksheet.Name == "ToolList")
                    {
                        oSheetToolList = worksheet;
                        //Extract template keywords and cell adresses
                        VarsListToolList = WriteFiles.ExtractTemplateData(oSheetToolList);
                        Has_ToolList = true;
                    }
                }


                List<string> NCProg_Toolpaths = new List<string>();
                foreach (String NCProg in listNCProgsSelected.Items)
                {
                    if (NCProgDetails == "")
                    {
                        NCProgDetails = NCProg;
                    }
                    else
                    {
                        NCProgDetails = NCProgDetails + Environment.NewLine + NCProg;
                    }
                    NCProg_Toolpaths = PowerMILLAutomation.GetNCProgToolpathes(NCProg);
                    foreach (string Toolpath in NCProg_Toolpaths)
                    {
                        NCProgDetails = NCProgDetails + Environment.NewLine + "    " + Toolpath;
                    }
                    Project.TotalTime = Project.TotalTime + double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';'" + NCProg + "').Statistics.TotalTime"));

                    //Extract the PowerMILL parameters found in the template from the current NCProgram toolapths
                    WriteFiles.ExtractData(NCProg, out ToolData, out ToolpathData);

                    if (Has_ProjectSummaryFull)
                    {
                        foreach (ToolpathInfo Toolpath in ToolpathData)
                        {
                            ProjectToolpathData.Add(new ToolpathInfo
                            {
                                Name = Toolpath.Name,
                                Description = Toolpath.Description,
                                Notes = Toolpath.Notes,
                                ToolName = Toolpath.ToolName,
                                ToolNumber = Toolpath.ToolNumber,
                                ToolDiameter = Toolpath.ToolDiameter,
                                ToolType = Toolpath.ToolType,
                                ToolCutterLength = Toolpath.ToolCutterLength,
                                ToolHolderName = Toolpath.ToolHolderName,
                                ToolOverhang = Toolpath.ToolOverhang,
                                ToolNumberOfFlutes = Toolpath.ToolNumberOfFlutes,
                                ToolLengthOffset = Toolpath.ToolLengthOffset,
                                ToolRadOffset = Toolpath.ToolRadOffset,
                                ToolpathType = Toolpath.ToolpathType,
                                ToolTipRadius = Toolpath.ToolTipRadius,
                                ToolDescription = Toolpath.ToolDescription,
                                Thickness = Toolpath.Thickness,
                                AxialThickness = Toolpath.AxialThickness,
                                CutterComp = Toolpath.CutterComp,
                                Feed = Toolpath.Feed,
                                Speed = Toolpath.Speed,
                                IPT = Toolpath.IPT,
                                SFM = Toolpath.SFM,
                                PlungeFeed = Toolpath.PlungeFeed,
                                SkimFeed = Toolpath.SkimFeed,
                                Coolant = Toolpath.Coolant,
                                Stepover = Toolpath.Stepover,
                                DOC = Toolpath.DOC,
                                GeneralAxisType = Toolpath.GeneralAxisType,
                                Statistic_Time = Toolpath.Statistic_Time,
                                TPWorkplane = Toolpath.TPWorkplane,
                                Tolerance = Toolpath.Tolerance,
                                RapidHeight = Toolpath.RapidHeight,
                                SkimHeight = Toolpath.SkimFeed,
                                ToolpathMinX = Toolpath.ToolpathMinX,
                                ToolpathMinY = Toolpath.ToolpathMinY,
                                ToolpathMinZ = Toolpath.ToolpathMinZ,
                                ToolpathMaxX = Toolpath.ToolpathMaxX,
                                ToolpathMaxY = Toolpath.ToolpathMaxY,
                                ToolpathMaxZ = Toolpath.ToolpathMaxZ,
                                FirstLeadInType = Toolpath.FirstLeadInType,
                                SecondLeadInType = Toolpath.SecondLeadInType,
                                FirstLeadOutType = Toolpath.FirstLeadOutType,
                                SecondLeadOutType = Toolpath.SecondLeadOutType,
                                CuttingDistance = Toolpath.CuttingDistance,
                                NCProgName = NCProg
                            });
                        }
                    }

                    if (Has_NCProgSummary)
                    {
                        //Write the Excel document
                        int Index = oSheetNCProgSummary.Index;
                        oSheetNCProgSummary.Copy(oSheetNCProgSummary, Type.Missing);

                        if (NCProg.Length > 31)
                        {
                            oWB.Sheets[Index].Name = NCProg.Replace("*","").Substring(0, 30);
                        }
                        else
                        {
                            oWB.Sheets[Index].Name = NCProg;
                        }
                        oSheetNewNCProgSummary = oWB.Sheets[Index];

                        if (Has_ToolpathDetails)
                        {
                            WriteFiles.CreateExcelFile(NCProg, ToolpathData, ToolData, Project, oSheetNewNCProgSummary, VarsListNCProgSummary, oWB, true, Project_Path, oSheetNewNCProgSummary, NCProgDetails);
                        }
                        else
                        {
                            WriteFiles.CreateExcelFile(NCProg, ToolpathData, ToolData, Project, oSheetNewNCProgSummary, VarsListNCProgSummary, oWB, false, Project_Path, oSheetNewNCProgSummary, NCProgDetails);
                        }

                    }
                    if (Has_ToolpathDetails)
                    {
                        //Write the Excel document
                        WriteFiles.CreateExcelFile(NCProg, ToolpathData, ToolData, Project, oSheetToolpathDetails, VarsListToolpathDetails, oWB, false, Project_Path, oSheetNewNCProgSummary, NCProgDetails);
                    }
                    if (Has_ToolList)
                    {
                        //Write the Excel document
                        int Index = oSheetToolList.Index;
                        oSheetToolList.Copy(oSheetToolList, Type.Missing);
                        if (NCProg.Length > 25)
                        {
                            oWB.Sheets[Index].Name = NCProg.Replace("*","").Substring(0, 24) + "-Tools";
                        }
                        else
                        {
                            oWB.Sheets[Index].Name = NCProg + "-Tools";
                        }
                        oWB.Sheets[Index].Name = NCProg + "-Tools";
                        oSheetNewToolList = oWB.Sheets[Index];

                        WriteFiles.CreateExcelFile(NCProg, ToolpathData, ToolData, Project, oSheetNewToolList, VarsListToolList, oWB, false, Project_Path, oSheetNewNCProgSummary, NCProgDetails);
                    }


                }
                Project.CombinedNCTPList = NCProgDetails;
                if (Has_ProjectSummary)
                {
                    WriteFiles.CreateExcelFile("None", ToolpathData, ToolData, Project, oSheetProjectSummary, VarsListProjectSummary, oWB, true, Project_Path, oSheetProjectSummary, NCProgDetails);
                }
                if (Has_ProjectSummaryFull)
                {
                    WriteFiles.CreateExcelFile("None", ProjectToolpathData, ToolData, Project, oSheetProjectSummaryFull, VarsListProjectSummaryFull, oWB, true, Project_Path, oSheetProjectSummaryFull, NCProgDetails);

                }

                if (Has_NCProgSummary)
                {
                    oSheetNCProgSummary.Delete();
                }
                if (Has_ToolpathDetails)
                {
                    oSheetToolpathDetails.Delete();
                }
                if (Has_ToolList)
                {
                    oSheetToolList.Delete();
                }
                oXL.DisplayAlerts = true;
                oWB.Sheets[1].Activate();
                oWB.Save();
                MessageBox.Show("SetupSheet exported successfully");
            }
        }

        public static FileInfo GetNewestFile(DirectoryInfo directory)
        {
            return directory.GetFiles()
                .Union(directory.GetDirectories().Select(d => GetNewestFile(d)))
                .OrderByDescending(f => (f == null ? DateTime.MinValue : f.LastWriteTime))
                .FirstOrDefault();
        }

        private void Get_Picture_Click(object sender, RoutedEventArgs e)
        {
           string New_PictureName = "";
           string Path = PowerMILLAutomation.ExecuteEx("print $project_pathname(0)") + "\\SetupSheets_files\\snapshots";

            if (listNCProgsSelected.Items.Count == 0)
            {
                MessageBox.Show("Please select at least 1 NC Program");
            }
            else
            {
                string NCProg = NCProg = listNCProgsSelected.Items[0].ToString();

                PowerMILLAutomation.ExecuteEx("KEEP SNAPSHOT NCPROGRAM '" + NCProg + "' CURRENT");

                FileInfo Picture = GetNewestFile(new DirectoryInfo(@Path));

                Path = PowerMILLAutomation.ExecuteEx("print $project_pathname(0)") + "\\Excel_Setupsheet";

                bool DirectoryExist = Directory.Exists(Path);

                if (DirectoryExist == false)
                {
                    Directory.CreateDirectory(Path);
                }

                Path = Path + "//";

                if (Toolpath_Pics.IsChecked == true)
                {
                    New_PictureName = Picture_List.Text;
                    ConvertToUsableName(New_PictureName, out New_PictureName);
                    New_PictureName = Path + "TP-" + New_PictureName + ".png";
                }
                else if (NCProg_Pics.IsChecked == true)
                {
                    New_PictureName = NCProg + "-" + Picture_List.Text;
                    ConvertToUsableName(New_PictureName, out New_PictureName);
                    New_PictureName = Path + "NCP-" + New_PictureName + ".png";
                }
                else if (Project_Pics.IsChecked == true)
                {
                    New_PictureName = Picture_List.Text;
                    ConvertToUsableName(New_PictureName, out New_PictureName);
                    New_PictureName = Path + "PRJ-" + New_PictureName + ".png";
                }


                if (File.Exists(New_PictureName))
                {
                    Setup_Picture.Source = null;
                    File.Delete(New_PictureName);
                }
                File.Move(Picture.FullName, New_PictureName);

                BitmapImage image = new BitmapImage();
                using (var stream = File.OpenRead(New_PictureName))
                {
                    image.BeginInit();
                    image.CacheOption = BitmapCacheOption.OnLoad;
                    image.StreamSource = stream;
                    image.EndInit();
                }
                Setup_Picture.Source = image;
            }


        }


        private void Del_Picture_Click(object sender, RoutedEventArgs e)
        {
            if (listNCProgsSelected.Items.Count == 0)
            {
                MessageBox.Show("Please select at least 1 NC Program");
            }
            else
            {
                string NCProg = listNCProgsSelected.Items[0].ToString();
                string PictureName = "";
                if (Toolpath_Pics.IsChecked == true)
                {
                    PictureName = Picture_List.Text;
                    ConvertToUsableName(PictureName, out PictureName);
                    PictureName = PowerMILLAutomation.ExecuteEx("print $project_pathname(0)") + "\\Excel_Setupsheet\\" + "TP-" + PictureName + ".png";
                }
                else if (NCProg_Pics.IsChecked == true)
                {
                    PictureName = NCProg + "-" + Picture_List.Text;
                    ConvertToUsableName(PictureName, out PictureName);
                    PictureName = PowerMILLAutomation.ExecuteEx("print $project_pathname(0)") + "\\Excel_Setupsheet\\" + "NCP-" + PictureName + ".png";
                }
                else if (Project_Pics.IsChecked == true)
                {
                    PictureName = Picture_List.Text;
                    ConvertToUsableName(PictureName, out PictureName);
                    PictureName = PowerMILLAutomation.ExecuteEx("print $project_pathname(0)") + "\\Excel_Setupsheet\\" + "PRJ-" + PictureName + ".png";
                }

                Setup_Picture.Source = null;
                File.Delete(PictureName);
            }
        }

         private void Picture_List_DropDownClosed(object sender, EventArgs e)
        {
            UpdatePicture();
        }

        private void Detailed_Checked(object sender, RoutedEventArgs e)
        {
            List<string> NCProg_Toolpaths = new List<string>();

            Picture_List.Items.Clear();
            if (listNCProgsSelected.Items.Count > 0)
            {

                foreach (String NCProg in listNCProgsSelected.Items)
                {
                    NCProg_Toolpaths = PowerMILLAutomation.GetNCProgToolpathes(NCProg);
                    for (int i = 0; i <= NCProg_Toolpaths.Count - 1; i++)
                    {
                        Picture_List.Items.Add(NCProg_Toolpaths[i]);
                    }
                }
            }
            Picture_List.SelectedIndex = 0;
            UpdatePicture();
        }

        private void Summary_Checked(object sender, RoutedEventArgs e)
        {
            Picture_List.Items.Clear();
            Picture_List.Items.Add("1");
            Picture_List.Items.Add("2");
            Picture_List.Items.Add("3");
            Picture_List.Items.Add("4");
            Picture_List.Items.Add("5");
            Picture_List.SelectedIndex = 0;
            UpdatePicture();
        }

        private void Picture_List_KeyUp(object sender, KeyEventArgs e)
        {
            Control ctl;
            ctl = (Control)sender;
            if (e.Key == Key.Down || e.Key == Key.Up)
            {
                UpdatePicture();
            }
        }

        private void UpdatePicture()
        {
            string PictureName = "";
            string Path = "";


            if (listNCProgsSelected.Items.Count > 0)
            {
                Path = PowerMILLAutomation.ExecuteEx("print $project_pathname(0)") + "\\Excel_Setupsheet\\";
                string NCProg = NCProg = listNCProgsSelected.Items[0].ToString();

                if (Toolpath_Pics.IsChecked == true)
                {
                    PictureName = Picture_List.Text;
                    ConvertToUsableName(PictureName, out PictureName);
                    PictureName = Path + "TP-" + PictureName + ".png";
                }
                else if (NCProg_Pics.IsChecked == true)
                {
                    PictureName = NCProg + "-" + Picture_List.Text;
                    ConvertToUsableName(PictureName, out PictureName);
                    PictureName = Path + "NCP-" + PictureName + ".png";
                }
                else if (Project_Pics.IsChecked == true)
                {
                    PictureName = Picture_List.Text;
                    ConvertToUsableName(PictureName, out PictureName);
                    PictureName = Path + "PRJ-" + PictureName + ".png";
                }

                if (File.Exists(PictureName))
                {
                    BitmapImage image = new BitmapImage();
                    using (var stream = File.OpenRead(PictureName))
                    {
                        image.BeginInit();
                        image.CacheOption = BitmapCacheOption.OnLoad;
                        image.StreamSource = stream;
                        image.EndInit();
                    }
                    Setup_Picture.Source = image;
                }
                else
                {
                    Setup_Picture.Source = null;
                }
            }
        }

        private void ConvertToUsableName(string OriginalName, out string NewName)
        {
            NewName = "";
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                NewName = OriginalName.Replace(c, '_');
            }
            NewName = NewName.Replace("*","");
        }
        private void Project_Pics_Checked(object sender, RoutedEventArgs e)
        {
            Picture_List.Items.Clear();
            Picture_List.Items.Add("1");
            Picture_List.Items.Add("2");
            Picture_List.Items.Add("3");
            Picture_List.Items.Add("4");
            Picture_List.Items.Add("5");
            Picture_List.SelectedIndex = 0;
            UpdatePicture();
        }
    }
}
