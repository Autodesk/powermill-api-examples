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
using System.Windows;
using System.Diagnostics;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Xml.Serialization;

using Delcam.Plugins.Framework;
using Delcam.Plugins.Events;

using PowerMILLExporter;

namespace PowerMILLVERICUTInterfacePlugin.Forms
{
    public partial class VericutPaneWPF : UserControl
    {
        #region plugin_data

        private IPluginCommunicationsInterface oComm;
        ProjectData proj_data = new ProjectData();
        PluginSettings plugin_settings = new PluginSettings();
        int active_setup = -1;
        bool set_setup_settings = false;

        #endregion

        #region main_plugin_code

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="comms"></param>
        public VericutPaneWPF(IPluginCommunicationsInterface comms)
        {
            EventLogger.InitializeEventLogger();
            Messages.PluginName = "PowerMILL to VERICUT";
            InitializeComponent();
            //
            oComm = comms;

            // create the pmill automation object to send pmill commands
            //pmill = new PowerMILLAutomation(oComm);
            PowerMILLAutomation.SetVariables(oComm);

            // translate plugin static text
            ButExport.Content = "Export";
            // load and display PM units
            SetPluginDefaults();
            LoadUnits();
            LoadPluginSettings();

            // subscribe to some events
            comms.EventUtils.Subscribe(new EventSubscription("EntityCreated", EntityCreated));
            comms.EventUtils.Subscribe(new EventSubscription("EntityDeleted", EntityDeleted));
            comms.EventUtils.Subscribe(new EventSubscription("ProjectClosed", ProjectClosed));
            comms.EventUtils.Subscribe(new EventSubscription("ProjectOpened", ProjectOpened));
            comms.EventUtils.Subscribe(new EventSubscription("EntityRenamed", EntityRenamed));
            comms.EventUtils.Subscribe(new EventSubscription("UnitsChanged", UnitsChanged));

            proj_data = new ProjectData();
            LoadProjectComponents();

            string ini_fpath = PowerMILLAutomation.GetParameterValueTerse("project_pathname(0)");
            if (!String.IsNullOrEmpty(ini_fpath) && !ini_fpath.Equals("None", StringComparison.OrdinalIgnoreCase))
            {
                ProjectData.ProjectPath = ini_fpath.Replace('/', System.IO.Path.DirectorySeparatorChar);
                ProjectData.ProjectName = System.IO.Path.GetFileNameWithoutExtension(ProjectData.ProjectPath);

                ProjectData.ProjectPluginDataPath = System.IO.Path.Combine(ini_fpath, @"plugin_data\VERICUTInterface\interface.xml");
                LoadProjectDataFromIni();
            }

        }

        private void ReadPowermillUCSs()
        {
            proj_data.sUCSs.Clear();
            proj_data.sUCSs.Add("Global");
            proj_data.sUCSs.AddRange(PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.Workplanes));

            List<string> selected_wp = new List<string>();
            foreach (string name in listWorkplanesSelected.Items)
            {
                selected_wp.Add(name);
            }

            cToUCS.Items.Clear();
            cbAttachWP.Items.Clear();
            cbModelWP.Items.Clear();
            cbNCWP.Items.Clear();
            listWorkplanes.Items.Clear();
            listWorkplanesSelected.Items.Clear();
            cbCutStockTransitionWP.Items.Clear();
            foreach (string ucs in proj_data.sUCSs)
            {
                cToUCS.Items.Add(ucs);
                cbAttachWP.Items.Add(ucs);
                if (selected_wp.Contains(ucs))
                {
                    listWorkplanesSelected.Items.Add(ucs);
                }
                else
                {
                    listWorkplanes.Items.Add(ucs);
                }
                cbModelWP.Items.Add(ucs);
                cbNCWP.Items.Add(ucs);
                cbCutStockTransitionWP.Items.Add(ucs);
            }
            cToUCS.SelectedIndex = 0;
            if (active_setup >= 0 && active_setup < proj_data.setup_infos.Count() &&
                cbAttachWP.Items.Contains(proj_data.setup_infos[active_setup].attach_workplane))
            {
                cbAttachWP.SelectedItem = proj_data.setup_infos[active_setup].attach_workplane;
            }
            else
            {
                cbAttachWP.SelectedIndex = 0;
            }
            cbCutStockTransitionWP.SelectedIndex = 0;
            if (!String.IsNullOrEmpty(proj_data.cut_stock_csys))
                if (cbCutStockTransitionWP.Items.Contains(proj_data.cut_stock_csys))
                    cbCutStockTransitionWP.SelectedItem = proj_data.cut_stock_csys;
        }

        private void SetPluginDefaults()
        {
            //create all 

            cbToolLibrary.SelectedIndex = 0;

            /* NC defaults */
            rExportNC.IsChecked = true;
            rBlockStockModel.IsChecked = true;
            rPartModels.IsChecked = true;
            rFixtureModels.IsChecked = true;
            rSelectWPs.IsChecked = true;
            cbToolLibrary.SelectedIndex = 0;
        }

        #endregion

        #region "PowerMILLEvents"

        /// <summary>
        /// Action of Project units changed event
        /// </summary>
        /// <param name="event_name"></param>
        /// <param name="event_arguments"></param>
        void UnitsChanged(string event_name, Dictionary<string, string> event_arguments)
        {
            if (ProjectData.PmillUnitsIsInch)
            {
                tModelExportTol.Text = "0.004";
                tStockExportTol.Text = "0.004";
                tFixtureExportTol.Text = "0.004";
                tToolExportTol.Text = "0.004";
            }
            else
            {
                tModelExportTol.Text = "0.1";
                tStockExportTol.Text = "0.1";
                tFixtureExportTol.Text = "0.1";
                tToolExportTol.Text = "0.1";
            }
        }

        /// <summary>
        /// Action of Entity Created event
        /// </summary>
        /// <param name="event_name"></param>
        /// <param name="event_arguments"></param>
        void EntityCreated(string event_name, Dictionary<string, string> event_arguments)
        {
            string entity_type = event_arguments["EntityType"];
            string entity_name = event_arguments["Name"];

            switch (entity_type)
            {
                case "Model":
                    EventLogger.WriteToEvengLog(String.Format("Model {0} was created. Update interface.", entity_name));
                    listModels.Items.Add(entity_name);
                    listFixtures.Items.Add(entity_name);
                    break;

                case "Ncprogram":
                    EventLogger.WriteToEvengLog(String.Format("NC program {0} was created. Update interface.", entity_name));
                    listNCProgs.Items.Add(entity_name);
                    break;

                case "Stockmodel":
                    EventLogger.WriteToEvengLog(String.Format("Stock model {0} was created. Update interface.", entity_name));
                    listStockModels.Items.Add(entity_name);
                    break;

                case "Toolpath":
                    EventLogger.WriteToEvengLog(String.Format("Toolpath {0} was created. Update interface.", entity_name));
                    listBlocks.Items.Add(entity_name);
                    break;

                case "Workplane":
                    EventLogger.WriteToEvengLog(String.Format("Workplane {0} was created. Update interface.", entity_name));
                    ReadPowermillUCSs();
                    break;

            }
        }

        /// <summary>
        /// Action of Entity Created event
        /// </summary>
        /// <param name="event_name"></param>
        /// <param name="event_arguments"></param>
        void EntityRenamed(string event_name, Dictionary<string, string> event_arguments)
        {
            string entity_type = event_arguments["EntityType"];
            string entity_orig_name = event_arguments["PreviousName"];
            string entity_new_name = event_arguments["Name"];

            switch (entity_type)
            {
                case "Model":
                    EventLogger.WriteToEvengLog(String.Format("Model {0} was renamed to {1}. Update interface.", entity_orig_name, entity_new_name));
                    if (listModels.Items.Contains(entity_orig_name))
                    {
                        listModels.Items.Remove(entity_orig_name);
                        listModels.Items.Add(entity_new_name);
                    }
                    if (listModelsSelected.Items.Contains(entity_orig_name))
                    {
                        listModelsSelected.Items.Remove(entity_orig_name);
                        listModelsSelected.Items.Add(entity_new_name);
                    }
                    if (listFixtures.Items.Contains(entity_orig_name))
                    {
                        listFixtures.Items.Remove(entity_orig_name);
                        listFixtures.Items.Add(entity_new_name);
                    }
                    if (listFixturesSelected.Items.Contains(entity_orig_name))
                    {
                        listFixturesSelected.Items.Remove(entity_orig_name);
                        listFixturesSelected.Items.Add(entity_new_name);
                    }
                    break;

                case "Ncprogram":
                    EventLogger.WriteToEvengLog(String.Format("NC program {0} was renamed to {1}. Update interface.", entity_orig_name, entity_new_name));
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
                    EventLogger.WriteToEvengLog(String.Format("Stock model {0} was renamed to {1}. Update interface.", entity_orig_name, entity_new_name));
                    if (listStockModels.Items.Contains(entity_orig_name))
                    {
                        listStockModels.Items.Remove(entity_orig_name);
                        listStockModels.Items.Add(entity_new_name);
                    }
                    if (listStockModelsUsed.Items.Contains(entity_orig_name))
                    {
                        listStockModelsUsed.Items.Remove(entity_orig_name);
                        listStockModelsUsed.Items.Add(entity_new_name);
                    }
                    break;

                case "Toolpath":
                    EventLogger.WriteToEvengLog(String.Format("Toolpath {0} was renamed to {1}. Update interface.", entity_orig_name, entity_new_name));
                    if (listBlocks.Items.Contains(entity_orig_name))
                    {
                        listBlocks.Items.Remove(entity_orig_name);
                        listBlocks.Items.Add(entity_new_name);
                    }
                    if (listBlocksUsed.Items.Contains(entity_orig_name))
                    {
                        listBlocksUsed.Items.Remove(entity_orig_name);
                        listBlocksUsed.Items.Add(entity_new_name);
                    }
                    break;

                case "Workplane":
                    EventLogger.WriteToEvengLog(String.Format("Toolpath {0} was renamed to {1}. Update interface.", entity_orig_name, entity_new_name));
                    if (listWorkplanesSelected.Items.Contains(entity_orig_name))
                    {
                        listWorkplanesSelected.Items.Remove(entity_orig_name);
                        listWorkplanesSelected.Items.Add(entity_new_name);
                    }
                    ReadPowermillUCSs();
                    break;
            }
        }

        /// <summary>
        /// Action of Project Opened event
        /// </summary>
        /// <param name="event_name"></param>
        /// <param name="event_arguments"></param>
        void ProjectOpened(string event_name, Dictionary<string, string> event_arguments)
        {
            EventLogger.WriteToEvengLog(String.Format("Project {0} opened.", event_arguments["Path"]));

            ClearItems();

            proj_data = new ProjectData();
            SetPluginDefaults();
            LoadProjectComponents();
            LoadProjectDataFromIni();
        }

        private void ClearItems()
        {
            proj_data.Clear();
            listModels.Items.Clear();
            listFixtures.Items.Clear();
            listNCProgs.Items.Clear();
            listNCProgsSelected.Items.Clear();
            listBlocks.Items.Clear();
            listBlocksUsed.Items.Clear();
            listStockModels.Items.Clear();
            listStockModelsUsed.Items.Clear();
            cbAttachWP.Items.Clear();
            cbModelWP.Items.Clear();
            cbNCWP.Items.Clear();
            cProgName.Items.Clear();
            cToUCS.Items.Clear();
            listWorkplanes.Items.Clear();
            listWorkplanesSelected.Items.Clear();
            cbCutStockTransitionWP.Items.Clear();
            cbSetups.Items.Clear();
        }

        private void LoadProjectComponents()
        {
            //EventLogger.WriteToEvengLog("Get all models and populate UI");
            List<string> models = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.Models);
            foreach (string model in models)
            {
                listModels.Items.Add(model);
                listFixtures.Items.Add(model);
            }


            //EventLogger.WriteToEvengLog("Get all nc programs and populate UI");
            List<string> ncprogs = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.NCPrograms);
            foreach (string ncprog in ncprogs)
                listNCProgs.Items.Add(ncprog);

            //EventLogger.WriteToEvengLog("Get all nc programs and populate UI");
            ReadPowermillUCSs();

            cProgName.Items.Add("Work Offsets");
            cProgName.Items.Add("Program Zero");
            cProgName.SelectedIndex = 0;

            tRegister.Text = "54";

            List<string> stock_models = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.StockModels);
            EventLogger.WriteToEvengLog("Load list of stock models");
            foreach (string name in stock_models)
                listStockModels.Items.Add(name);

            List<string> toolpaths = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.Toolpaths);
            EventLogger.WriteToEvengLog("Load list of blocks/toolpaths");
            foreach (string name in toolpaths)
                listBlocks.Items.Add(name);

        }

        /// <summary>
        /// Action of Entity Deleted event
        /// </summary>
        /// <param name="event_name"></param>
        /// <param name="event_arguments"></param>
        void EntityDeleted(string event_name, Dictionary<string, string> event_arguments)
        {
            string entity_type = event_arguments["EntityType"];
            string entity_name = event_arguments["Name"];

            switch (entity_type)
            {
                case "Model":
                    EventLogger.WriteToEvengLog(String.Format("Model {0} was deleted. Update interface.", entity_name));
                    // models
                    if (listModels.Items.Contains(entity_name))
                        listModels.Items.Remove(entity_name);
                    else if (listModelsSelected.Items.Contains(entity_name))
                        listModelsSelected.Items.Remove(entity_name);
                    // clamp
                    if (listFixtures.Items.Contains(entity_name))
                        listFixtures.Items.Remove(entity_name);
                    else if (listFixturesSelected.Items.Contains(entity_name))
                        listFixturesSelected.Items.Remove(entity_name);
                    break;

                case "Ncprogram":
                    EventLogger.WriteToEvengLog(String.Format("NC program {0} was deleted. Update interface.", entity_name));
                    if (listNCProgs.Items.Contains(entity_name))
                        listNCProgs.Items.Remove(entity_name);
                    else if (listNCProgsSelected.Items.Contains(entity_name))
                        listNCProgsSelected.Items.Remove(entity_name);
                    break;

                case "Stockmodel":
                    EventLogger.WriteToEvengLog(String.Format("Stock model {0} was deleted. Update interface.", entity_name));
                    if (listStockModels.Items.Contains(entity_name))
                        listStockModels.Items.Remove(entity_name);
                    if (listStockModelsUsed.Items.Contains(entity_name))
                        listStockModelsUsed.Items.Remove(entity_name);
                    break;

                case "Toolpath":
                    EventLogger.WriteToEvengLog(String.Format("Toolpath {0} was deleted. Update interface.", entity_name));
                    if (listBlocks.Items.Contains(entity_name))
                        listBlocks.Items.Remove(entity_name);
                    if (listBlocksUsed.Items.Contains(entity_name))
                        listBlocksUsed.Items.Remove(entity_name);
                    break;

                case "Workplane":
                    EventLogger.WriteToEvengLog(String.Format("Toolpath {0} was deleted. Update interface.", entity_name));
                    ReadPowermillUCSs();
                    if (cbAttachWP.Items.Contains(entity_name))
                        cbAttachWP.Items.Remove(entity_name);
                    if (cbModelWP.Items.Contains(entity_name))
                        cbModelWP.Items.Remove(entity_name);
                    if (cbNCWP.Items.Contains(entity_name))
                        cbNCWP.Items.Remove(entity_name);
                    break;

            }
        }

        /// <summary>
        /// Action of Project Closed event
        /// </summary>
        void ProjectClosed(string event_name, Dictionary<string, string> event_arguments)
        {
            EventLogger.WriteToEvengLog("Project closed");
            tOutputDir.Text = "";
            tProjTemplate.Text = "";
            tNumSetups.Text = "";
            active_setup = -1;
            SavePluginSettings();
            ProjectData.ProjectPluginDataPath = "";
            ClearItems();
        }

        #endregion

        #region "FormHandling*

        /// <summary>
        /// When the table workplane is selected
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CboWorkplane_DropDownClosed(object sender, EventArgs e)
        {
            //if (CboWorkplane.Text.Trim() == "")
            //{
            //    PowerMILLAutomation.Execute("DEACTIVATE WORKPLANE");
            //}
            //else
            //{
            //    PowerMILLAutomation.Execute("ACTIVATE Workplane '" + CboWorkplane.Text + "'");
            //}
        }

        /// <summary>
        /// initialisations of the form 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void UserControl_Loaded(object sender, RoutedEventArgs e)
        {
            //ButOnOff.IsChecked = false;
            //ButOnOff_CheckedChanged(sender, e);
        }

        /// <summary>
        /// load and display units
        /// </summary>
        private void LoadUnits()
        {
            ProjectData.PmillUnitsIsInch = ((PowerMILLAutomation.Units == PowerMILLAutomation.enumUnit.mm) ? false : true);
            //lUnits.Content = "Units" + ": " + PowerMILLAutomation.Units.ToString().ToUpper();
        }

        /// <summary>
        /// Get output directory folder
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnFileExportPath_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog oFolderDialog = new System.Windows.Forms.FolderBrowserDialog();
            //if (oFolderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            //    txtOutputDir.Text = oFolderDialog.SelectedPath;
        }

        private void btnSaveSettings_Click(object sender, RoutedEventArgs e)
        {
            VerifyProjectData(false);
            SavePluginSettings();
            if (String.IsNullOrEmpty(ProjectData.ProjectPluginDataPath))
            {
                string ini_fpath = PowerMILLAutomation.ExecuteEx("PRINT VALUE PROJECTPATH").Replace('/', System.IO.Path.DirectorySeparatorChar);
                if (!String.IsNullOrEmpty(ini_fpath) && !ini_fpath.Equals("None", StringComparison.OrdinalIgnoreCase))
                    ProjectData.ProjectPluginDataPath = System.IO.Path.Combine(ini_fpath, @"plugin_data\VERICUTInterface\interface.xml");
            }
            if (!String.IsNullOrEmpty(ProjectData.ProjectPluginDataPath))
            {
                if (proj_data.setup_infos.Count == 0)
                    SetSetupsPlaceHolders();

                proj_data.output_dir = tOutputDir.Text;
                proj_data.proj_template = tProjTemplate.Text;
                if (cbSetups.SelectedIndex >= 0)
                    proj_data.setup_infos[cbSetups.SelectedIndex] = SetSetupSettings();

                SaveProjectSettings();
            }
        }

        /// <summary>
        /// Send data to VERICUT
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void butStart_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (!VerifyProjectData(true))
                {
                    return;
                }

                ButExport.IsEnabled = false;

                //PowerMILLAutomation.ShowMessageInProgressBar("Start export");
                bool bIsSTLMM = PowerMILLAutomation.IsSTLOutputInMM();
                string sShowTPDialog = PowerMILLAutomation.ExecuteEx("print par terse 'Options.ToolpathOpts.AutoRaise'");
                if (sShowTPDialog != "0")
                {
                    PowerMILLAutomation.ExecuteEx(String.Format("$Options.ToolpathOpts.AutoRaise={0}", 0));
                    EventLogger.WriteToEvengLog("Checkbox Raise dialog is checked in Options -> Toolpaths -> Activation dialog. Uncheck it");
                }

                if (ProjectData.PmillUnitsIsInch != !bIsSTLMM)
                    PowerMILLAutomation.SetSTLOutputUnits(!ProjectData.PmillUnitsIsInch);

                if (proj_data.setup_infos.Count == 0)
                    SetSetupsPlaceHolders();
                if (proj_data.setup_infos.Count == 0)
                {
                    Messages.ShowError(Properties.Resources.IDS_NoSetups);
                    return;
                }

                proj_data.output_dir = tOutputDir.Text;
                proj_data.proj_template = tProjTemplate.Text;
                if (0 <= cbSetups.SelectedIndex && cbSetups.SelectedIndex < cbSetups.Items.Count)
                {
                    proj_data.setup_infos[cbSetups.SelectedIndex] = SetSetupSettings();
                }

                EventLogger.WriteToEvengLog("Check form data");
                if (!CheckFormData()) return;

                EventLogger.WriteToEvengLog("Save interface settings");
                SavePluginSettings();
                SaveProjectSettings();

                // Check if output directory is empty or not
                if (String.IsNullOrWhiteSpace(tOutputDir.Text))
                {
                    Messages.ShowError(Properties.Resources.IDS_NoOutputDirectory);
                    return;
                }

                // Try to create the output directory. If it already exists this does nothing.
                try
                {
                    Directory.CreateDirectory(tOutputDir.Text);
                }
                catch
                {
                    Messages.ShowError(Properties.Resources.IDS_FolderNotCreated, tOutputDir.Text);
                    return;
                }

                // If the output directory is not empty, check whether or not to continue.
                if (Directory.EnumerateFiles(tOutputDir.Text, "*", SearchOption.AllDirectories).Any())
                {
                    if (Messages.ShowWarning(Properties.Resources.IDS_DirNotEmpty) == MessageBoxResult.No)
                    {
                        return;
                    }
                }

                InitializeExportSettings();

                if (String.IsNullOrWhiteSpace(proj_data.proj_template))
                {
                    Messages.ShowError(Properties.Resources.IDS_NoVericutTemplate);
                    return;
                }
                string project_fpath = ExportData.CreateOutputFiles(proj_data, proj_data.output_dir, plugin_settings);
                if (project_fpath == "*") return;
                if (string.IsNullOrEmpty(project_fpath))
                {
                    Messages.ShowError(Properties.Resources.IDS_ExportErrors);
                    return;
                }
                else
                {
                    if (Convert.ToBoolean(cbOpenVericut.IsChecked))
                        StartVericut(project_fpath);
                }

                Messages.ShowMessage(Properties.Resources.IDS_SuccessfullExport);
            }
            catch (Exception Ex)
            {
                EventLogger.WriteToEvengLog(Ex.Message + Environment.NewLine + Ex.StackTrace);
                Messages.ShowError(Properties.Resources.IDS_Exception, Ex.Message);
            }
            finally
            {
                ButExport.IsEnabled = true;
            }
        }

        private void StartVericut(string project_fpath)
        {
            try
            {
                if (!File.Exists(plugin_settings.vericut_fpath))
                {
                    EventLogger.WriteToEvengLog(
                        String.Format("Vericut path {0} is invalid. The file doesn't exist.", plugin_settings.vericut_fpath));
                    Messages.ShowError(
                        String.Format("Vericut path {0} is invalid. The file doesn't exist.", plugin_settings.vericut_fpath));
                    return;
                }
                if (!File.Exists(project_fpath))
                {
                    EventLogger.WriteToEvengLog(
                        String.Format("Vericut project {0} doesn't exist. Can't open it in VERICUT.", project_fpath));
                    Messages.ShowError(
                        String.Format("Vericut project {0} doesn't exist. Can't open it in VERICUT.", project_fpath));
                    return;
                }

                using (Process proc = new Process())
                {
                    proc.StartInfo.FileName = plugin_settings.vericut_fpath;
                    proc.StartInfo.Arguments = String.Format("\"{0}\"", project_fpath);
                    proc.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;
                    proc.Start();
                    proc.WaitForExit(0);
                    if (proc != null)
                        proc.Dispose();
                }
            }
            catch (Exception Ex)
            {
                EventLogger.WriteToEvengLog(Ex.Message);
            }
        }

        private void InitializeExportSettings()
        {
            if (plugin_settings.export_options == null)
                plugin_settings.export_options = new ExportOptions();
            plugin_settings.export_options.is_inch = ProjectData.PmillUnitsIsInch;
            plugin_settings.export_options.block_tol = Convert.ToDouble(tStockExportTol.Text);
            plugin_settings.export_options.model_tol = Convert.ToDouble(tModelExportTol.Text);
            plugin_settings.export_options.fixture_tol = Convert.ToDouble(tFixtureExportTol.Text);
            plugin_settings.export_options.tool_tol = Convert.ToDouble(tToolExportTol.Text);
            plugin_settings.vericut_fpath = tVericutPath.Text;
            plugin_settings.start_vericut = Convert.ToBoolean(cbOpenVericut.IsChecked);
            plugin_settings.tool_id_use_name = Convert.ToBoolean(cbToolName.IsChecked);
            plugin_settings.tool_id_use_num = Convert.ToBoolean(cbToolNumber.IsChecked);
        }

        /// <summary>
        /// Check that all required form fields were set
        /// </summary>
        /// <returns></returns>
        private bool CheckFormData()
        {
            foreach (SetupInfo setup_info in proj_data.setup_infos)
            {

                PowerMILLAutomation.Execute("DIALOGS MESSAGE OFF");
                foreach (NCProgramInfo ncprog_info in setup_info.nc_progs)
                {
                    if (!String.IsNullOrWhiteSpace(ncprog_info.sName))
                    {
                        string sProgCN = PowerMILLAutomation.ExecuteEx("EDIT NCPROGRAM '" + ncprog_info.sName + "' LIST");

                        string[] sTabCNInfos = sProgCN.Split((char)13);
                        int iSlashIndex = sTabCNInfos[2].IndexOf('/');
                        if (iSlashIndex < 0) iSlashIndex = 0;
                        int iSpaceIndex = sTabCNInfos[2].LastIndexOf((char)32, iSlashIndex);

                        // Get the NCProg File Path
                        ncprog_info.sPath = sTabCNInfos[2].Remove(0, iSpaceIndex).Trim();
                    }
                }
                PowerMILLAutomation.Execute("DIALOGS MESSAGE ON");

                if (setup_info.nc_progs == null)
                {
                    Messages.ShowError(String.Format("Setup {0}: At least one nc program should be selected to export.", setup_info.name));
                    return false;
                }
                foreach (NCProgramInfo nc_prog in setup_info.nc_progs)
                {
                    if (String.IsNullOrEmpty(nc_prog.sName) && setup_info.tools_to_use != 1)
                    {
                        Messages.ShowError(String.Format("Setup {0}: An existing tool library has to be used, if selecting an existing nc program.", setup_info.name));
                        return false;
                    }
                }
                if (setup_info.tools_to_use == 1 || setup_info.tools_to_use == 2)
                {
                    if (String.IsNullOrEmpty(setup_info.tls_fpath))
                    {
                        Messages.ShowError(String.Format("Setup {0}: Tools library file isn't specified in the template or the form.", setup_info.name));
                        return false;
                    }
                    if (!File.Exists(setup_info.tls_fpath))
                    {
                        Messages.ShowError(String.Format("Setup {0}: Tools library file specified in the template or the form doesn't exist.", setup_info.name));
                        return false;
                    }
                }
            }
            return true;
        }

        #endregion

        #region plugin_helpers

        /// <summary>
        /// Method to verify consistensy of the project data read from saved project with
        /// settings PowerMill project.
        /// </summary>
        private bool VerifyProjectData(bool is_error)
        {
            string warning_message = "";

            // Initialize global PowerMill project info.
            InitProjectSettings();

            // Verify global plugin data.
            SetupInfo.VerifyWorkplane(ref proj_data.cut_stock_csys, ref warning_message, "");
            SetupInfo.VerifyDirectory(ref proj_data.output_dir, ref warning_message, "");
            SetupInfo.VerifyFile(ref proj_data.proj_template, ref warning_message, "");

            // Verify Vericut setup data.
            foreach (SetupInfo setup_info in proj_data.setup_infos)
            {
                if (!File.Exists(setup_info.template))
                {
                    string msg = "";
                    SetupInfo.VerifyFile(ref setup_info.template, ref msg, " ");
                    setup_info.template = proj_data.proj_template;
                }
                string str = setup_info.VerifySetupInfo(" ", is_error);
                if (!String.IsNullOrWhiteSpace(str))
                {
                    warning_message += "=====================================\n";
                    warning_message += String.Format(Properties.Resources.IDS_SetupName, setup_info.name) + "\n";
                    warning_message += str;
                }
            }

            if (!String.IsNullOrWhiteSpace(warning_message))
            {
                if (is_error)
                {
                    return MessageBox.Show(warning_message,
                        Messages.PluginName, MessageBoxButton.OKCancel, MessageBoxImage.Error) == MessageBoxResult.OK;
                }
                else
                {
                    MessageBox.Show(warning_message,
                        Messages.PluginName, MessageBoxButton.OK, MessageBoxImage.Warning);
                }

                return false;
            }

            return true;
        }

        private void InitProjectSettings()
        {
            string ini_fpath = PowerMILLAutomation.GetParameterValueTerse("project_pathname(0)");
            if (!String.IsNullOrEmpty(ini_fpath) && !ini_fpath.Equals("None", StringComparison.OrdinalIgnoreCase))
            {
                ProjectData.ProjectPath = ini_fpath.Replace('/', System.IO.Path.DirectorySeparatorChar);
                ProjectData.ProjectName = System.IO.Path.GetFileNameWithoutExtension(ProjectData.ProjectPath);
                ProjectData.ProjectPluginDataPath = System.IO.Path.Combine(ini_fpath, @"plugin_data\VERICUTInterface\interface.xml");
            }
        }

        #endregion

        private WorkPlaneOrigin GetOXYMatriceForNCProgram(string NCProgName, ref bool is_error)
        {
            is_error = false;
            string wp_name = "";

            try
            {
                if (NCProgName.Trim() == "") return null;


                /*
                 * wp = GetNCWorkPlane(NCProgName);
                 * if (wp == null)
                {
                    is_error = true;
                    return new WorkPlaneOrigin();
                }
                WorkPlaneOrigin wpO = GetWorkplaneOriginInRelationToActiveWorkplane(PowerMILLAutomation.oToken, wp.Name);
                Replace with:*/
                wp_name = PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.Name", NCProgName)).Trim();

                if (String.IsNullOrEmpty(wp_name))
                {
                    is_error = true;
                }
                WorkPlaneOrigin wpO = MainModule.GetWorkplaneCoords(PowerMILLAutomation.oToken, wp_name, true, null);
                if (wpO == null)
                {
                    wpO = new WorkPlaneOrigin();
                }
                if (wpO == null)
                {
                    is_error = true;
                    //return new WorkPlaneOrigin(wp.Origin.X, wp.Origin.Y, wp.Origin.Z,
                    //                           wp.XAxis.X, wp.XAxis.Y, wp.XAxis.Z,
                    //                           wp.YAxis.X, wp.YAxis.Y, wp.YAxis.Z,
                    //                           wp.ZAxis.X, wp.ZAxis.Y, wp.ZAxis.Z,
                    //                           wp.XAngle, wp.YAngle, wp.ZAngle);
                    //Replace with:
                    return new WorkPlaneOrigin(
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.Origin[0]", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.Origin[1]", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.Origin[2]", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.XAxis[0]", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.XAxis[1]", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.XAxis[2]", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.YAxis[0]", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.YAxis[1]", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.YAxis[2]", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.ZAxis[0]", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.ZAxis[1]", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.ZAxis[2]", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.XAngle", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.YAngle", NCProgName)).Trim()),
                       Convert.ToDouble(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram', '{0}').OutputWorkplane.ZAngle", NCProgName)).Trim()));
                }
                return wpO;

            }
            catch (Exception ex)
            {
                Messages.ShowError(Properties.Resources.IDS_ExGetNCProgWorkplane, ex.Message);
                is_error = true;
                return null;
            }
            finally
            {
            }
        }

        private PowerMILL.Workplane GetWorkPlane(string workplane_name)
        {
            foreach (PowerMILL.Workplane oWorkplane in PowerMILLAutomation.oPServices.Project.Workplanes)
                if (oWorkplane.Name == workplane_name) return oWorkplane;

            return null;
        }

        private PowerMILL.Workplane GetNCWorkPlane(string ncprog_name)
        {
            foreach (PowerMILL.NCProgram oNCProg in PowerMILLAutomation.oPServices.Project.NCPrograms)
                if (oNCProg.Name == ncprog_name) return null;// oNCProg.OutputWorkplane;

            return null;
        }

        #region "IniFileRead/WriteData"

        private void SaveProjectSettings()
        {
            try
            {
                if (!Directory.Exists(System.IO.Path.GetDirectoryName(ProjectData.ProjectPluginDataPath)))
                    Directory.CreateDirectory(System.IO.Path.GetDirectoryName(ProjectData.ProjectPluginDataPath));
                System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(proj_data.GetType());
                System.Xml.XmlWriterSettings settings = new System.Xml.XmlWriterSettings();
                settings.NewLineChars = Environment.NewLine;
                settings.NewLineOnAttributes = true;
                settings.Indent = true;
                settings.NewLineHandling = System.Xml.NewLineHandling.Replace;
                settings.OmitXmlDeclaration = true;
                settings.CloseOutput = true;
                System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(ProjectData.ProjectPluginDataPath, settings);
                serializer.Serialize(writer, proj_data);
                writer.Close();
            }
            catch { }
        }

        private void SavePluginSettings()
        {
            try
            {
                InitializeExportSettings();
                string settings_fpath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "powermill_vericut_settings.xml");
                if (File.Exists(settings_fpath))
                    File.Delete(settings_fpath);

                System.Xml.Serialization.XmlSerializer serializer = new System.Xml.Serialization.XmlSerializer(plugin_settings.GetType());
                System.Xml.XmlWriterSettings settings = new System.Xml.XmlWriterSettings();
                settings.NewLineChars = Environment.NewLine;
                settings.NewLineOnAttributes = true;
                settings.Indent = true;
                settings.NewLineHandling = System.Xml.NewLineHandling.Replace;
                settings.OmitXmlDeclaration = true;
                settings.CloseOutput = true;
                System.Xml.XmlWriter writer = System.Xml.XmlWriter.Create(settings_fpath, settings);
                serializer.Serialize(writer, plugin_settings);
                writer.Close();
                tProjTemplate_TextChanged(null, null);
            }
            catch { }
        }

        private void LoadProjectDataFromIni()
        {
            if (String.IsNullOrEmpty(ProjectData.ProjectPluginDataPath))
                ProjectData.ProjectPluginDataPath = System.IO.Path.Combine(
                                                    PowerMILLAutomation.GetParameterValueTerse("project_pathname(0)").Replace('/', System.IO.Path.DirectorySeparatorChar),
                                                    @"plugin_data\VERICUTInterface\interface.xml");
            active_setup = -1;
            cbSetups.Items.Clear();

            if (!File.Exists(ProjectData.ProjectPluginDataPath)) return;
            TextReader r = new StreamReader(ProjectData.ProjectPluginDataPath);
            System.Xml.Serialization.XmlSerializer s = new System.Xml.Serialization.XmlSerializer(typeof(ProjectData));

            proj_data = (ProjectData)s.Deserialize(r);
            r.Close();

            VerifyProjectData(false);

            tOutputDir.Text = proj_data.output_dir;
            tProjTemplate.Text = proj_data.proj_template;

            if (proj_data.setup_infos != null)
            {
                cbSetups.IsEnabled = true;
                cbSetups.SelectedIndex = 0;
                tNumSetups.Text = proj_data.setup_infos.Count.ToString();
                SetUISettingsForSetup(0);
                set_setup_settings = true;
            }
            else
                tNumSetups.Text = "0";

            if (!String.IsNullOrWhiteSpace(proj_data.cut_stock_csys))
            {
                cBlockTransition.IsChecked = true;
            }
            else
            {
                cBlockTransition.IsChecked = false;
            }
            cBlockTransition_Click(null, null);
        }

        private void SetUISettingsForSetup(int index)
        {
            if (index < 0) return;
            if (proj_data.setup_infos.Count > 0)
            {
                tSetupName.Text = proj_data.setup_infos[index].name;
                UpdateUIAttachCompsAndSubsystems(proj_data.setup_infos[index].setup_attach_components, proj_data.setup_infos[index].setup_subsystems);

                /* Set nc program saved data */
                if (proj_data.setup_infos[index].nc_progs != null)
                    if (proj_data.setup_infos[index].nc_progs.Count > 0)
                    {
                        foreach (NCProgramInfo nc_prog in proj_data.setup_infos[index].nc_progs)
                        {
                            if (!String.IsNullOrEmpty(nc_prog.sName))
                            {
                                rExportNC.IsChecked = true;
                                listNCProgs.Items.Remove(nc_prog.sName);
                                if (!listNCProgsSelected.Items.Contains(nc_prog.sName))
                                    listNCProgsSelected.Items.Add(nc_prog.sName);
                            }
                            else if (!String.IsNullOrEmpty(nc_prog.sPath))
                            {
                                rUseNC.IsChecked = true;
                                if (!listNCProgsBrowserSelected.Items.Contains(nc_prog.sPath))
                                    listNCProgsBrowserSelected.Items.Add(nc_prog.sPath);
                            }
                        }
                    }
                /* Set work offsets */
                if (proj_data.setup_infos[index].offsets != null)
                {
                    if (proj_data.setup_infos[index].offsets.Count > 0)
                    {
                        dgOffsets.Items.Clear();
                        foreach (WorkOffset offset in proj_data.setup_infos[index].offsets)
                            dgOffsets.Items.Add(offset);
                        dgOffsets.SelectedIndex = 0;
                        DataGrid_SelectionChanged(null, null);
                    }
                }
                /* Set workplane saved data */
                if (proj_data.setup_infos[index].workplanes_to_export != null)
                    if (proj_data.setup_infos[index].workplanes_to_export.Count > 0)
                    {
                        foreach (string wp_name in proj_data.setup_infos[index].workplanes_to_export)
                        {
                            listWorkplanes.Items.Remove(wp_name);
                            if (!listWorkplanesSelected.Items.Contains(wp_name))
                                listWorkplanesSelected.Items.Add(wp_name);
                        }
                    }
                if (cbAttachWP.Items.Contains(proj_data.setup_infos[index].attach_workplane))
                {
                    cbAttachWP.SelectedItem = proj_data.setup_infos[index].attach_workplane;
                }
                else
                {
                    cbAttachWP.SelectedIndex = 0;
                }
                if (cbModelWP.Items.Contains(proj_data.setup_infos[index].model_workplane))
                {
                    cbModelWP.SelectedItem = proj_data.setup_infos[index].model_workplane;
                }
                else
                {
                    cbModelWP.SelectedIndex = 0;
                }
                if (cbNCWP.Items.Contains(proj_data.setup_infos[index].nc_workplane))
                {
                    cbNCWP.SelectedItem = proj_data.setup_infos[index].nc_workplane;
                }
                else
                {
                    cbNCWP.SelectedIndex = 0;
                }

                if (cbAttachWPTo.Items.Contains(proj_data.setup_infos[index].attach_workplane_to))
                {
                    cbAttachWPTo.SelectedItem = proj_data.setup_infos[index].attach_workplane_to;
                }
                else
                {
                    cbAttachWPTo.SelectedIndex = 0;
                }
                /* If transforming stock from setup to setup and cut stock transition is selected, block and model tabs should be disabled.*/
                if (String.IsNullOrEmpty(proj_data.cut_stock_csys) || cbSetups.SelectedIndex == 0)
                {
                    eBlock.IsEnabled = true;
                    ePart.IsEnabled = true;
                    /* Set block saved data */
                    if (!String.IsNullOrEmpty(proj_data.setup_infos[index].block_toolpath))
                    {
                        listBlocksUsed.Items.Clear();
                        if (!listBlocksUsed.Items.Contains(proj_data.setup_infos[index].block_toolpath))
                            listBlocksUsed.Items.Add(proj_data.setup_infos[index].block_toolpath);
                        if (listBlocks.Items.Contains(proj_data.setup_infos[index].block_toolpath))
                            listBlocks.Items.Remove(proj_data.setup_infos[index].block_toolpath);
                        rBlockToolpath.IsChecked = true;
                    }
                    if (!String.IsNullOrEmpty(proj_data.setup_infos[index].block_stockmodel))
                    {
                        listStockModelsUsed.Items.Clear();
                        listStockModelsUsed.Items.Add(proj_data.setup_infos[index].block_stockmodel);
                        if (listStockModels.Items.Contains(proj_data.setup_infos[index].block_stockmodel))
                            listStockModels.Items.Remove(proj_data.setup_infos[index].block_stockmodel);
                        rBlockStockModel.IsChecked = true;
                    }
                    if (!String.IsNullOrEmpty(proj_data.setup_infos[index].block_stl_fpath))
                    {
                        TxtBlockSTL.Items.Clear();
                        TxtBlockSTL.Items.Add(proj_data.setup_infos[index].block_stl_fpath);
                        rBlockStockFile.IsChecked = true;
                    }
                    if (cbAttachStockTo.Items.Contains(proj_data.setup_infos[index].block_attach_to))
                    {
                        cbAttachStockTo.SelectedItem = proj_data.setup_infos[index].block_attach_to;
                    }
                    else
                    {
                        cbAttachStockTo.SelectedIndex = 0;
                    }
                    cBlockTransition.IsChecked = !String.IsNullOrEmpty(proj_data.cut_stock_csys);
                    cbCutStockTransitionWP.SelectedIndex = 0;
                    if (!String.IsNullOrEmpty(proj_data.cut_stock_csys))
                        if (cbCutStockTransitionWP.Items.Contains(proj_data.cut_stock_csys))
                            cbCutStockTransitionWP.SelectedItem = proj_data.cut_stock_csys;
                    /* Set part/design saved data */
                    if (proj_data.setup_infos[index].part_models != null)
                    {
                        if (proj_data.setup_infos[index].part_models.Count > 0)
                        {
                            listModelsSelected.Items.Clear();
                            foreach (string part_model in proj_data.setup_infos[index].part_models)
                            {
                                listModelsSelected.Items.Add(part_model);
                                if (listModels.Items.Contains(part_model))
                                    listModels.Items.Remove(part_model);
                            }
                            rPartModels.IsChecked = true;
                        }
                    }
                    if (proj_data.setup_infos[index].part_stl_fpaths != null)
                    {
                        if (proj_data.setup_infos[index].part_stl_fpaths.Count > 0)
                        {
                            listModelsBrowserSelected.Items.Clear();
                            foreach (string part_fpath in proj_data.setup_infos[index].part_stl_fpaths)
                                listModelsBrowserSelected.Items.Add(part_fpath);
                            rPartSTLs.IsChecked = true;
                        }
                    }
                    if (cbAttachModelsTo.Items.Contains(proj_data.setup_infos[index].part_attach_to))
                    {
                        cbAttachModelsTo.SelectedItem = proj_data.setup_infos[index].part_attach_to;
                    }
                    else
                    {
                        cbAttachModelsTo.SelectedIndex = 0;
                    }
                }
                else
                {
                    //eBlock.IsEnabled = false;
                    //eBlock.IsExpanded = false;
                    //ePart.IsEnabled = false;
                    //ePart.IsExpanded = false;
                }
                /* Set fixture saved data */
                if (proj_data.setup_infos[index].fixture_models != null)
                {
                    if (proj_data.setup_infos[index].fixture_models.Count > 0)
                    {
                        listFixturesSelected.Items.Clear();
                        foreach (string fixture_model in proj_data.setup_infos[index].fixture_models)
                        {
                            listFixturesSelected.Items.Add(fixture_model);
                            if (listFixtures.Items.Contains(fixture_model))
                                listFixtures.Items.Remove(fixture_model);
                        }
                        rFixtureModels.IsChecked = true;
                    }
                }
                if (proj_data.setup_infos[index].fixture_stl_fpaths != null)
                {
                    if (proj_data.setup_infos[index].fixture_stl_fpaths.Count > 0)
                    {
                        listFixtureBrowserSelected.Items.Clear();
                        foreach (string fixture_fpath in proj_data.setup_infos[index].fixture_stl_fpaths)
                            listFixtureBrowserSelected.Items.Add(fixture_fpath);
                        rFixtureSTLs.IsChecked = true;
                    }
                }
                if (cbAttachModelsTo.Items.Contains(proj_data.setup_infos[index].fixture_attach_to))
                {
                    cbAttachFixtureTo.SelectedItem = proj_data.setup_infos[index].fixture_attach_to;
                }
                else
                {
                    cbAttachFixtureTo.SelectedIndex = 0;
                }
                if (proj_data.setup_infos[index].tools_to_use >= 0)
                {
                    cbToolLibrary.SelectedIndex = proj_data.setup_infos[index].tools_to_use;
                    if (proj_data.setup_infos[index].tools_from_template)
                        rToolLibFromTemplate.IsChecked = true;
                    else
                        rToolLibFromTemplate.IsChecked = false;
                }
            }
        }

        private void LoadPluginSettings()
        {
            string settings_fpath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "powermill_vericut_settings.xml");
            bool use_default_settings = false;

            if (!File.Exists(settings_fpath))
            {
                use_default_settings = true;
            }
            else
            {
                try
                {
                    TextReader r = new StreamReader(settings_fpath);
                    XmlSerializer s = new XmlSerializer(typeof(PluginSettings));

                    plugin_settings = (PluginSettings)s.Deserialize(r);
                    r.Close();
                }
                catch (InvalidOperationException e)
                {
                    string warning_message =
                        "Reverting to the default VERICUT plugin settings as the powermill_vericut_settings.xml file could not be read:";
                    warning_message += "\n" + e.Message;
                    MessageBox.Show(warning_message, Messages.PluginName, MessageBoxButton.OK, MessageBoxImage.Warning);
                    
                    use_default_settings = true;
                }
            }

            if (use_default_settings)
            { 
                if (ProjectData.PmillUnitsIsInch)
                {
                    tModelExportTol.Text = "0.004";
                    tStockExportTol.Text = "0.004";
                    tFixtureExportTol.Text = "0.004";
                    tToolExportTol.Text = "0.004";
                }
                else
                {
                    tModelExportTol.Text = "0.1";
                    tStockExportTol.Text = "0.1";
                    tFixtureExportTol.Text = "0.1";
                    tToolExportTol.Text = "0.1";
                }
            }
            else
            {
                tVericutPath.Text = plugin_settings.vericut_fpath;
                cbToolName.IsChecked = plugin_settings.tool_id_use_name;
                cbToolNumber.IsChecked = plugin_settings.tool_id_use_num;
                cbOpenVericut.IsChecked = plugin_settings.start_vericut;
                if (ProjectData.PmillUnitsIsInch == plugin_settings.export_options.is_inch) //settings are in the same units as project
                {
                    tModelExportTol.Text = plugin_settings.export_options.model_tol.ToString();
                    tStockExportTol.Text = plugin_settings.export_options.block_tol.ToString();
                    tFixtureExportTol.Text = plugin_settings.export_options.fixture_tol.ToString();
                    tToolExportTol.Text = plugin_settings.export_options.tool_tol.ToString();
                }
                else if (ProjectData.PmillUnitsIsInch && !plugin_settings.export_options.is_inch) //project is inch and settings are in mm
                {
                    tModelExportTol.Text = (Math.Round(plugin_settings.export_options.model_tol / 25.4, 3)).ToString();
                    tStockExportTol.Text = (Math.Round(plugin_settings.export_options.block_tol / 25.4, 3)).ToString();
                    tFixtureExportTol.Text = (Math.Round(plugin_settings.export_options.fixture_tol / 25.4, 3)).ToString();
                    tToolExportTol.Text = (Math.Round(plugin_settings.export_options.tool_tol / 25.4, 3)).ToString();
                }
                else if (!ProjectData.PmillUnitsIsInch && plugin_settings.export_options.is_inch) //project is mm and settings are in inch
                {
                    tModelExportTol.Text = (plugin_settings.export_options.model_tol * 25.4).ToString();
                    tStockExportTol.Text = (plugin_settings.export_options.block_tol * 25.4).ToString();
                    tFixtureExportTol.Text = (plugin_settings.export_options.fixture_tol * 25.4).ToString();
                    tToolExportTol.Text = (plugin_settings.export_options.tool_tol * 25.4).ToString();
                }
            }
        }

        #endregion

        private void SelectBlock()
        {
        }

        private void SelectClamps()
        {
        }

        private void bBrowseOutputDir_Click(object sender, RoutedEventArgs e)
        {
            //string proj_dir = PowerMILLAutomation.ExecuteEx("PRINT VALUE PROJECTPATH");
            System.Windows.Forms.FolderBrowserDialog oFolderDialog = new System.Windows.Forms.FolderBrowserDialog();
            oFolderDialog.Description = "Select directory for the plugin output files";
            //if (!String.IsNullOrEmpty(proj_dir))
            //{
            //    oFolderDialog.RootFolder = Environment.SpecialFolder.Desktop;
            //    oFolderDialog.SelectedPath = proj_dir;
            //}
            if (oFolderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                tOutputDir.Text = oFolderDialog.SelectedPath;
                proj_data.output_dir = tOutputDir.Text;
            }
        }

        private void rExportNC_Checked(object sender, RoutedEventArgs e)
        {
            ShowNCSelectorControls();
        }

        private void rUseNC_Checked(object sender, RoutedEventArgs e)
        {
            ShowNCSelectorControls();
        }

        private void ShowNCSelectorControls()
        {
            gridBrowseNCs.Visibility = ((bool)rExportNC.IsChecked ? System.Windows.Visibility.Hidden : System.Windows.Visibility.Visible);
            gridSelectNCs.Visibility = ((bool)rExportNC.IsChecked ? System.Windows.Visibility.Visible : System.Windows.Visibility.Hidden);
        }

        private void cbToolLibrary_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (cbToolLibrary.SelectedIndex == 2 || cbToolLibrary.SelectedIndex == 1)
            {
                rToolLibFromTemplate.IsEnabled = true;
                rToolLibFile.IsEnabled = true;
                tToolsFile.IsEnabled = true;
                bToolsFile.IsEnabled = true;
            }
            else
            {
                rToolLibFromTemplate.IsEnabled = false;
                rToolLibFile.IsEnabled = false;
                tToolsFile.IsEnabled = false;
                bToolsFile.IsEnabled = false;
            }
        }

        #region Expanders

        private void ExpendNCProgExpander(object sender, RoutedEventArgs e)
        {
            CollapseExpandExpanders(true, false, false, false, false, false, false);
        }

        private void ExpendWorkOffsetsExpander(object sender, RoutedEventArgs e)
        {
            CollapseExpandExpanders(false, true, false, false, false, false, false);
        }

        private void ExpendBlockExpander(object sender, RoutedEventArgs e)
        {
            CollapseExpandExpanders(false, false, false, true, false, false, false);
        }

        private void ExpendFixturesExpander(object sender, RoutedEventArgs e)
        {
            CollapseExpandExpanders(false, false, false, false, false, true, false);
        }

        private void ExpendPartExpander(object sender, RoutedEventArgs e)
        {
            CollapseExpandExpanders(false, false, false, false, true, false, false);
        }

        private void ExpendToolsExpander(object sender, RoutedEventArgs e)
        {
            CollapseExpandExpanders(false, false, false, false, false, false, true);
        }

        private void ExpendWorkplanesExpander(object sender, RoutedEventArgs e)
        {
            CollapseExpandExpanders(false, false, true, false, false, false, false);
        }

        private void CollapseExpandExpanders(bool nc_progs_expanded, bool work_offsets_expanded, bool ucs_expanded, bool block_expanded,
                                             bool part_expanded, bool fixt_expanded, bool tools_expanded)
        {
            eNCProgs.IsExpanded = nc_progs_expanded;
            eWorkOffsets.IsExpanded = work_offsets_expanded;
            eUCS.IsExpanded = ucs_expanded;
            eBlock.IsExpanded = block_expanded;
            ePart.IsExpanded = part_expanded;
            eFixtures.IsExpanded = fixt_expanded;
            eTools.IsExpanded = tools_expanded;
            CalcAndSetPaneHeight();
        }

        private void CalcAndSetPaneHeight()
        {
            int height = 780;

            //if (eNCProgs.IsExpanded)
            //    height = 730;
            //if (eUCS.IsExpanded)
            //    height -= 200;
            //if (eBlock.IsExpanded)
            //    height -= 100;
            PowerMILLAutomation.oPServices.AdjustPaneHeight(PowerMILLAutomation.oToken, 0, height);
            //gridMain.Height = height - 70;
            Thickness margin = tabControl1.Margin;
            margin.Bottom = 55;
            tabControl1.Margin = margin;
        }

        #endregion

        #region NC program controls

        private void bNCRight_Click(object sender, RoutedEventArgs e)
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

        private void bNCAllRight_Click(object sender, RoutedEventArgs e)
        {
            for (int i = listNCProgs.Items.Count - 1; i >= 0; i--)
            {
                object SelectedItem = listNCProgs.Items[i];
                listNCProgsSelected.Items.Insert(0, SelectedItem); // to keep the same order...
                listNCProgs.Items.Remove(SelectedItem);
            }
        }

        private void bNCLeft_Click(object sender, RoutedEventArgs e)
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

        private void bNCAllLeft_Click(object sender, RoutedEventArgs e)
        {
            for (int i = listNCProgsSelected.Items.Count - 1; i >= 0; i--)
            {
                object SelectedItem = listNCProgsSelected.Items[i];
                listNCProgs.Items.Insert(0, SelectedItem); // to keep the same order...
                listNCProgsSelected.Items.Remove(SelectedItem);
            }

        }

        private void bNCBrowse_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog oFileDlg = new Microsoft.Win32.OpenFileDialog();
            oFileDlg.Filter = "NC program file" + " (*.*)|*.*";
            oFileDlg.FileName = "";
            oFileDlg.CheckFileExists = true;
            oFileDlg.Title = "Select NC program files";
            oFileDlg.Multiselect = true;
            if (oFileDlg.ShowDialog() == true)
            {
                foreach (string fpath in oFileDlg.FileNames)
                {
                    listNCProgsBrowserSelected.Items.Add(fpath);
                }
            }

        }

        private void bNCDelete_Click(object sender, RoutedEventArgs e)
        {
            while (listNCProgsBrowserSelected.SelectedItems.Count > 0)
                listNCProgsBrowserSelected.Items.Remove(listNCProgsBrowserSelected.SelectedItems[0]);
        }

        private void ButBlockSTL_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog oFileDlg = new Microsoft.Win32.OpenFileDialog();
            oFileDlg.Filter = "STL file" + " (*.stl)|*.stl";
            oFileDlg.FileName = "";
            oFileDlg.CheckFileExists = true;
            oFileDlg.Title = "Select STL file";
            oFileDlg.Multiselect = false;
            if (oFileDlg.ShowDialog() == true)
            {
                TxtBlockSTL.Items.Clear();
                TxtBlockSTL.Items.Add(oFileDlg.FileName);
            }

        }

        private void ButBlockRemoveSTL_Click(object sender, RoutedEventArgs e)
        {
            while (TxtBlockSTL.SelectedItems.Count > 0)
                TxtBlockSTL.Items.Remove(TxtBlockSTL.SelectedItems[0]);
        }

        private void bModelsBrowse_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog oFileDlg = new Microsoft.Win32.OpenFileDialog();
            oFileDlg.Filter = "STL file" + " (*.stl)|*.stl";
            oFileDlg.FileName = "";
            oFileDlg.CheckFileExists = true;
            oFileDlg.Title = "Select STL file";
            oFileDlg.Multiselect = false;
            if (oFileDlg.ShowDialog() == true)
            {
                listModelsBrowserSelected.Items.Clear();
                listModelsBrowserSelected.Items.Add(oFileDlg.FileName);
            }

        }

        private void bModelsDelete_Click(object sender, RoutedEventArgs e)
        {
            while (listModelsBrowserSelected.SelectedItems.Count > 0)
                listModelsBrowserSelected.Items.Remove(listModelsBrowserSelected.SelectedItems[0]);
        }

        private void bFixtureBrowse_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog oFileDlg = new Microsoft.Win32.OpenFileDialog();
            oFileDlg.Filter = "STL file" + " (*.stl)|*.stl";
            oFileDlg.FileName = "";
            oFileDlg.CheckFileExists = true;
            oFileDlg.Title = "Select STL file";
            oFileDlg.Multiselect = true;
            if (oFileDlg.ShowDialog() == true)
            {
                foreach (string fpath in oFileDlg.FileNames)
                    listFixtureBrowserSelected.Items.Add(fpath);
            }

        }

        private void bFixtureDelete_Click(object sender, RoutedEventArgs e)
        {
            while (listFixtureBrowserSelected.SelectedItems.Count > 0)
                listFixtureBrowserSelected.Items.Remove(listFixtureBrowserSelected.SelectedItems[0]);
        }

        private void bToolsFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog oFileDlg = new Microsoft.Win32.OpenFileDialog();
            oFileDlg.Filter = "Tools file" + " (*.tls)|*.tls";
            oFileDlg.FileName = "";
            oFileDlg.CheckFileExists = true;
            oFileDlg.Title = "Select .tls file";
            oFileDlg.Multiselect = false;
            if (oFileDlg.ShowDialog() == true)
                tToolsFile.Text = oFileDlg.FileName;

        }

        #endregion

        #region Workplanes controls

        private void bWPRight_Click(object sender, RoutedEventArgs e)
        {
            if (listWorkplanes.SelectedItems.Count > 0)
            {
                for (int i = listWorkplanes.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object SelectedItem = listWorkplanes.SelectedItems[i];
                    listWorkplanesSelected.Items.Insert(0, SelectedItem); // to keep the same order
                    listWorkplanes.Items.Remove(SelectedItem);
                }
            }
        }

        private void bWPAllRight_Click(object sender, RoutedEventArgs e)
        {
            for (int i = listWorkplanes.Items.Count - 1; i >= 0; i--)
            {
                object SelectedItem = listWorkplanes.Items[i];
                listWorkplanesSelected.Items.Insert(0, SelectedItem); // to keep the same order...
                listWorkplanes.Items.Remove(SelectedItem);
            }
        }

        private void bWPLeft_Click(object sender, RoutedEventArgs e)
        {
            if (listWorkplanesSelected.SelectedItems.Count > 0)
            {
                for (int i = listWorkplanesSelected.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object SelectedItem = listWorkplanesSelected.SelectedItems[i];
                    listWorkplanes.Items.Insert(0, SelectedItem); // to keep the same order...
                    listWorkplanesSelected.Items.Remove(SelectedItem);
                }
            }
        }

        private void bWPAllLeft_Click(object sender, RoutedEventArgs e)
        {
            for (int i = listWorkplanesSelected.Items.Count - 1; i >= 0; i--)
            {
                object SelectedItem = listWorkplanesSelected.Items[i];
                listWorkplanes.Items.Insert(0, SelectedItem); // to keep the same order...
                listWorkplanesSelected.Items.Remove(SelectedItem);
            }

        }

        #endregion

        private void cbSetup_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (active_setup >= 0 && set_setup_settings)
                proj_data.setup_infos[active_setup] = SetSetupSettings();
            active_setup = cbSetups.SelectedIndex;
            LoadProjectComponents();
            SetUISettingsForSetup(active_setup);
        }

        private SetupInfo SetSetupSettings()
        {
            SetupInfo setup = new SetupInfo();
            string sNCStatus;
            NCProgramInfo nc_prog;

            setup.name = tSetupName.Text;
            setup.template = (String.IsNullOrEmpty(tSetupTemplate.Text) ? tProjTemplate.Text : tSetupTemplate.Text);

            /* NC programs */
            setup.nc_progs = null;
            if (Convert.ToBoolean(rExportNC.IsChecked))
            {
                setup.nc_progs = new List<NCProgramInfo>();
                for (int i = 0; i <= listNCProgsSelected.Items.Count - 1; i++)
                {
                    string nc_name = listNCProgsSelected.Items[i].ToString();

                    bool bIsError = false;
                    // Add NCProg To the list
                    string sCNCFPath = Utilities.CheckNCProgram(nc_name, out sNCStatus);
                    if (string.IsNullOrEmpty(sCNCFPath))
                    {
                        PowerMILLAutomation.Execute("DIALOGS MESSAGE OFF");
                        PowerMILLAutomation.Execute(String.Format("ACTIVATE NCPROGRAM \"{0}\" KEEP NCPROGRAM ; YES", nc_name));
                        PowerMILLAutomation.Execute("DIALOGS MESSAGE ON");
                        sCNCFPath = Utilities.CheckNCProgram(nc_name, out sNCStatus);
                    }
                    if (!string.IsNullOrEmpty(sCNCFPath))
                    {
                        // If CNFile does exist, return it
                        if (!System.IO.File.Exists(sCNCFPath))
                        {
                            Messages.ShowError(String.Format(Properties.Resources.IDS_NCFileDoesntExist, sCNCFPath, nc_name));
                            break;
                        }
                        nc_prog = new NCProgramInfo();
                        nc_prog.sName = nc_name;
                        nc_prog.sPath = sCNCFPath;
                        nc_prog.sToolRef = PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram','{0}').CoordinateType", nc_name)).Trim();
                        if (Convert.ToInt32(PowerMILLAutomation.ExecuteEx(String.Format("print = entity_exists(entity('ncprogram','{0}').ModelLocation)", nc_name))) != 0)
                        {
                            setup.model_workplane = PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram','{0}').ModelLocation.Name", nc_name)).Trim();
                            nc_prog.sAttachWorkplane = setup.model_workplane;
                            PowerMILLAutomation.Execute("Activate Workplane '" + setup.model_workplane + "'");
                            nc_prog.oNcOXY = GetOXYMatriceForNCProgram(nc_name, ref bIsError);
                        }
                        else
                        {
                            nc_prog.oNcOXY = new WorkPlaneOrigin();
                        }
                        setup.nc_progs.Add(nc_prog);
                    }
                    else
                    {
                        Messages.ShowError(Properties.Resources.IDS_NCStatusWrong, nc_name, sNCStatus);
                        break;
                    }
                }
            }
            else
            {
                setup.nc_progs = new List<NCProgramInfo>();
                foreach (string item in listNCProgsBrowserSelected.Items)
                {
                    nc_prog = new NCProgramInfo();
                    nc_prog.sPath = item;
                    setup.nc_progs.Add(nc_prog);
                }
                setup.model_workplane = (string)cbModelWP.SelectedItem;
                setup.nc_workplane = (string)cbNCWP.SelectedItem;
            }

            if (dgOffsets.Items.Count > 0)
            {
                setup.offsets = new List<WorkOffset>();
                for (int i = 0; i < dgOffsets.Items.Count; i++)
                    setup.offsets.Add((WorkOffset)dgOffsets.Items[i]);
                dgOffsets.SelectedIndex = 0;
                DataGrid_SelectionChanged(null, null);
            }

            /* Workplanes */
            setup.workplanes_to_export = new List<string>();
            foreach (string workplane in listWorkplanesSelected.Items)
                setup.workplanes_to_export.Add(workplane);
            setup.attach_workplane = cbAttachWP.Text;
            setup.attach_workplane_to = cbAttachWPTo.Text;

            /* Block */
            setup.block_toolpath = "";
            setup.block_stockmodel = "";
            setup.block_stl_fpath = "";
            if (Convert.ToBoolean(rBlockToolpath.IsChecked) && listBlocksUsed.Items.Count > 0)
                setup.block_toolpath = listBlocksUsed.Items[0].ToString();
            else if (Convert.ToBoolean(rBlockStockFile.IsChecked) && TxtBlockSTL.Items.Count > 0)
                setup.block_stl_fpath = TxtBlockSTL.Items[0].ToString();
            else if (Convert.ToBoolean(rBlockStockModel.IsChecked) && listStockModelsUsed.Items.Count > 0)
                setup.block_stockmodel = listStockModelsUsed.Items[0].ToString();
            setup.block_attach_to = cbAttachStockTo.Text;

            if (active_setup == 0)
                proj_data.cut_stock_csys = (Convert.ToBoolean(cBlockTransition.IsChecked) ? cbCutStockTransitionWP.Text : "");

            /* Parts */
            setup.part_models = null;
            setup.part_stl_fpaths = null;
            if (Convert.ToBoolean(rPartModels.IsChecked))
            {
                setup.part_models = new List<string>();
                foreach (string item in listModelsSelected.Items)
                    setup.part_models.Add(item);
            }
            else
            {
                setup.part_stl_fpaths = new List<string>();
                foreach (string item in listModelsBrowserSelected.Items)
                    setup.part_stl_fpaths.Add(item);
            }
            setup.part_attach_to = cbAttachModelsTo.Text;

            /* Fixtures */
            setup.fixture_models = null;
            setup.fixture_stl_fpaths = null;
            if (Convert.ToBoolean(rFixtureModels.IsChecked))
            {
                setup.fixture_models = new List<string>();
                foreach (string item in listFixturesSelected.Items)
                    setup.fixture_models.Add(item);
            }
            else
            {
                setup.fixture_stl_fpaths = new List<string>();
                foreach (string item in listFixtureBrowserSelected.Items)
                    setup.fixture_stl_fpaths.Add(item);
            }
            setup.fixture_attach_to = cbAttachFixtureTo.Text;

            /* Tools */
            setup.tools_to_use = -1;
            setup.tls_fpath = "";
            setup.oTools = null;

            setup.tools_to_use = cbToolLibrary.SelectedIndex;
            if (cbToolLibrary.SelectedIndex == 0) /* Create new library */
            {
                setup.tools_to_export = new List<string>();
                //foreach (string item in listToolsSelected.Items)
                //    setup.tools_to_export.Add(item);
            }
            else if (cbToolLibrary.SelectedIndex == 1 || cbToolLibrary.SelectedIndex == 2) /* Use library tools or add tools to a library*/
            {
                setup.tools_from_template = Convert.ToBoolean(rToolLibFromTemplate.IsChecked);
                if (Convert.ToBoolean(rToolLibFile.IsChecked))
                    setup.tls_fpath = tToolsFile.Text;
                if (Convert.ToBoolean(rToolLibFromTemplate.IsChecked))
                    setup.tls_fpath = ExportData.GetToolLibraryFilePath(setup.template);
            }
            if (cbToolLibrary.SelectedIndex == 2) /* Add tools to library */
            {
                setup.tools_to_export = new List<string>();
                //foreach (string item in listToolsSelected.Items)
                //    setup.tools_to_export.Add(item);
            }

            return setup;
        }

        /// <summary>
        /// Fill the NCProg List
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        private bool AddNCProgToTheList(ref List<NCProgramInfo> nc_progs)
        {
            bool bReturn = true;
            string sNCStatus;
            NCProgramInfo oNCProg;
            // Clear the Selected NCProgList
            if (nc_progs != null)
                nc_progs.Clear();
            else
                nc_progs = new List<NCProgramInfo>();

            if (listNCProgsSelected.Items.Count > 0)
            {
                int i = 0;
                //bool bIsError = false;
                // Add NCProg To the list
                string sCNFileName = Utilities.CheckNCProgram(listNCProgsSelected.Items[i].ToString(), out sNCStatus);
                if (!string.IsNullOrEmpty(sCNFileName))
                {
                    // If CNFile does exist, return it
                    if (!System.IO.File.Exists(sCNFileName))
                    {
                        Messages.ShowError(Properties.Resources.IDS_NCFileDoesntExist, sCNFileName, listNCProgsSelected.Items[i].ToString());
                        bReturn = false;
                    }
                    else
                    {
                        oNCProg = new NCProgramInfo();
                        oNCProg.sName = listNCProgsSelected.Items[i].ToString();
                        oNCProg.sPath = sCNFileName;
                        oNCProg.sToolRef = PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram','{0}').CoordinateType", oNCProg.sName)).Trim();
                        if (Convert.ToInt32(PowerMILLAutomation.ExecuteEx(String.Format("print = entity_exists(entity('ncprogram','{0}').ModelLocation)", oNCProg.sName))) != 0)
                        {
                            oNCProg.sAttachWorkplane = PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram','{0}').ModelLocation.Name", oNCProg.sName)).Trim();
                        }
                        else
                        {
                            Messages.ShowError(Properties.Resources.IDS_WorkplaneModelNotSet, listNCProgsSelected.Items[i].ToString());
                            return false;
                        }
                        //PowerMILLAutomation.Execute("Activate Workplane '" + oNCProg.sAttachWorkplane + "'");
                        //oNCProg.oOXY = GetOXYMatriceForNCProgram(oNCProg.sName, ref bIsError);
                        nc_progs.Add(oNCProg);
                        bReturn = true;
                    }
                }
                else
                {
                    Messages.ShowError(Properties.Resources.IDS_NCStatusWrong, listNCProgsSelected.Items[i].ToString(), sNCStatus);
                    bReturn = false;
                }
                return bReturn;
            }

            // Print Error Message
            if (!bReturn)
                Messages.ShowError(Properties.Resources.IDS_NCVerify);

            return bReturn;
        }

        private void tNumSetups_TextChanged(object sender, TextChangedEventArgs e)
        {
            int num_setups = 0;
            try
            {
                num_setups = Convert.ToInt32(tNumSetups.Text);
            }
            catch (Exception)
            { }
            if (num_setups > 0)
            {
                active_setup = -1;
                cbSetups.Items.Clear();
                for (int i = 1; i <= num_setups; i++)
                    cbSetups.Items.Add(i);
                active_setup = 0;
            }
            if (proj_data.setup_infos.Count < num_setups)
            {
                for (int i = proj_data.setup_infos.Count; i < num_setups; i++)
                    proj_data.setup_infos.Add(new SetupInfo());
            }
            else if (proj_data.setup_infos.Count > num_setups)
            {
                while (proj_data.setup_infos.Count > num_setups)
                    proj_data.setup_infos.RemoveAt(num_setups);
            }
        }

        private void bBrowseProjTemplate_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog oFileDlg = new Microsoft.Win32.OpenFileDialog();
            oFileDlg.Filter = "VCPROJECT file" + " (*.vcproject)|*.vcproject";
            oFileDlg.FileName = "";
            oFileDlg.CheckFileExists = true;
            oFileDlg.Title = "Select template VERICUT project";
            if (oFileDlg.ShowDialog() == true)
            {
                tProjTemplate.Text = oFileDlg.FileName;
            }
        }

        private void rPartModels_Checked(object sender, RoutedEventArgs e)
        {
            gridSelectParts.Visibility = System.Windows.Visibility.Visible;
            gridBrowseParts.Visibility = System.Windows.Visibility.Hidden;
        }

        private void rPartSTLs_Checked(object sender, RoutedEventArgs e)
        {
            gridSelectParts.Visibility = System.Windows.Visibility.Hidden;
            gridBrowseParts.Visibility = System.Windows.Visibility.Visible;
        }

        private void rFixtureModels_Checked(object sender, RoutedEventArgs e)
        {
            gridSelectFixture.Visibility = System.Windows.Visibility.Visible;
            gridBrowseFixture.Visibility = System.Windows.Visibility.Hidden;
        }

        private void rFixtureSTLs_Checked(object sender, RoutedEventArgs e)
        {
            gridSelectFixture.Visibility = System.Windows.Visibility.Hidden;
            gridBrowseFixture.Visibility = System.Windows.Visibility.Visible;

        }

        private void rBlockExport_Checked(object sender, RoutedEventArgs e)
        {
            gridSelectToolpath.Visibility = System.Windows.Visibility.Visible;
            gridBrowseBlocks.Visibility = System.Windows.Visibility.Hidden;
            gridSelectStockModels.Visibility = System.Windows.Visibility.Hidden;
        }

        private void rBlockStockModel_Checked(object sender, RoutedEventArgs e)
        {
            gridSelectToolpath.Visibility = System.Windows.Visibility.Hidden;
            gridBrowseBlocks.Visibility = System.Windows.Visibility.Hidden;
            gridSelectStockModels.Visibility = System.Windows.Visibility.Visible;
        }

        private void rBlockStockFile_Checked(object sender, RoutedEventArgs e)
        {
            gridSelectToolpath.Visibility = System.Windows.Visibility.Hidden;
            gridBrowseBlocks.Visibility = System.Windows.Visibility.Visible;
            gridSelectStockModels.Visibility = System.Windows.Visibility.Hidden;
        }

        private void bBrowseSetupTemplate_Click(object sender, RoutedEventArgs e)
        {
            if (proj_data.setup_infos.Count == 0)
                SetSetupsPlaceHolders();
            if (cbSetups.SelectedIndex < 0)
            {
                Messages.ShowError("You need to select setup first.");
                return;
            }

            Microsoft.Win32.OpenFileDialog oFileDlg = new Microsoft.Win32.OpenFileDialog();
            oFileDlg.Filter = "VCPROJECT file" + " (*.vcproject)|*.vcproject";
            oFileDlg.FileName = "";
            oFileDlg.CheckFileExists = true;
            oFileDlg.Title = "Select template VERICUT project for the setup";
            if (oFileDlg.ShowDialog() == true)
            {
                tSetupTemplate.Text = oFileDlg.FileName;
                proj_data.setup_infos[cbSetups.SelectedIndex].template = tSetupTemplate.Text;
                proj_data.setup_infos[cbSetups.SelectedIndex].VerifySetupInfo("", false);
                UpdateUIAttachCompsAndSubsystems(proj_data.setup_infos[cbSetups.SelectedIndex].setup_attach_components, proj_data.setup_infos[cbSetups.SelectedIndex].setup_subsystems);
            }
        }

        private void SetSetupsPlaceHolders()
        {
            int num_setups = 0;
            try
            {
                if (!String.IsNullOrWhiteSpace(tNumSetups.Text))
                {
                    num_setups = Convert.ToInt32(tNumSetups.Text);
                }
            }
            catch (Exception)
            { }
            if (proj_data.setup_infos.Count < num_setups)
            {
                for (int i = proj_data.setup_infos.Count; i < num_setups; i++)
                    proj_data.setup_infos.Add(new SetupInfo());

            }
        }

        private void UpdateUIAttachCompsAndSubsystems(List<string> attach_components, List<string> sybsystems)
        {
            cFromComp.Items.Clear();
            cbAttachStockTo.Items.Clear();
            cbAttachModelsTo.Items.Clear();
            cbAttachFixtureTo.Items.Clear();
            cbAttachWPTo.Items.Clear();
            cSubsystem.Items.Clear();
            if (attach_components != null)
            {
                foreach (string component in attach_components)
                {
                    cFromComp.Items.Add(component);
                    cbAttachStockTo.Items.Add(component);
                    cbAttachModelsTo.Items.Add(component);
                    cbAttachFixtureTo.Items.Add(component);
                    cbAttachWPTo.Items.Add(component);
                }
            }
            if (sybsystems != null)
            {
                foreach (string subsystem in sybsystems)
                    cSubsystem.Items.Add(subsystem);
            }
            cSubsystem.SelectedIndex = 0;
            cFromComp.SelectedIndex = 0;
            cToUCS.SelectedIndex = 0;
        }

        #region block controls

        private void ButBlockAdd_Click(object sender, RoutedEventArgs e)
        {
            if (listBlocks.SelectedItems.Count > 0)
            {
                if (listBlocksUsed.Items.Count > 0)
                {
                    for (int i = listBlocksUsed.Items.Count - 1; i >= 0; i--)
                    {
                        object SelectedItem = listBlocksUsed.Items[i];
                        listBlocks.Items.Insert(0, SelectedItem); // to keep the same order...
                        listBlocksUsed.Items.Remove(SelectedItem);
                    }
                }
                object SelectedItem2 = listBlocks.SelectedItems[0];
                listBlocksUsed.Items.Insert(0, SelectedItem2); // to keep the same order
                listBlocks.Items.Remove(SelectedItem2);
            }

            if (listBlocks.SelectedItems.Count > 0)
            {
                object SelectedItem = listBlocks.SelectedItems[0];
                listBlocksUsed.Items.Insert(0, SelectedItem); // to keep the same order
                listBlocks.Items.Remove(SelectedItem);
            }
        }

        private void ButBlockRemove_Click(object sender, RoutedEventArgs e)
        {
            if (listBlocksUsed.SelectedItems.Count > 0)
            {
                for (int i = listBlocksUsed.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object SelectedItem = listBlocksUsed.SelectedItems[i];
                    listBlocks.Items.Insert(0, SelectedItem); // to keep the same order...
                    listBlocksUsed.Items.Remove(SelectedItem);
                }
            }
        }

        private void ButStockModelsAdd_Click(object sender, RoutedEventArgs e)
        {
            if (listStockModels.SelectedItems.Count > 0)
            {
                if (listStockModelsUsed.Items.Count > 0)
                {
                    for (int i = listStockModelsUsed.Items.Count - 1; i >= 0; i--)
                    {
                        object SelectedItem = listStockModelsUsed.Items[i];
                        listStockModels.Items.Insert(0, SelectedItem); // to keep the same order...
                        listStockModelsUsed.Items.Remove(SelectedItem);
                    }
                }
                object SelectedItem2 = listStockModels.SelectedItems[0];
                listStockModelsUsed.Items.Insert(0, SelectedItem2); // to keep the same order
                listStockModels.Items.Remove(SelectedItem2);
            }
        }

        private void ButStockModelsRemove_Click(object sender, RoutedEventArgs e)
        {
            if (listStockModelsUsed.SelectedItems.Count > 0)
            {
                for (int i = listStockModelsUsed.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object SelectedItem = listStockModelsUsed.SelectedItems[i];
                    listStockModels.Items.Insert(0, SelectedItem); // to keep the same order...
                    listStockModelsUsed.Items.Remove(SelectedItem);
                }
            }
        }

        #endregion

        #region model controls

        private void bModelsRight_Click(object sender, RoutedEventArgs e)
        {
            if (listModels.SelectedItems.Count > 0)
            {
                for (int i = listModels.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object SelectedItem = listModels.SelectedItems[i];
                    listModelsSelected.Items.Insert(0, SelectedItem); // to keep the same order
                    listModels.Items.Remove(SelectedItem);
                }
            }
        }

        private void bModelsAllRight_Click(object sender, RoutedEventArgs e)
        {
            for (int i = listModels.Items.Count - 1; i >= 0; i--)
            {
                object SelectedItem = listModels.Items[i];
                listModelsSelected.Items.Insert(0, SelectedItem); // to keep the same order...
                listModels.Items.Remove(SelectedItem);
            }
        }

        private void bModelsLeft_Click(object sender, RoutedEventArgs e)
        {
            if (listModelsSelected.SelectedItems.Count > 0)
            {
                for (int i = listModelsSelected.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object SelectedItem = listModelsSelected.SelectedItems[i];
                    listModels.Items.Insert(0, SelectedItem); // to keep the same order...
                    listModelsSelected.Items.Remove(SelectedItem);
                }
            }
        }

        private void bModelsAllLeft_Click(object sender, RoutedEventArgs e)
        {
            for (int i = listModelsSelected.Items.Count - 1; i >= 0; i--)
            {
                object SelectedItem = listModelsSelected.Items[i];
                listModels.Items.Insert(0, SelectedItem); // to keep the same order...
                listModelsSelected.Items.Remove(SelectedItem);
            }

        }

        #endregion

        #region fixture controls

        private void bFixtureRight_Click(object sender, RoutedEventArgs e)
        {
            if (listFixtures.SelectedItems.Count > 0)
            {
                for (int i = listFixtures.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object SelectedItem = listFixtures.SelectedItems[i];
                    listFixturesSelected.Items.Insert(0, SelectedItem); // to keep the same order
                    listFixtures.Items.Remove(SelectedItem);
                }
            }
        }

        private void bFixtureAllRight_Click(object sender, RoutedEventArgs e)
        {
            for (int i = listFixtures.Items.Count - 1; i >= 0; i--)
            {
                object SelectedItem = listFixtures.Items[i];
                listFixturesSelected.Items.Insert(0, SelectedItem); // to keep the same order...
                listFixtures.Items.Remove(SelectedItem);
            }
        }

        private void bFixtureLeft_Click(object sender, RoutedEventArgs e)
        {
            if (listFixturesSelected.SelectedItems.Count > 0)
            {
                for (int i = listFixturesSelected.SelectedItems.Count - 1; i >= 0; i--)
                {
                    object SelectedItem = listFixturesSelected.SelectedItems[i];
                    listFixtures.Items.Insert(0, SelectedItem); // to keep the same order...
                    listFixturesSelected.Items.Remove(SelectedItem);
                }
            }
        }

        private void bFixtureAllLeft_Click(object sender, RoutedEventArgs e)
        {
            for (int i = listFixturesSelected.Items.Count - 1; i >= 0; i--)
            {
                object SelectedItem = listFixturesSelected.Items[i];
                listFixtures.Items.Insert(0, SelectedItem); // to keep the same order...
                listFixturesSelected.Items.Remove(SelectedItem);
            }

        }

        #endregion

        private void rSelectAlignedWPs_Checked(object sender, RoutedEventArgs e)
        {
            FindWPsAlignedWithAttachWPAndUpdateUI();
        }

        private void cbAttachWP_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            FindWPsAlignedWithAttachWPAndUpdateUI();
        }

        private void FindWPsAlignedWithAttachWPAndUpdateUI()
        {
            WorkPlaneOrigin orig;
            string attach_wp;
            List<string> wps = new List<string>();
            attach_wp = Convert.ToString(cbAttachWP.SelectedValue);
            if (Convert.ToBoolean(rSelectAlignedWPs.IsChecked) && !String.IsNullOrEmpty(attach_wp))
            {
                wps.Add(attach_wp);
                WorkPlaneOrigin attach_wp_orig = new WorkPlaneOrigin();
                if (cbAttachWP.SelectedValue.ToString().ToUpper() != "GLOBAL")
                {
                    attach_wp_orig = MainModule.GetWorkplaneCoords(PowerMILLAutomation.oToken, Convert.ToString(cbAttachWP.SelectedValue), false, Properties.Resources.IDS_WorkplaneFailedToGet);
                }
                foreach (string wp in proj_data.sUCSs)
                {
                    if (!wp.Equals(attach_wp))
                    {
                        if (wp.ToUpper() == "GLOBAL")
                        {
                            orig = new WorkPlaneOrigin();
                        }
                        else
                        {
                            orig = MainModule.GetWorkplaneCoords(PowerMILLAutomation.oToken, wp, false, Properties.Resources.IDS_WorkplaneFailedToGet);
                        }
                        if (orig.dXi == attach_wp_orig.dXi &&
                            orig.dXj == attach_wp_orig.dXj &&
                            orig.dXk == attach_wp_orig.dXk &&
                            orig.dYi == attach_wp_orig.dYi &&
                            orig.dYj == attach_wp_orig.dYj &&
                            orig.dYk == attach_wp_orig.dYk)
                            wps.Add(wp);
                    }
                }
                wps.Sort();
                List<string> other_wps = new List<string>(proj_data.sUCSs);
                foreach (string wp in wps)
                    other_wps.Remove(wp);
                listWorkplanes.Items.Clear();
                listWorkplanesSelected.Items.Clear();
                foreach (string wp in wps)
                    listWorkplanesSelected.Items.Add(wp);
                foreach (string wp in other_wps)
                    listWorkplanes.Items.Add(wp);
            }
        }

        private void rSelectUsedTools_Checked(object sender, RoutedEventArgs e)
        {
            if (listNCProgsSelected.Items.Count <= 0) return;
        }

        private void DataGrid_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (dgOffsets.SelectedIndex < 0) return;
            WorkOffset selected_offset = (WorkOffset)dgOffsets.Items[dgOffsets.SelectedIndex];

            if (cProgName.Items.Contains(selected_offset.name))
            {
                cProgName.SelectedItem = selected_offset.name;
            }
            else
            {
                cProgName.SelectedIndex = 0;
            }
            tRegister.Text = selected_offset.register;
            if (cSubsystem.Items.Contains(selected_offset.subsystem))
            {
                cSubsystem.SelectedItem = selected_offset.subsystem;
            }
            else
            {
                cSubsystem.SelectedIndex = 0;
            }
            if (cFromComp.Items.Contains(selected_offset.from_component))
            {
                cFromComp.SelectedItem = selected_offset.from_component;
            }
            else
            {
                cFromComp.SelectedIndex = 0;
            }
            if (cToUCS.Items.Contains(selected_offset.to_ucs))
            {
                cToUCS.SelectedItem = selected_offset.to_ucs;
            }
            else
            {
                cToUCS.SelectedIndex = 0;
            }
        }

        private void bAddWorkOffset_Click(object sender, RoutedEventArgs e)
        {
            if (!String.IsNullOrEmpty((string)cProgName.SelectedItem) &&
                !String.IsNullOrEmpty(tRegister.Text) &&
                !String.IsNullOrEmpty((string)cSubsystem.SelectedItem) &&
                !String.IsNullOrEmpty((string)cFromComp.SelectedItem) &&
                !String.IsNullOrEmpty((string)cToUCS.SelectedItem))
            {
                WorkOffset new_offset = new WorkOffset()
                {
                    name = (string)cProgName.SelectedItem,
                    register = tRegister.Text,
                    subsystem = (string)cSubsystem.SelectedItem,
                    from_component = (string)cFromComp.SelectedItem,
                    to_ucs = (string)cToUCS.SelectedItem
                };
                List<WorkOffset> temp_list = new List<WorkOffset>();
                foreach (WorkOffset offset in dgOffsets.Items)
                {
                    if (offset.name == new_offset.name &&
                        offset.register == new_offset.register &&
                        offset.subsystem == new_offset.subsystem &&
                        offset.to_ucs == new_offset.to_ucs &&
                        offset.from_component == new_offset.from_component)
                    {
                        Messages.ShowError(Properties.Resources.IDS_WorkOffsetExists);
                        return;
                    }
                }
                dgOffsets.Items.Add(new_offset);
            }
        }

        private void bDeleteWorkOffset_Click(object sender, RoutedEventArgs e)
        {
            if (dgOffsets.SelectedIndex < 0) return;
            dgOffsets.Items.RemoveAt(dgOffsets.SelectedIndex);
        }

        private void bModifyWorkOffset_Click(object sender, RoutedEventArgs e)
        {
            if (dgOffsets.SelectedIndex < 0) return;
            dgOffsets.Items[dgOffsets.SelectedIndex] = new WorkOffset()
            {
                name = (string)cProgName.SelectedItem,
                register = tRegister.Text,
                subsystem = (string)cSubsystem.SelectedItem,
                from_component = (string)cFromComp.SelectedItem,
                to_ucs = (string)cToUCS.SelectedItem
            };
        }

        private void bBrowseVericutPath_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog oFileDlg = new Microsoft.Win32.OpenFileDialog();
            oFileDlg.Filter = "VERICUT" + " (*.bat)|*.bat";
            oFileDlg.FileName = "";
            oFileDlg.CheckFileExists = true;
            oFileDlg.Title = "Select VERICUT path";
            if (oFileDlg.ShowDialog() == true)
                tVericutPath.Text = oFileDlg.FileName;
        }

        private void cBlockTransition_Click(object sender, RoutedEventArgs e)
        {
            if (cBlockTransition.IsChecked == true)
            {
                cbCutStockTransitionWP.Focusable = true;
                cbCutStockTransitionWP.IsEnabled = true;
                cbCutStockTransitionWP.Visibility = Visibility.Visible;
            }
            else
            {
                cbCutStockTransitionWP.IsEditable = false;
                cbCutStockTransitionWP.Focusable = false;
                cbCutStockTransitionWP.IsEnabled = false;
                cbCutStockTransitionWP.Visibility = Visibility.Hidden;
            }
        }

        private void tProjTemplate_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (File.Exists(tProjTemplate.Text.Trim()))
            {
                proj_data.proj_template = tProjTemplate.Text.Trim();
                if (proj_data.setup_infos.Count > 0)
                {
                    for (int i = 0; i < proj_data.setup_infos.Count; ++i)
                    {
                        proj_data.setup_infos[i].ReloadTemplate(proj_data.proj_template, false);
                    }
                    if (sender != null)
                    {
                        VerifyProjectData(false);
                    }
                    SetUISettingsForSetup(active_setup < 0 ? 0 : active_setup);
                }
            }
        }

        private void tSetupTemplate_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (File.Exists(tSetupTemplate.Text.Trim()) &&
                active_setup >= 0 && active_setup < proj_data.setup_infos.Count())
            {
                proj_data.setup_infos[active_setup].ReloadTemplate(tSetupTemplate.Text.Trim(), false);
                string warning_message = proj_data.setup_infos[active_setup].VerifySetupInfo("", false);
                if (!String.IsNullOrWhiteSpace(warning_message))
                {
                    MessageBox.Show(warning_message,
                        Messages.PluginName, MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                SetUISettingsForSetup(active_setup);
            }
        }
    }
}
