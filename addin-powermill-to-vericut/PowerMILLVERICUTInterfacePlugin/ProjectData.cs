// -----------------------------------------------------------------------
// Copyright 2018 Autodesk, Inc. All rights reserved.
// 
// Use of this software is subject to the terms of the Autodesk license
// agreement provided at the time of installation or download, or which
// otherwise accompanies this software in either electronic or hard copy form.
// -----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using PowerMILLExporter;

namespace PowerMILLVERICUTInterfacePlugin
{
    [Serializable]
    public class WorkOffset
    {
        public string name { get; set; }
        public string register { get; set; }
        public string subsystem { get; set; }
        public string from_component { get; set; }
        public string to_ucs { get; set; }
    }

    [Serializable]
    public class SetupInfo
    {
        public string name;
        
        public string template;
        
        public List<string> setup_attach_components = null;
        public List<string> setup_subsystems = null;

        //public List<string> nc_names = null;
        //public List<string> nc_fpaths = null;
        //public WorkPlaneOrigin nc_origin = null;
        public List<NCProgramInfo> nc_progs = null;
        //public string nc_tool_ref = "";

        public List<WorkOffset> offsets;

        public List<string> workplanes_to_export;
        public string attach_workplane = "";
        public string model_workplane = "";
        public string nc_workplane = "";
        public string attach_workplane_to = "";
      
        public string block_toolpath = "";
        public string block_stockmodel = "";
        public string block_stl_fpath = "";
        public string block_attach_to = "";

        public List<string> part_models = null;
        public List<string> part_stl_fpaths = null;
        public string part_attach_to = "";

        public List<string> fixture_models = null;
        public List<string> fixture_stl_fpaths = null;
        public string fixture_attach_to = "";

        public int tools_to_use = -1;
        public bool tools_from_template = true;
        public List<string> tools_to_export = null;
        [XmlIgnore]
        public PowerMILLExporter.Tools.Tools oTools;
        public string tls_fpath = "";

        public List<string> toolpaths = new List<string>();

        // List of tools
        public static int iInitToolNum;

        private List<string> pmtoolpaths;
        private List<string> nc_programs;
        private List<string> models;
        private List<string> stock_models;

        internal string VerifySetupInfo(string prefix, bool is_error)
        {
            string warning_message = "";

            // Verify setup template and load components and subsystems.
            bool verify_components = true;
            if (!VerifyFile(ref template, ref warning_message, prefix))
            {
                warning_message += prefix + String.Format(Properties.Resources.IDS_NoTemplateVerification) + "\n";
                verify_components = false;
                template = "";
            }
            if (String.IsNullOrWhiteSpace(template))
            {
                verify_components = false;
            }
            if (verify_components)
            {
                setup_attach_components = FindAllAttachComponents(System.IO.Path.GetFullPath(template));
                setup_attach_components.Sort();
                setup_subsystems = VericutMachineReader.GetMachineSubsystems(System.IO.Path.GetFullPath(template));
                setup_subsystems.Sort();
            }
                
            pmtoolpaths = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.Toolpaths);
            nc_programs = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.NCPrograms);
            models = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.Models);
            stock_models = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.StockModels);

            VerifyWorkplane(ref attach_workplane, ref warning_message, prefix);
            if (verify_components)
            {
                VerifyComponent(ref attach_workplane_to, ref warning_message, prefix);
            }

            if (verify_components)
            {
                VerifyComponent(ref block_attach_to, ref warning_message, prefix);
            }
            if (is_error)
            {
                VerifyFile(ref block_stl_fpath, ref warning_message, prefix);
            }

            VerifyStockModel(ref block_stockmodel, ref warning_message, prefix);
            if (is_error)
            {
                VerifyFile(ref block_toolpath, ref warning_message, prefix);
            }

            if (verify_components)
            {
                VerifyComponent(ref fixture_attach_to, ref warning_message, prefix);
            }

            VerifyModels(fixture_models, ref warning_message, prefix);
            if (is_error)
            {
                VerifyFiles(fixture_stl_fpaths, ref warning_message, prefix);
            }

            VerifyWorkplane(ref model_workplane, ref warning_message, prefix);

            VerifyNCPrograms(nc_progs, ref warning_message, prefix);
            VerifyWorkplane(ref nc_workplane, ref warning_message, prefix);

            if (verify_components)
            {
                VerifyWorkOffsets(offsets, ref warning_message, prefix);
            }

            if (verify_components)
            {
                VerifyComponent(ref part_attach_to, ref warning_message, prefix);
            }
            VerifyModels(part_models, ref warning_message, prefix);
            if (is_error)
            {
                VerifyFiles(part_stl_fpaths, ref warning_message, prefix);
                VerifyFile(ref tls_fpath, ref warning_message, prefix);
            }

            VerifyToolpaths(toolpaths, ref warning_message, prefix);
            VerifyWorkplanes(workplanes_to_export, ref warning_message, prefix);

            return warning_message;
        }

        internal void VerifyWorkOffsets(List<WorkOffset> work_offsets, ref string warning_message, string prefix)
        {
            if (work_offsets == null) return;

            //List<WorkOffset> new_work_offsets = new List<WorkOffset>();
            foreach (WorkOffset work_offset in work_offsets)
            {
                if (VerifyWorkOffset(work_offset, ref warning_message, prefix))
                {
                    //new_work_offsets.Add(work_offset);
                }
                else
                {
                    warning_message += prefix + String.Format(
                        String.Format(Properties.Resources.IDS_InvalidWorkOffset), name, work_offset.name) + "\n";
                }
            }
            //work_offsets = new_work_offsets;
        }

        internal bool VerifyWorkOffset(WorkOffset work_offset, ref string warning_message, string prefix)
        {
            if (work_offset == null) return false;

            bool is_valid = true;

            string str = work_offset.from_component;
            is_valid = is_valid && VerifyComponent(ref str, ref warning_message, prefix);

            foreach (string subsystem in setup_subsystems)
            {
                bool exist = false;
                if (subsystem.Trim().ToUpper() == work_offset.subsystem.Trim().ToUpper())
                {
                    exist = true;
                    break;
                }
                if (!exist)
                {
                    is_valid = false;
                }
            }

            str = work_offset.to_ucs;
            is_valid = is_valid && VerifyWorkplane(ref str, ref warning_message, prefix);

            return is_valid;
        }

        internal void VerifyToolpaths(List<string> toolpaths, ref string warning_message, string prefix)
        {
            if (toolpaths == null) return;

            List<string> new_toolpaths = new List<string>();
            foreach (string toolpath in toolpaths)
            {
                if (VerifyToolpath(toolpath, ref warning_message, prefix))
                {
                    new_toolpaths.Add(toolpath);
                }
            }
            toolpaths = new_toolpaths;
        }

        internal bool VerifyToolpath(string toolpath, ref string warning_message, string prefix)
        {
            if (!String.IsNullOrWhiteSpace(toolpath))
            {
                bool exist = false;

                foreach (string s in pmtoolpaths)
                {
                    if (s.Trim().ToUpper() == toolpath.Trim().ToUpper())
                    {
                        exist = true;
                        break;
                    }
                }

                if (!exist)
                {
                    warning_message += prefix + String.Format(
                        Properties.Resources.IDS_ToolpathDoesntExist, toolpath) + "\n";

                    return false;
                }
            }
            else
            {
                return false;
            }

            return true;
        }

        internal void VerifyNCPrograms(List<NCProgramInfo> nc_programs, ref string warning_message, string prefix)
        {
            if (nc_programs == null) return;

            List<NCProgramInfo> new_nc_programs = new List<NCProgramInfo>();
            foreach (NCProgramInfo nc_program in nc_programs)
            {
                if (VerifyNCProgram(nc_program, ref warning_message, prefix))
                {
                    new_nc_programs.Add(nc_program);
                }
            }
            nc_programs = new_nc_programs;
        }

        internal bool VerifyNCProgram(NCProgramInfo nc_program, ref string warning_message, string prefix)
        {
           if (nc_program == null) return false;

           if (!String.IsNullOrWhiteSpace(nc_program.sName))
           {
               bool exist = false;

               foreach (string s in nc_programs)
               {
                   if (s.Trim().ToUpper() == nc_program.sName.Trim().ToUpper())
                   {
                       exist = true;
                       break;
                   }
               }

               if (!exist)
               {
                   warning_message += prefix + String.Format(
                       Properties.Resources.IDS_NCProgramDoesntExist, nc_program.sName) + "\n";

                   return false;
               }
           }
           else
           {
               return false;
           }

           return true;
        }

        internal void VerifyFiles(List<string> files, ref string warning_message, string prefix)
        {
            if (files == null) return;

            //List<string> new_files = new List<string>();
            for (int i = 0; i < files.Count(); ++i)
            {
                string str = files[i];
                VerifyFile(ref str, ref warning_message, prefix);
                //if (!String.IsNullOrWhiteSpace(str))
                //{
                //    new_files.Add(str);
               // }
            }
            //files = new_files;
        }

        internal void VerifyModels(List<string> models, ref string warning_message, string prefix)
        {
            if (models == null) return;

            List<string> new_models = new List<string>();
            for(int i = 0; i < models.Count(); ++i) {
                string str = models[i];
                VerifyModel(ref str, ref warning_message, prefix);
                if (!String.IsNullOrWhiteSpace(str))
                {
                    new_models.Add(str);
                }
            }
            models = new_models;
        }

        internal void VerifyModel(ref string model, ref string warning_message, string prefix)
        {
            if (!String.IsNullOrWhiteSpace(model))
            {
                bool exist = false;

                foreach (string m in models)
                {
                    if (m.Trim().ToUpper() == model.Trim().ToUpper())
                    {
                        exist = true;
                        break;
                    }
                }

                if (!exist)
                {
                    warning_message += prefix + String.Format(
                        Properties.Resources.IDS_NoModel + "\n", model);
                    model = "";
                }
            }
        }

        internal void VerifyStockModel(ref string stock, ref string warning_message, string prefix)
        {
            if (!String.IsNullOrWhiteSpace(stock))
            {
                bool exist = false;

                foreach (string model in stock_models)
                {
                    if (model.Trim().ToUpper() == stock.Trim().ToUpper())
                    {
                        exist = true;
                        break;
                    }
                }

                if (!exist)
                {
                    warning_message += prefix + String.Format(
                        Properties.Resources.IDS_NoStockModel + "\n", stock);
                    stock = "";
                }
            }
        }

        internal bool VerifyComponent(ref string cname, ref string warning_message, string prefix)
        {
            if (!String.IsNullOrWhiteSpace(cname))
            {
                bool exist = false;
                foreach (string comp in setup_attach_components)
                {
                    if (comp.Trim().ToUpper() == cname.Trim().ToUpper())
                    {
                        exist = true;
                        break;
                    }
                }

                if (!exist)
                {
                    warning_message += prefix + String.Format(
                        Properties.Resources.IDS_NoComponent, cname, name) + "\n";
                    return false;
                }
            }

            return true;
        }

        internal static void VerifyWorkplanes(List<string> workplanes, ref string warning_message, string prefix)
        {
            if (workplanes == null) return;

            List<string> new_workplanes = new List<string>();
            foreach (string workplane in workplanes)
            {
                string str = workplane;
                if (VerifyWorkplane(ref str, ref warning_message, prefix))
                {
                    new_workplanes.Add(workplane);
                }
            }
            workplanes = new_workplanes;
        }

        internal static bool VerifyWorkplane(ref string wp, ref string warning_message, string prefix)
        {
            if (String.IsNullOrWhiteSpace(wp) || wp.Trim().ToUpper() == "GLOBAL")
            {
                wp = "Global";
                return true;
            }

            bool exist = false;

            foreach (PowerMILL.Workplane oWorkplane in PowerMILLAutomation.oPServices.Project.Workplanes)
            {
                if (oWorkplane.Name.Trim().ToUpper() == wp.Trim().ToUpper())
                {
                    exist = true;
                    break;
                }
            }

            if (!exist)
            {
                warning_message += prefix + String.Format(Properties.Resources.IDS_WorkplaneDoesntExist, wp) + "\n";
                wp = "Global";

                return false;
            }

            return true;
        }

        internal static void VerifyDirectory(ref string path, ref string warning_message, string prefix)
        {
            if (!String.IsNullOrWhiteSpace(path) && !Directory.Exists(path))
            {
                warning_message += prefix + String.Format(Properties.Resources.IDS_DirNotExist, path) + "\n";
                //path = "";
            }
        }

        internal static bool VerifyFile(ref string path, ref string warning_message, string prefix)
        {
            if (!String.IsNullOrWhiteSpace(path) && !File.Exists(path))
            {
                warning_message += prefix + String.Format(Properties.Resources.IDS_FileNotExist, path) + "\n";
                //path = "";
                return false;
            }
            return true;
        }

        private List<string> FindAllAttachComponents(string template_fpath)
        {
            string fcontent;
            string str_to_add;
            List<string> components = new List<string>();
            MatchCollection allMatches;
            Regex expr;

            try
            {
                fcontent = File.ReadAllText(template_fpath);
                //expr = new Regex(@"<Component Name=""([a-z]|[A-Z]|[0-9][)*""");
                expr = new Regex(@"<Component Name="".*?""");
                allMatches = expr.Matches(fcontent);
                foreach (Match match in allMatches)
                {
                    str_to_add = match.Value.Replace("<Component Name=\"", "").Replace("\"", "").Trim();
                    if (!components.Contains(str_to_add))
                        components.Add(str_to_add);
                }
            }
            catch (Exception Ex)
            {
                EventLogger.WriteExceptionToEvengLog(Ex.Message, "FindAllAttachComponents");
            }
            finally
            {
                fcontent = "";
                allMatches = null;
                expr = null;
            }
            return components;
        }

        internal void ReloadTemplate(string path, bool top_level)
        {
            if (top_level && !String.IsNullOrWhiteSpace(template))
            {
                if (File.Exists(template))
                {
                    return;
                }
                template = "";
            }
            string full_path = null;
            try
            {
                full_path = System.IO.Path.GetFullPath(path);
            }
            catch
            {
                full_path = null;
            }
            if (full_path != null) {
                if (!top_level)
                {
                    if (File.Exists(full_path) && String.IsNullOrWhiteSpace(template))
                    {
                        template = path;
                    }
                }
                setup_attach_components = FindAllAttachComponents(full_path);
                setup_attach_components.Sort();
                setup_subsystems = VericutMachineReader.GetMachineSubsystems(full_path);
                setup_subsystems.Sort();
            }
        }

        //private List<string> FindAllSubsystems(string template_fpath)
        //{
        //    string fcontent;
        //    string str_to_add;
        //    List<string> subsystems = new List<string>();
        //    MatchCollection allMatches;
        //    Regex expr;

        //    try
        //    {
        //        fcontent = File.ReadAllText(template_fpath);
        //        expr = new Regex(@"Subsystem="".*?""");
        //        allMatches = expr.Matches(fcontent);
        //        foreach (Match match in allMatches)
        //        {
        //            str_to_add = match.Value.Replace("Subsystem=\"", "").Replace("\"", "").Trim();
        //            if (!subsystems.Contains(str_to_add))
        //                subsystems.Add(str_to_add);
        //        }
        //        subsystems.Sort();
        //    }
        //    catch (Exception Ex)
        //    {
        //        EventLogger.WriteExceptionToEvengLog(Ex.Message, "FindAllSubsystems");
        //    }
        //    finally
        //    {
        //        fcontent = "";
        //        allMatches = null;
        //        expr = null;
        //    }
        //    return subsystems;

        //}

    }

    [Serializable]
    public class ProjectData
    {
        public static bool PmillUnitsIsInch { get; set; }
        public static string ProjectPath { get; set; }
        public static string ProjectName { get; set; }
        public static string ProjectPluginDataPath { get; set; }
        public string cut_stock_csys = "";
        public List<SetupInfo> setup_infos = new List<SetupInfo>();
        public string output_dir = "";
        public string proj_template = "";
        public List<string> sUCSs = new List<string>();

        public void Clear()
        {
            ProjectPath = "";
            ProjectName = "";
            ProjectPluginDataPath = "";
            cut_stock_csys = "";
            setup_infos.Clear();
            output_dir = "";
            proj_template = "";
            sUCSs.Clear();
        }
    }
}
