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
using System.Text;
using System.Text.RegularExpressions;
using Delcam.Plugins.Framework;

namespace PowerMILLExporter
{
    public static class PowerMILLAutomation
    {
        // variables
        public static PowerMILL.PluginServices oPServices;
        public static string oToken;

        // enum
        public enum enumUnit
        {
            mm, inches, undefined
        }

        public enum enumEntity
        {
            MachineTools, NCPrograms, Toolpaths, Tools, Boundaries, Patterns, FeatureSets, Workplanes, Models, StockModels
        }

        /// <summary>
        /// new instance of the class
        /// </summary>
        /// <param name="comms"></param>
        //public static PowerMILLAutomation(IPluginCommunicationsInterface comms)
        //{
        //    oPServices = comms.Services;
        //    oToken = comms.Token;
        //}

        public static void SetVariables(IPluginCommunicationsInterface comms)
        {
            oPServices = comms.Services;
            oToken = comms.Token;
        }

        public static bool IsSTLOutputInMM()
        {
            string str = GetParameterValueTerse("Powermill.Export.OutputMetricSTL");
            bool value = true;

            if (!String.IsNullOrEmpty(str))
                value = Convert.ToBoolean(Convert.ToInt32(str));

            return value;
        }

        public static void SetSTLOutputUnits(bool bIsMM)
        {
            Execute(String.Format("$Powermill.Export.OutputMetricSTL = \"{0}\"", (bIsMM ? "TRUE" : "FALSE")));
        }

        /// <summary>
        /// This execute a PowerMILL command
        /// </summary>
        /// <param name="Command"></param>
        public static void Execute(string Command)
        {
            oPServices.QueueCommand(oToken, Command);
        }

        public static void ShowMessageInProgressBar(string Message)
        {
            oPServices.StartProgressBreak(oToken, Message);
            oPServices.EndProgressBreak(oToken);
            oPServices.UpdateProgressBreak(oToken, 10);
        }

        /// <summary>
        /// This execute a PowerMILL command and return the answer in a string
        /// </summary>
        /// <param name="Command"></param>
        /// <returns></returns>
        public static string ExecuteEx(string Command)
        {
            object sResponse;
            string sOutput;
            int error = oPServices.DoCommandEx(oToken, Command, out sResponse);
            if (error != 0)
                if (!Command.Equals("PRINT VALUE PROJECTPATH"))
                {
                    throw new Exception("Failed to obtain command value: " + Command);
                }

            sOutput = (sResponse as string).Trim();
            if (sOutput.Contains(Command))
            {
                List<string> sTabCNInfos = new List<string>(sOutput.Split((char)13));
                if (sTabCNInfos.Count > 0)
                {
                    sTabCNInfos.RemoveAt(0);
                    if (sTabCNInfos.Count > 0)
                    {
                        sTabCNInfos.RemoveAt(0);
                        return String.Join(((char)13).ToString(), sTabCNInfos.ToArray()).Trim();
                    }

                }
                
            }
            else
                return sOutput;
            return "";
        }

        /// <summary>
        /// Turn off dialogs
        /// </summary>
        public static void DialogsOFF()
        {
            Execute("DIALOGS MESSAGE OFF");
            //Execute("DIALOGS ERROR OFF");
        }

        /// <summary>
        /// Turn on dialogs
        /// </summary>
        public static void DialogsON()
        {
            Execute("DIALOGS MESSAGE ON");
            //Execute("DIALOGS ERROR ON");
        }

        /// <summary>
        /// Returns a list of PM entity name
        /// </summary>
        public static List<string> GetListOf(enumEntity Entity)
        {

            List<string> oList = new List<string>();

            switch (Entity)
            {
                case enumEntity.MachineTools:
                    {
                        foreach (PowerMILL.MachineTool oEntity in oPServices.Project.MachineTools)
                        {
                            oList.Add(oEntity.Name);
                        }
                        return oList;
                    }
                case enumEntity.NCPrograms:
                    {
                        foreach (PowerMILL.NCProgram oEntity in oPServices.Project.NCPrograms)
                        {
                            oList.Add(oEntity.Name);
                        }
                        return oList;
                    }
                case enumEntity.Toolpaths:
                    {
                        foreach (PowerMILL.Toolpath oEntity in oPServices.Project.Toolpaths)
                        {
                            oList.Add(oEntity.Name);
                        }
                        return oList;
                    }
                case enumEntity.Tools:
                    {
                        foreach (PowerMILL.Tool oEntity in oPServices.Project.Tools)
                        {
                            oList.Add(oEntity.Name);
                        }
                        return oList;
                    }
                case enumEntity.Boundaries:
                    {
                        foreach (PowerMILL.Boundary oEntity in oPServices.Project.Boundaries)
                        {
                            oList.Add(oEntity.Name);
                        }
                        return oList;
                    }
                case enumEntity.Patterns:
                    {
                        foreach (PowerMILL.Pattern oEntity in oPServices.Project.Patterns)
                        {
                            oList.Add(oEntity.Name);
                        }
                        return oList;
                    }
                case enumEntity.FeatureSets:
                    {
                        foreach (PowerMILL.FeatureSet oEntity in oPServices.Project.FeatureSets)
                        {
                            oList.Add(oEntity.Name);
                        }
                        return oList;
                    }
                case enumEntity.Workplanes:
                    {
                        foreach (PowerMILL.Workplane oEntity in oPServices.Project.Workplanes)
                        {
                            oList.Add(oEntity.Name);
                        }
                        return oList;
                    }
                case enumEntity.Models:
                    {
                        foreach (PowerMILL.Model oEntity in oPServices.Project.Models)
                        {
                            oList.Add(oEntity.Name);
                        }
                        return oList;
                    }
                case enumEntity.StockModels:
                    {
                        foreach (PowerMILL.StockModel oEntity in oPServices.Project.StockModels)
                        {
                            oList.Add(oEntity.Name);
                        }
                        return oList;
                    }
                default:
                    {
                        throw new Exception("Unknown entity (GetListOf): " + Entity.ToString());
                    }
            }
        }

        /// <summary>
        /// Returns a parameter
        /// </summary>
        public static string GetParameterValue(string Parameter)
        {
            return oPServices.GetParameterValue(oToken, Parameter);
        }

        /// <summary>
        /// Returns a parameter (terse)
        /// </summary>
        public static string GetParameterValueTerse(string Parameter)
        {
            return oPServices.GetParameterValueTerse(oToken, Parameter);
        }

        /// <summary>
        /// return the PM units
        /// </summary>
        public static enumUnit Units
        {
            get
            {
                if (GetParameterValueTerse("units").ToUpper() == "METRIC")
                {
                    return enumUnit.mm;
                }
                else
                {
                    return enumUnit.inches;
                }
            }
        }

        public static void ActivateWorkplane(string WorkplaneName)
        {
            if (String.IsNullOrEmpty(WorkplaneName) || WorkplaneName.ToUpper() == "GLOBAL")
                Execute("DEACTIVATE WORKPLANE");
            else
                Execute("ACTIVATE WORKPLANE '" + WorkplaneName + "'");

        }

        public static void CreateTransformWorkplane(string WorkplaneName)
        {
            Execute("DEACTIVATE Workplane");
            Execute("CREATE WORKPLANE '" + WorkplaneName + "'");

        }

        public static void DeleteTransformWorkplane(string WorkplaneName)
        {
            Execute("DELETE WORKPLANE '" + WorkplaneName + "'");
        }

        public static void SetExportTolerance(double Tolerance)
        {
            ExecuteEx(String.Format("$Powermill.Export.TriangleTolerance = \"{0}\"", Tolerance));
        }

        /// <summary>
        /// Export CAD models (PART)
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static void ExportModels(List<string> ModelList, string Tol, string DestDir, out List<string> ModelListPaths, out bool bReturn)
        {
            string sOutputPath;

            EventLogger.WriteToEvengLog("Export models");
            bReturn = true;
            ModelListPaths = new List<string>();
            if (ModelList == null) return;
            if (ModelList.Count == 0) return;

            ExecuteEx(String.Format("$Powermill.Export.TriangleTolerance = \"{0}\"", Tol));
            ModelListPaths = new List<string>();

            if (ModelList.Count == 1)
            {
                sOutputPath = ExportModel(ModelList[0], DestDir);
                if (!String.IsNullOrEmpty(sOutputPath))
                    ModelListPaths.Add(sOutputPath);
            }
            else
            {
                sOutputPath = ExportAllModels(ModelList, DestDir);
                ModelListPaths.Add(sOutputPath);
                if (!String.IsNullOrEmpty(sOutputPath))
                    ModelListPaths.Add(sOutputPath);
            }

            for (int i = 0; i < ModelListPaths.Count; i++)
            {
                bReturn = System.IO.File.Exists(ModelListPaths[i]);
                if (!bReturn) break;
            }
        }

        /// <summary>
        /// Export Clamp List
        /// </summary>
        /// <remarks></remarks>
        public static void ExportClamps(List<string> ClampList, string Tol, string DestDir, out List<string> ClampListPaths, out bool bReturn)
        {
            string sOutputPath;

            EventLogger.WriteToEvengLog("Export clamps");
            bReturn = true;
            ClampListPaths = new List<string>();
            if (ClampList == null) return;
            if (ClampList.Count == 0) return;

            PowerMILLAutomation.ExecuteEx(String.Format("$Powermill.Export.TriangleTolerance = \"{0}\"", Tol));
            ClampListPaths = new List<string>();
            foreach (string model in ClampList)
            {
                sOutputPath = ExportModel(model, DestDir);
                if (!String.IsNullOrEmpty(sOutputPath))
                    ClampListPaths.Add(sOutputPath);
            }

            for (int i = 0; i < ClampListPaths.Count; i++)
            {
                bReturn = System.IO.File.Exists(ClampListPaths[i]);
                if (!bReturn) break;
            }
        }

        /// <summary>
        /// Export the SelectedToolpath Block
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static void ExportBlock(List<string> Blocks, string Tol, string DestDir, out List<string> BlockPaths, out bool bReturn)
        {
            string sOutputPath;

            EventLogger.WriteToEvengLog("Export block(s)");
            bReturn = true;
            BlockPaths = new List<string>();
            if (Blocks == null) return;
            if (Blocks.Count == 0) return;

            PowerMILLAutomation.ExecuteEx(String.Format("$Powermill.Export.TriangleTolerance = \"{0}\"", Tol));
            BlockPaths = new List<string>();
            foreach (string model in Blocks)
            {
                sOutputPath = ExportBlock(model, DestDir);
                if (!String.IsNullOrEmpty(sOutputPath))
                    BlockPaths.Add(sOutputPath);
            }

            for (int i = 0; i < BlockPaths.Count; i++)
            {
                bReturn = System.IO.File.Exists(BlockPaths[i]);
                if (!bReturn) break;
            }
        }

        /// <summary>
        /// Export the Selected stock model block
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static void ExportStockModel(List<string> StockModelNames, string Tol, string DestDir, out List<string> BlockPaths, out bool bReturn)
        {
            string sOutputPath;

            EventLogger.WriteToEvengLog("Export block(s)");
            BlockPaths = new List<string>();
            bReturn = true;
            if (StockModelNames == null) return;
            if (StockModelNames.Count == 0) return;

            PowerMILLAutomation.ExecuteEx(String.Format("$Powermill.Export.TriangleTolerance = \"{0}\"", Tol));
            BlockPaths = new List<string>();

            foreach (string model in StockModelNames)
            {
                sOutputPath = ExportStockModel(model, DestDir);
                if (!String.IsNullOrEmpty(sOutputPath))
                    BlockPaths.Add(sOutputPath);
            }

            for (int i = 0; i < BlockPaths.Count; i++)
            {
                bReturn = System.IO.File.Exists(System.IO.Path.Combine(DestDir, BlockPaths[i]));
                if (!bReturn) break;
            }
        }

        public static string ExportModel(string model_name, string output_dir)
        {
            if (String.IsNullOrEmpty(model_name)) return "";

            string sOutputPath = Path.Combine(output_dir, model_name + ".stl");
            EventLogger.WriteToEvengLog(String.Format("Export model {0} to {1}", model_name, sOutputPath));
            ActivateWorkplane("");
            Execute("EDIT MODEL '" + model_name + "' DESELECT ALL");
            Execute(String.Format("EXPORT MODEL '{0}' '{1}' YES", model_name, sOutputPath));

            return sOutputPath;
        }

        public static string ExportAllModels(List<string> models, string output_dir)
        {
            if (models == null) return "";
            if (models.Count == 0) return "";

            string mult_models = "";
            ActivateWorkplane("");
            Execute("EDIT MODEL ALL DESELECT ALL");
            foreach (string model in models)
            {
                mult_models += String.Format("'{0}' ", model);
                Execute("EDIT MODEL '" + model + "' SELECT ALL");
            }
            EventLogger.WriteToEvengLog("Select models " + mult_models);
            string sOutputPath = Path.Combine(output_dir, "models.stl");
            EventLogger.WriteToEvengLog(String.Format("Export models {0} to {1}", mult_models, sOutputPath));
            //Execute(String.Format("EXPORT MODEL {0} '{1}' YES", mult_models, sOutputPath));
            //EXPORT MODEL ALL 'C:\Vericut\IMTS_Demo\Vericut_Project\Part\all.stl' YES
            Execute(String.Format("EXPORT MODEL ALL '{0}' YES", sOutputPath));

            return sOutputPath;
        }

        public static string ExportBlock(string toolpath_name, string output_dir)
        {
            if (String.IsNullOrEmpty(toolpath_name)) return "";

            string sOutputPath = Path.Combine(output_dir, toolpath_name + ".stl");
            EventLogger.WriteToEvengLog(String.Format("Export block for toolpath {0} to {1}", toolpath_name, sOutputPath));
            Execute("DEACTIVATE Toolpath");
            Execute(String.Format("ACTIVATE Toolpath '{0}'", toolpath_name));
            //Execute("EDIT BLOCK RESET");
            Execute(String.Format("EXPORT BLOCK ; '{0}' YES", sOutputPath));

            return sOutputPath;
        }


        public static string ExportStockModel(string stock_model_name, string output_dir)
        {
            if (String.IsNullOrEmpty(stock_model_name)) return "";

            EventLogger.WriteToEvengLog("Export stock model " + stock_model_name);
            string sOutputPath = Path.Combine(output_dir, stock_model_name + ".stl");
            EventLogger.WriteToEvengLog(String.Format("Export stock model {0} to {1}", stock_model_name, sOutputPath));
            ActivateWorkplane("");
            //Execute(String.Format("TRANSFORM TYPE WORKPLANE TRANSFORM MODEL '{0}'", ListModels[i]));
            Execute("EDIT MODEL '" + stock_model_name + "' DESELECT ALL");
            Execute(String.Format("EXPORT STOCKMODEL '{0}' '{1}' YES", stock_model_name, sOutputPath));

            return sOutputPath;
        }

        /// <summary>
        /// Export and translate the ToolHolder
        /// </summary>
        /// <param name="tool">Name of the tool</param>
        /// <param name="Type">Type to export (ToolHolder or ToolShank)</param>
        /// <returns>IGS file Path</returns>
        /// <remarks></remarks>
        public static string PMILLExportToolSegment(string tool, string Type, string outputDir)
        {

            // Change the name of export because of special characters
            string sExportToolName = tool;
            string sOutputPath = "";
            string output_segment = "";

            System.Text.RegularExpressions.Regex oRegex = new System.Text.RegularExpressions.Regex("\\W");
            sExportToolName = oRegex.Replace(sExportToolName, "");

            // Defines input and output path
            switch (Type)
            {
                case "TOOLHOLDER":
                    outputDir = System.IO.Path.Combine(outputDir, "Holders");
                    output_segment = "HOLDER";
                    break;
                case "TOOLSHANK":
                    outputDir = System.IO.Path.Combine(outputDir, "Shanks");
                    output_segment = "SHANK";
                    break;
                case "TOOL_PROFILE":
                    outputDir = System.IO.Path.Combine(outputDir, "Tools");
                    output_segment = "TIP";
                    break;
            }
            if (!System.IO.Directory.Exists(outputDir))
                System.IO.Directory.CreateDirectory(outputDir);

            sOutputPath = System.IO.Path.Combine(outputDir, sExportToolName + ".stl");

            // Export the ToolHolder of Active Tool in DGK
            PowerMILLAutomation.ExecuteEx(String.Format("ACTIVATE TOOL '{0}'", tool));
            PowerMILLAutomation.ExecuteEx(String.Format("EDIT TOOL ; EXPORT_STL {0} \"{1}\" YES", output_segment, sOutputPath)); // "YES" to overwrite if already exists

            if (System.IO.File.Exists(sOutputPath))
            {
                // Return the IGS Path
                return sOutputPath;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Get the NCProgToolpathes
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        public static List<string> GetNCProgToolpathes(string NCProgName)
        {
            List<string> sToolpathesList = new List<string>();

            // Get NC prog Infos
            PowerMILLAutomation.Execute("DIALOGS MESSAGE OFF");
            string sProgCN = PowerMILLAutomation.ExecuteEx("EDIT NCPROGRAM '" + NCProgName + "' LIST");
            PowerMILLAutomation.Execute("DIALOGS MESSAGE ON");

            // Split the Line
            string[] sTabCNInfos = Regex.Split(sProgCN, Environment.NewLine);// sProgCN.Split(Convert.ToChar("\n\r"));

            // Add NC prog to the List
            for (int i = 4; i <= sTabCNInfos.Length - 2; i++)
            {
                if (!string.IsNullOrEmpty(sTabCNInfos[i].Trim()))
                    sToolpathesList.Add(sTabCNInfos[i]);
            }

            return sToolpathesList;
        }


    }
}
