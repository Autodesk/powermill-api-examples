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
using System.Text;
using System.Text.RegularExpressions;
using PowerMILLExporter.Tools;
using System.Xml.Serialization;

namespace PowerMILLExporter
{
    public class WorkPlaneOrigin
    {
        // Matrice OXY for models
        public double dX { get; set; }
        public double dY { get; set; }
        public double dZ { get; set; }
        public double dXi { get; set; }
        public double dXj { get; set; }
        public double dXk { get; set; }
        public double dYi { get; set; }
        public double dYj { get; set; }
        public double dYk { get; set; }
        public double dZi { get; set; }
        public double dZj { get; set; }
        public double dZk { get; set; }
        public double dXangle { get; set; }
        public double dYangle { get; set; }
        public double dZangle { get; set; }

        public WorkPlaneOrigin()
        {
            dX = dY = dZ = 0;
            dXi = 1;
            dXj = 0;
            dXk = 0;
            dYi = 0;
            dYj = 1;
            dYk = 0;
            dZi = 0;
            dZj = 0;
            dZk = 1;
            dXangle = dYangle = dZangle = 0;
        }

        public WorkPlaneOrigin(double x, double y, double z,
                               double xi, double xj, double xk,
                               double yi, double yj, double yk,
                               double zi, double zj, double zk,
                               double aX, double aY, double aZ)
        {
            dX = x;
            dY = y;
            dZ = z;
            dXi = xi;
            dXj = xj;
            dXk = xk;
            dYi = yi;
            dYj = yj;
            dYk = yk;
            dZi = zi;
            dZj = zj;
            dZk = zk;
            dXangle = aX;
            dYangle = aY;
            dZangle = aZ;
        }

        public string ToXYZString()
        {
            return String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0},{1},{2},{3},{4},{5},{6},{7},{8}", dX, dY, dZ, dXi, dXj, dXk, dYi, dYj, dYk);
        }
    }

    public class NCProgramInfo
    {
        //Program name
        public string sName { get; set; }
        //Program file path
        public string sPath { get; set; }

        //Matrice OXY for nc code
        public WorkPlaneOrigin oNcOXY { get; set; }

        //List of NCProgram tools
        public List<NCProgramTool> oTools { get; set; }

        //Tool reference: Centre or Tip
        public string sToolRef { get; set; }

        //Attach workplane
        public string sAttachWorkplane { get; set; }

        //NC code workplane
        public string sNCWorkplane { get; set; }

        public bool bExport { get; set; }

        public NCProgramInfo()
        {
            sPath = "";
            oNcOXY = new WorkPlaneOrigin();
            oTools = new List<NCProgramTool>();
            sToolRef = "";
            sAttachWorkplane = "";
            sNCWorkplane = "";
        }
    }

    public class Point3D
    {
        public double X, Y, Z;

        public Point3D(double x, double y, double z)
        {
            X = x;
            Y = y;
            Z = z;
        }
    }

    [Serializable]
    public class ExportOptions
    {
        public bool is_inch;
        public double block_tol;
        public double model_tol;
        public double fixture_tol;
        public double tool_tol;

        public ExportOptions() {}
    }

    public class MainModule
    {
        /// <summary>
        /// Get ToolNumbers and ToolName in NCProgram
        /// </summary>
        /// <param name="NCProgramName"></param>
        /// <remarks></remarks>
        public static void FillNCProgramTools(string NCProgramName, List<string> ToolpathesList, int NC_Index, out List<NCProgramTool> Tools, ref int InitToolNumber)
        {
            Tools = null;
            string Tool_Name;
            bool First_Tool = true;
            int Text_Blocks = 0;
            List<string> project_Tool_List = new List<string>();

            for (int i = 0; i < ToolpathesList.Count; i++)
            {
                project_Tool_List = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.Tools);
                NCProgramTool NCProgTool = new NCProgramTool();
                Tool_Name = PowerMILLAutomation.GetParameterValueTerse(String.Format("components(entity('ncprogram','{0}'))[{1}].Tool.Name", NCProgramName, i)).Trim();
                //This section was added to match every nc program tool with a tool from the project.  If no match is found, it means that the toolpath has no tool, so it is a text block.
                //The text block variable will be increased of one and used to compensate the list that matches toolpath with tools because text blocks do not appear into the ncprog toolpath list but do in the tool list.
                if (Tool_Name.ToUpper().Contains("ERROR"))
                {
                    Text_Blocks = Text_Blocks + 1;
                    continue;
                }
                NCProgTool.Name = Tool_Name;
                NCProgTool.Number = Convert.ToInt32(PowerMILLAutomation.GetParameterValueTerse(String.Format("entity('ncprogram','{0}').nctoolpath[{1}].ToolNumber.Value", NCProgramName, (i-Text_Blocks))).Trim());
                if (First_Tool && NC_Index == 0)
                    InitToolNumber = NCProgTool.Number;
                NCProgTool.Toolpath = ToolpathesList[(i- Text_Blocks)];
                if (Tools == null) Tools = new List<NCProgramTool>();
                Tools.Add(NCProgTool);
                First_Tool = false;

            }
        }

        /// <summary>
        /// Check if the Tool Already Exist in the List
        /// </summary>
        /// <param name="ToolName"></param>
        /// <param name="ToolType"></param>
        /// <returns></returns>CheckExistingTool(sToolName, iToolNumber, sToolType))
        /// <remarks></remarks>
        public static bool CheckExistingTool(string ToolName, int ToolNumber, string ToolType, Tools.Tools oTools)
        {
            try
            {
                // Case function of the Type
                switch (ToolType)
                {
                    case "end_mill":
                    case "tip_radiused":
                    case "ball_nosed":
                        return oTools.oMillingTools.Any(tool => tool.Name == ToolName && tool.ToolNumber == ToolNumber);
                    case "tipped_disc":
                    case "dovetail":
                        return oTools.oTopAndSideMillTools.Any(tool => tool.Name == ToolName && tool.ToolNumber == ToolNumber);
                    case "drill":
                        return oTools.oDrillingTools.Any(tool => tool.Name == ToolName && tool.ToolNumber == ToolNumber);
                    case "tap":
                    case "thread_mill":
                        return oTools.oTapTools.Any(tool => tool.Name == ToolName && tool.ToolNumber == ToolNumber);
                    case "taper_spherical":
                        return oTools.oTaperSphericalTools.Any(tool => tool.Name == ToolName && tool.ToolNumber == ToolNumber);
                    case "taper_tipped":
                        return oTools.oTaperTippedTools.Any(tool => tool.Name == ToolName && tool.ToolNumber == ToolNumber);
                    case "off_centre_tip_rad":
                        return oTools.oSpecialMILLTools.Any(tool => tool.Name == ToolName && tool.ToolNumber == ToolNumber);
                    case "form":
                    case "barrel":
                    case "routing":
                        return oTools.o2DTools.Any(tool => tool.Name == ToolName && tool.ToolNumber == ToolNumber);
                }
            }
            catch (Exception Ex)
            {
                Console.WriteLine(Ex.Message);
            }
            return true;
        }

        /// <summary>
        /// Copy the programs into the output directory
        /// </summary>
        /// <returns>Exeption if copy failed</returns>
        /// <remarks></remarks>
        public static bool CopySelectedNCPrograms(List<NCProgramInfo> NCPrograms, string OutputDir)
        {
            bool bReturn = true;
            if (NCPrograms == null) return bReturn;

            EventLogger.WriteToEvengLog("Copy nc programs to the output folder");
            for (int i = 0; i < NCPrograms.Count; i++)
            {
                try
                {
                    string sNCProgPath = System.IO.Path.Combine(OutputDir, System.IO.Path.GetFileName(NCPrograms[i].sPath));
                    System.IO.File.Copy(NCPrograms[i].sPath, sNCProgPath);
                    NCPrograms[i].sPath = sNCProgPath;
                    bReturn = System.IO.File.Exists(sNCProgPath);
                    if (!bReturn) break;
                }
                catch 
                {
                    bReturn = false;
                    break;
                }
            }
            return true;
        }

        /// <summary>
        /// Check, if current PowerMill project has a named workplane.
        /// </summary>
        /// <param name="wp">Workplane name to check.</param>
        /// <returns>true - workplane exists, false - no workplane.</returns>
        public static bool PowerMillHasWorkplane(string wp)
        {
            if (String.IsNullOrWhiteSpace(wp) || wp.ToUpper() == "GLOBAL") return true;

            // Turn messages off, if necessary
            bool messages_displayed = PowerMILLAutomation.oPServices.RequestInformation("MessagesDisplayed") == "true";
            if (messages_displayed)
            {
                PowerMILLAutomation.oPServices.DoCommand(PowerMILLAutomation.oToken, "DIALOGS MESSAGE OFF");
            }

            // Check the workplane exists
            string response = PowerMILLAutomation.GetParameterValueTerse("entity_exists('workplane', '" + wp + "')");

            // Turn messages back on, if they were on originally
            if (messages_displayed)
            {
                PowerMILLAutomation.oPServices.DoCommand(PowerMILLAutomation.oToken, "DIALOGS MESSAGE ON");
            }

            return response == "1";
        }

        /// <summary>
        /// Get coordinates of a workplanes
        /// </summary>
        /// <param name="token"></param>
        /// <param name="wp"></param>
        /// <param name="in_relation_to_active_wp"></param>
        /// <param name="err_msg"></param>
        /// <returns></returns>
        public static WorkPlaneOrigin GetWorkplaneCoords(string token, string wp, bool in_relation_to_active_wp, string err_msg)
        {
            WorkPlaneOrigin wpO = new WorkPlaneOrigin();

            if (wp.ToUpper() == "GLOBAL") {
                return wpO;
            }

            if (!PowerMillHasWorkplane(wp))
            {
                if (err_msg != null) Messages.ShowError(err_msg);
                return wpO;
            }

            // Turn messages off, if necessary
            bool messages_displayed = PowerMILLAutomation.oPServices.RequestInformation("MessagesDisplayed") == "true";
            if (messages_displayed)
            {
                PowerMILLAutomation.oPServices.DoCommand(PowerMILLAutomation.oToken, "DIALOGS MESSAGE OFF");
            }

            // Print the workplane info
            object wp_info;
            PowerMILLAutomation.oPServices.DoCommandEx(token, "SIZE WORKPLANE '" + wp + "'", out wp_info);

            // Turn messages back on, if they were on originally
            if (messages_displayed) PowerMILLAutomation.oPServices.DoCommand(token, "DIALOGS MESSAGE ON");

            string wp_info_str = wp_info as string;
            if (wp_info_str == null)
            {
                Messages.ShowError(err_msg);
                if (messages_displayed) PowerMILLAutomation.oPServices.DoCommand("DIALOGS MESSAGE ON", token);
                return null;
            }

            // Convert into lines
            string[] lines = wp_info_str.Split('\n');

            // Parse the WP properties
            List<Point3D> coords = new List<Point3D>();
            try
            {
                // Look for lines containing three coordinates
                Regex r = new Regex(String.Format("{0} +{0} +{0}", "(-?[0-9]+\\" + System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator + "[0-9]+)"));
                foreach (string line in lines)
                {
                    Match m = r.Match(line);
                    if (m.Success)
                    {
                        double x = double.Parse(m.Groups[1].ToString());
                        double y = double.Parse(m.Groups[2].ToString());
                        double z = double.Parse(m.Groups[3].ToString());
                        coords.Add(new Point3D(x, y, z));
                    }
                }
            }
            catch (Exception)
            {
                Messages.ShowError(err_msg);
                if (messages_displayed) PowerMILLAutomation.oPServices.DoCommand("DIALOGS MESSAGE ON", token);
                return new WorkPlaneOrigin();
            }

            // We expect either 5 values or 10 values, depending on whether the workplane is active
            if (coords.Count != 5 && coords.Count != 10)
            {
                Messages.ShowError(err_msg);
                if (messages_displayed) PowerMILLAutomation.oPServices.DoCommand("DIALOGS MESSAGE ON", token);
                return new WorkPlaneOrigin();
            }

            // Build the output info
            StringBuilder sb = new StringBuilder();
            if (in_relation_to_active_wp && coords.Count == 10)
            {
                wpO = new WorkPlaneOrigin(coords[5].X, coords[5].Y, coords[5].Z,
                                          coords[6].X, coords[6].Y, coords[6].Z,
                                          coords[7].X, coords[7].Y, coords[7].Z,
                                          coords[8].X, coords[8].Y, coords[8].Z,
                                          coords[9].X, coords[9].Y, coords[9].Z);
            }
            if (!in_relation_to_active_wp)// && coords.Count == 5)
            {
                wpO = new WorkPlaneOrigin(coords[0].X, coords[0].Y, coords[0].Z,
                                          coords[1].X, coords[1].Y, coords[1].Z,
                                          coords[2].X, coords[2].Y, coords[2].Z,
                                          coords[3].X, coords[3].Y, coords[3].Z,
                                          coords[4].X, coords[4].Y, coords[4].Z);

            }

            return wpO;
        }

    }
}
