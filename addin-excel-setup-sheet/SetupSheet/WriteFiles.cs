// -----------------------------------------------------------------------
// Copyright 2019 Autodesk, Inc. All rights reserved.
//
// Use of this software is subject to the terms of the Autodesk license
// agreement provided at the time of installation or download, or which
// otherwise accompanies this software in either electronic or hard copy form.
// -----------------------------------------------------------------------

using System;
using System.Windows;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;
using Delcam.Plugins.Framework;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Globalization;

namespace SetupSheet
{
    public class ProjectInfo
    {
        public string Name { get; set; }
        public string Path { get; set; }
        public string OrderNumber { get; set; }
        public string Programmer { get; set; }
        public string PartName { get; set; }
        public string Customer { get; set; }
        public string Date { get; set; }
        public string Notes { get; set; }
        public string MachModelsMaxX { get; set; }
        public string MachModelsMaxY { get; set; }
        public string MachModelsMaxZ { get; set; }
        public string MachModelsMinX { get; set; }
        public string MachModelsMinY { get; set; }
        public string MachModelsMinZ { get; set; }
        public string ModelsList { get; set; }
        public string CombinedNCTPList { get; set; }
        public string ExcelTemplate { get; set; }
        public double TotalTime { get; set; }
    }
    public class ToolInfo {
        public string Name { get; set; }
        public double Number { get; set; }
        public double Diameter { get; set; }
        public string ToolType { get; set; }
        public double CutterLength { get; set; }
        public string HolderName { get; set; }
        public double ToolOverhang { get; set; }
        public double NumberOfFlutes { get; set; }
        public double ToolTipRadius { get; set; }
        public double ToolRadOffset { get; set; }
        public double ToolLengthOffset { get; set; }
        public double ToolGaugeLength { get; set; }
        public string Description { get; set; }
    }

    public class ToolpathInfo
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string Notes { get; set; }
        public string ToolName { get; set; }
        public double ToolNumber { get; set; }
        public double ToolDiameter { get; set; }
        public string ToolType { get; set; }
        public double ToolCutterLength { get; set; }
        public string ToolHolderName { get; set; }
        public double ToolOverhang { get; set; }
        public double ToolNumberOfFlutes { get; set; }
        public double ToolTipRadius { get; set; }
        public double ToolRadOffset { get; set; }
        public double ToolLengthOffset { get; set; }
        public string ToolDescription { get; set; }
        public string ToolpathType { get; set; }
        public double ToolGaugeLength { get; set; }
        public string CutterComp { get; set; }
        public double Feed { get; set; }
        public double Speed { get; set; }
        public double IPT { get; set; }
        public double SFM { get; set; }
        public double PlungeFeed { get; set; }
        public double SkimFeed { get; set; }
        public string Coolant { get; set; }
        public double Stepover { get; set; }
        public double DOC { get; set; }
        public string FirstLeadInType { get; set; }
        public string SecondLeadInType { get; set; }
        public string FirstLeadOutType { get; set; }
        public string SecondLeadOutType { get; set; }
        public string Thickness { get; set; }
        public string AxialThickness { get; set; }
        public double Statistic_Time { get; set; }
        public string GeneralAxisType { get; set; }
        public string TPWorkplane { get; set; }
        public double RapidHeight { get; set; }
        public double Tolerance { get; set; }
        public double ToolpathMinX { get; set; }
        public double ToolpathMinY { get; set; }
        public double ToolpathMinZ { get; set; }
        public double ToolpathMaxX { get; set; }
        public double ToolpathMaxY { get; set; }
        public double ToolpathMaxZ { get; set; }
        public double SkimHeight { get; set; }
        public string TPPicture { get; set; }
        public double CuttingDistance { get; set; }
        public string NCProgName { get; set; }
    }

    class WriteFiles
    {

        public static void ExtractData(string NCProg, out List<ToolInfo> ToolData, out List<ToolpathInfo> ToolpathData)
        {
            List<string> NCProg_Toolpaths = new List<string>();
            List<string> Project_Workplanes = new List<string>();
            bool AlreadyExist = false;
            double Tool_Dia = 0;
            double Tool_Num = 0;
            double Tool_Length_Offset = 0;
            double Tool_Dia_Offset = 0;
            string Tool_Type="";
            double Cutter_Length=0;
            string Holder_Name="";
            double TOverhang=0;
            double FlutesQTY=0;
            double TipRad=0;
            string Axial_Thickness = "";
            string Radial_Thickness = "";
            string Thickness = "";
            string TPName = "";
            string ToolName="";
            string TPType="";
            double Feed=0;
            double Speed=0;
            double Plunge_Feed=0;
            double Skim_Feed=0;
            string Coolant="";
            double IPT=0;
            double SFM=0;
            string Use_Axial_Thickness = "0";
            string FirstLeadIn="";
            string SecondLeadIn="";
            string FirstLeadOut="";
            string SecondLeadOut="";
            string CutterComp = "";
            string ToolDescription = "";
            double Toolpath_Statistic = 0;
            double Stepover = 0;
            double DOC = 0;
            string AxisType = "";
            string ToolpathWP = "";
            string NCProgWP = "";
            string GeneralAxisType = "";
            string Description = "";
            string Notes = "";
            double Tolerance = 0;
            double RapidHeight = 0;
            double ToolpathMinX = 0;
            double ToolpathMinY = 0;
            double ToolpathMinZ = 0;
            double ToolpathMaxX = 0;
            double ToolpathMaxY = 0;
            double ToolpathMaxZ = 0;
            double SkimHeight = 0;
            double CuttingDistance = 0;
            double GaugeLength = 0;
            Dictionary<string, string> NCProgTpInfo = new Dictionary<string, string>();
            char Appos = (char)34;

            //CreateToolList as List to hold the tool infos
            ToolData = new List<ToolInfo>();

            //CreateToolpathList as List to hold the toolpath infos
            ToolpathData = new List<ToolpathInfo>();

                        
            //Creating a list of toolpaths and tools for each nc programs
            NCProg_Toolpaths = PowerMILLAutomation.GetNCProgToolpathes(NCProg);

            //Creating a list of all the workplanes contained in the project
            Project_Workplanes = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.Workplanes);

            NCProgWP = PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").outputworkplane.name");
            if (NCProgWP == "#ERROR: Invalid name")
            {
                NCProgWP = "world";
            }

            for (int j = 0; j <= NCProg_Toolpaths.Count - 1; j++)
                {
                //ToolDescription = PowerMILLAutomation.ExecuteEx("print par \"entity('toolpath', '" + NCProg_Toolpaths[j] + "')\"");

                Tool_Dia = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos +").tool.diameter"));
                Tool_Num = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").nctoolpath[" + j + "].toolnumber.value"));
                Tool_Length_Offset = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").nctoolpath[" + j + "].CutterLengthCompensation.Number.Value"));
                Tool_Dia_Offset = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").nctoolpath[" + j + "].CutterRadiusCompensation.Number.Value"));
                CutterComp = PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").nctoolpath[" + j + "].CutterRadiusCompensation.Type.Value");
                Tool_Type = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").tool.type");
                Cutter_Length = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").tool.length"));
                Holder_Name = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").tool.holdername");
                TOverhang = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").tool.overhang"));
                FlutesQTY = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").tool.flutes"));
                TipRad = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").tool.tipradius"));
                ToolDescription = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").tool.description");

                //Get toolpath info and build a list of all tools used in the current NC program
                TPName = NCProg_Toolpaths[j];
                if (CutterComp == "autotp")
                {
                    CutterComp = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").CNCCutterCompensation.Type");
                }
                ToolName = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").tool.name");
                Description = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").description");
                Notes = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").notes");
                TPType = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").strategy");
                TPType = TPSpecificParameters(TPType, NCProg_Toolpaths[j], out Stepover,out DOC);
                Feed = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").FeedRate.Cutting.Value"));
                Speed = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").SpindleSpeed.Value"));
                Plunge_Feed = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").FeedRate.Plunging.Value"));
                Skim_Feed = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").FeedRate.Rapid"));
                Coolant = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").coolant");
                IPT = Math.Round((Feed / Speed), 5);
                SFM = Math.Round((Math.PI * Tool_Dia * Speed) / 12, 0);
                FirstLeadIn = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").Connections.LeadIn[0].Type");
                SecondLeadIn = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").Connections.LeadIn[1].Type");
                FirstLeadOut = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").Connections.LeadOut[0].Type");
                SecondLeadOut = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").Connections.LeadOut[1].Type");
                Use_Axial_Thickness = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").UseAxialThickness");
                Radial_Thickness = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").Thickness");
                Axial_Thickness = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").AxialThickness");
                if (Use_Axial_Thickness == "0")
                {
                    Axial_Thickness= Radial_Thickness;
                }
                Toolpath_Statistic = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").statistics.totaltime"));
                ToolpathWP = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").workplane.name");
                AxisType = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").toolaxis.type");

                if ( !Project_Workplanes.Contains(ToolpathWP))
                {
                    ToolpathWP = "world";
                }
                ExtractAxisType(ToolpathWP, NCProgWP, AxisType, out GeneralAxisType);
                Tolerance = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").tolerance"));
                RapidHeight = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").Rapid.Plane.Distance"));
                SkimHeight = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").Connections.SkimDistance"));
                CuttingDistance = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").Statistics.CuttingMoves.Lengths.Linear")) + double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + NCProg_Toolpaths[j] + Appos + ").Statistics.CuttingMoves.Lengths.Arcs"));
                GetToolGaugeLength(ToolName, out GaugeLength);

                PowerMILLAutomation.GetToolpathLimits(NCProg_Toolpaths[j], out ToolpathMinX, out ToolpathMinY, out ToolpathMinZ, out ToolpathMaxX, out ToolpathMaxY, out ToolpathMaxZ);

                //Add the current toolpath info to the list
                ToolpathData.Add(new ToolpathInfo
                {
                    Name = TPName,
                    Description = Description,
                    Notes = Notes,
                    ToolName = ToolName,
                    ToolNumber = Tool_Num,
                    ToolDiameter = Tool_Dia,
                    ToolType = Tool_Type,
                    ToolCutterLength = Cutter_Length,
                    ToolHolderName = Holder_Name,
                    ToolOverhang = TOverhang,
                    ToolNumberOfFlutes = FlutesQTY,
                    ToolLengthOffset = Tool_Length_Offset,
                    ToolRadOffset = Tool_Dia_Offset,
                    ToolpathType = TPType,
                    ToolTipRadius = TipRad,
                    ToolDescription = ToolDescription,
                    ToolGaugeLength = GaugeLength,
                    Thickness = Radial_Thickness,
                    AxialThickness = Axial_Thickness,
                    CutterComp = CutterComp,
                    Feed = Feed,
                    Speed = Speed,
                    IPT = IPT,
                    SFM = SFM,
                    PlungeFeed = Plunge_Feed,
                    SkimFeed = Skim_Feed,
                    Coolant = Coolant,
                    Stepover = Stepover,
                    DOC = DOC,
                    GeneralAxisType = GeneralAxisType,
                    Statistic_Time = Toolpath_Statistic,
                    TPWorkplane = ToolpathWP,
                    Tolerance = Tolerance,
                    RapidHeight = RapidHeight,
                    SkimHeight = SkimHeight,
                    ToolpathMinX = ToolpathMinX,
                    ToolpathMinY = ToolpathMinY,
                    ToolpathMinZ = ToolpathMinZ,
                    ToolpathMaxX = ToolpathMaxX,
                    ToolpathMaxY = ToolpathMaxY,
                    ToolpathMaxZ = ToolpathMaxZ,
                    FirstLeadInType = FirstLeadIn,
                    SecondLeadInType = SecondLeadIn,
                    FirstLeadOutType = FirstLeadOut,
                    SecondLeadOutType = SecondLeadOut,
                    CuttingDistance = CuttingDistance,
                    NCProgName = NCProg
                    });

                    AlreadyExist = false;

                    if (ToolData.Any(Tool => Tool.Name == ToolName))
                    {
                        AlreadyExist = true;
                    }

                    if (AlreadyExist == false)
                    {
                        //Add the current tool info to the list
                        ToolData.Add(new ToolInfo
                        {
                            Name = ToolName,
                            Number = Tool_Num,
                            Diameter = Tool_Dia,
                            ToolType = Tool_Type,
                            CutterLength = Cutter_Length,
                            HolderName = Holder_Name,
                            ToolOverhang = TOverhang,
                            NumberOfFlutes = FlutesQTY,
                            ToolTipRadius = TipRad,
                            ToolLengthOffset = Tool_Length_Offset,
                            ToolRadOffset = Tool_Dia_Offset,
                            ToolGaugeLength = GaugeLength,
                            Description = ToolDescription
                        });
                    }

                }
            
            }

        private static void GetToolGaugeLength(string ToolName, out double GaugeLength)
        {
            char Appos = (char)34;
            string Gauge_Length_Address = "";
            PowerMILLAutomation.ExecuteEx("FORM TOOL " + Appos+ ToolName + Appos);
            string Tool_Type = PowerMILLAutomation.ExecuteEx("print $entity('tool';" + Appos + ToolName + Appos + ").Tool.Type");
            if (Tool_Type.Trim() == "end_mill")
            {
                Gauge_Length_Address = "Print formvalue pmlEndMillTool.Shell.LnkdToolHolder.TGaugeLength";
            }
            else if (Tool_Type.Trim() == "tip_radiused")
            {
                Gauge_Length_Address = "Print formvalue pmlTipRadTool.Shell.LnkdToolHolder.TGaugeLength";
            }
            else if (Tool_Type.Trim() == "ball_nosed")
            {
                Gauge_Length_Address = "Print formvalue pmlBallNosedTool.Shell.LnkdToolHolder.TGaugeLength";
            }
            else if (Tool_Type.Trim() == "taper_spherical")
            {
                Gauge_Length_Address = "Print formvalue pmlTapSphereTool.Shell.LnkdToolHolder.TGaugeLength";
            }
            else if (Tool_Type.Trim() == "taper_tipped")
            {
                Gauge_Length_Address = "Print formvalue pmlTapTipTool.Shell.LnkdToolHolder.TGaugeLength";
            }
            else if (Tool_Type.Trim() == "off_centre_tip_rad")
            {
                Gauge_Length_Address = "Print formvalue pmlOffCentreTool.Shell.LnkdToolHolder.TGaugeLength";
            }
            else if (Tool_Type.Trim() == "tipped_disc")
            {
                Gauge_Length_Address = "Print formvalue pmlTipDiscTool.Shell.LnkdToolHolder.TGaugeLength";
            }
            else if (Tool_Type.Trim() == "drill")
            {
                Gauge_Length_Address = "Print formvalue pmlDrillTool.Shell.LnkdToolHolder.TGaugeLength";
            }
            else if (Tool_Type.Trim() == "tap")
            {
                Gauge_Length_Address = "Print formvalue pmlTappingTool.Shell.LnkdToolHolder.TGaugeLength";
            }
            else if (Tool_Type.Trim() == "thread_mill")
            {
                Gauge_Length_Address = "Print formvalue pmlThreadMillTool.Shell.LnkdToolHolder.TGaugeLength";
            }
            else if (Tool_Type.Trim() == "dovetail")
            {
                Gauge_Length_Address = "Print formvalue pmlDovetailTool.Shell.LnkdToolHolder.TGaugeLength";
            }
            else if (Tool_Type.Trim() == "barrel")
            {
                Gauge_Length_Address = "Print formvalue pmlBarrelTool.Shell.LnkdToolHolder.TGaugeLength";
            }
            else if (Tool_Type.Trim() == "routing")
            {
                Gauge_Length_Address = "Print formvalue pmlRoutingTool.Shell.LnkdToolHolder.TGaugeLength";
            }
            else if (Tool_Type.Trim() == "form")
            {
                Gauge_Length_Address = "Print formvalue pmlFormTool.Shell.LnkdToolHolder.TGaugeLength";
            }

            //if tooltype is not supported yet, set the gauge length to 0 to avoid crash.
            if (Gauge_Length_Address != "")
            {
                GaugeLength = double.Parse(PowerMILLAutomation.ExecuteEx(Gauge_Length_Address));
            }
            else
            {
                GaugeLength = 0;
            }
            PowerMILLAutomation.ExecuteEx("TOOL ACCEPT");
        }
        private static void ExtractNCProg (string NCProg)
        {
            string NcProgPrint = PowerMILLAutomation.ExecuteEx("EDIT NCPROGRAM '" + NCProg + "' PRINT");

            // Split the Line
            string[] sTabCNInfos = Regex.Split(NcProgPrint, Environment.NewLine);// sProgCN.Split(Convert.ToChar("\n\r"));

            // Add NC prog to the List
            for (int i = 4; i <= sTabCNInfos.Length - 3; i++)
            {
                if (!string.IsNullOrEmpty(sTabCNInfos[i].Trim()))
                {

                }
                    
            }

        }

        private static void ExtractAxisType(string TPWorkplane, string NCPWorkplane, string TPAxisType, out string GeneralAxisType)
        {
            double TPAzimuth = 0;
            double TPElevation = 0;
            double NCAzimuth = 0;
            double NCElevation = 0;
            char Appos = (char)34;
            if (TPWorkplane != "world")
            {
                TPAzimuth = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('workplane';" + Appos + TPWorkplane + Appos + ").azimuth"));
                TPElevation = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('workplane';" + Appos + TPWorkplane + Appos + ").elevation"));
            }

            if (NCPWorkplane != "world")
            {
                NCAzimuth = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('workplane';" + Appos + NCPWorkplane + Appos + ").azimuth"));
                NCElevation = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('workplane';" + Appos + NCPWorkplane + Appos + ").elevation"));
            }
            GeneralAxisType = "3 axis";

            if (TPAxisType != "vertical")
            {
                GeneralAxisType = "5 axis";
            }
            else if (TPAzimuth != NCAzimuth || TPElevation != NCElevation)
            {
                GeneralAxisType = "3+2";
            }

        }

        public static string TPSpecificParameters(string TPType, string TPName, out double Stepover, out double DOC)
        {
            char Appos = (char)34;
            Stepover = 0;
            DOC = 0;
            string Other_Param = "";
            string Other_Param2 = "";
            CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;

            if (TPType == "raster")
            {
                TPType = "Raster";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "radial")
            {
                TPType = "Radial";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "spiral")
            {
                TPType = "Spiral";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "pattern")
            {
                TPType = "Pattern";
                Other_Param = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").PatternBasePosition");
                if (Other_Param == "drive_curve'")
                {
                    DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
                }
            }
            else if (TPType == "com_pattern")
            {
                TPType = "Commited Pattern";
            }
            else if (TPType == "com_boundary")
            {
                TPType = "Commited Boundary";
            }
            else if (TPType == "constantz")
            {
                TPType = "Constant Z";
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
            }
            else if (TPType == "offset_3d")
            {
                TPType = "3D Offset";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "pencil_corner")
            {
                TPType = "Pencil Corner";

            }
            else if (TPType == "automatic_corner")
            {
                TPType = "Automatic Corner";
            }
            else if (TPType == "along_corner")
            {
                TPType = "Along Corner";
            }
            else if (TPType == "multi_pencil_corner")
            {
                TPType = "Multipencil Corner";
            }
            else if (TPType == "rotary")
            {
                TPType = "4 Axis Rotary";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "point_projection")
            {
                TPType = "Point Projection";
            }
            else if (TPType == "line_projection")
            {
                TPType = "Line Projection";
                Other_Param = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").ProjectionStyle");
                if (Other_Param != "linear")
                {
                    Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                }
            }
            else if (TPType == "plane_projection")
            {
                TPType = "Plane Projection";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "curve_projection")
            {
                TPType = "Curve Projection";
                Other_Param = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").ProjectionStyle");
                if (Other_Param != "linear")
                {
                    Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                }
  
            }
            else if (TPType == "surface_proj")
            {
                TPType = "Surface Projection";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "profile")
            {
                TPType = "Profile";
                Other_Param = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").PatternBasePosition");
                if (Other_Param == "drive_curve")
                {
                    DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
                }
            }
            else if (TPType == "opti_constz")
            {
                TPType = "Optimised Constant Z";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "inter_constz")
            {
                TPType = "Steep and Shallow";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                Other_Param = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Cusp.Active");
                if (Other_Param == "1")
                {
                    DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
                }
            }
            else if (TPType == "swarf")
            {
                TPType = "Swarf";
                Other_Param = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").SwarfBasePosition");
                if (Other_Param != "automatic")
                {
                    DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
                }
            }
            else if (TPType == "embedded")
            {
                TPType = "Embedded Curves";
                Other_Param = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").MultipleCuts");
                if (Other_Param != "off")
                {
                    DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
                }
            }
            else if (TPType == "raster_area_clear")
            {
                TPType = "Raster Area Clearance";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").AreaClearance.ZHeights.Stepdown"));
            }
            else if (TPType == "offset_area_clear")
            {
                TPType = "Offset Area Clearance";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").AreaClearance.ZHeights.Stepdown"));
            }
            else if (TPType == "profile_area_clear")
            {
                TPType = "Profile Area Clearance";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").AreaClearance.ZHeights.Stepdown"));
            }
            else if (TPType == "drill")
            {
                TPType = "Drilling";
                Other_Param = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").drill.type");
                Other_Param2 = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").drill.spiral");
                if (Other_Param == "profile" && Other_Param2 == "1")
                {
                    Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                }
            }
            else if (TPType == "wireframe_swarf")
            {
                TPType = "Wireframe Swarf";
                Other_Param = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").SwarfBasePosition");
                if (Other_Param != "automatic")
                {
                    DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
                }
            }
            else if (TPType == "raster_flat")
            {
                TPType = "Raster Flat";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "offset_flat")
            {
                TPType = "Offset Flat";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "plunge")
            {
                TPType = "Plunge Milling";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "parametric_offset")
            {
                TPType = "Parametric Offset";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "surface_machine")
            {
                TPType = "Surface Machining";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "port_area_clear")
            {
                TPType = "Port Area Clearance";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
            }
            else if (TPType == "port_plunge")
            {
                TPType = "Port Plunge Finishing";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "port_spiral")
            {
                TPType = "Port Spiral Finishing";
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
            }
            else if (TPType == "method")
            {
                TPType = "Method";
            }
            else if (TPType == "blisk")
            {
                TPType = "Blisk/Impellor Machining";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                Other_Param = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").MultipleCuts");
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
            }
            else if (TPType == "blisk_hub")
            {
                TPType = "Blisk/Impellor Hub Finishing";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "blisk_blade")
            {
                TPType = "Blisk/Impellor Blade Finishing";
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
            }
            else if (TPType == "disc_profile")
            {
                TPType = "Disc Profile Machining";
            }
            else if (TPType == "curve_profile")
            {
                TPType = "Curve Profile Machining";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
            }
            else if (TPType == "curve_area_clear")
            {
                TPType = "Curve Area Clearance";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
            }
            else if (TPType == "face")
            {
                TPType = "Face Machining";
                Other_Param = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").PatternStyle");
                if (Other_Param !="one_pass")
                {
                    Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                }
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
            }
            else if (TPType == "chamfer")
            {
                TPType = "Chamfer Machining";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
            }
            else if (TPType == "wireframe_profile")
            {
                TPType = "Wireframe Profile Machining";
                Other_Param = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").MultipleCuts");
                if (Other_Param != "off")
                {
                    DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
                }
            }
            else if (TPType == "corner_clear")
            {
                TPType = "Corner Clearance";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "edge_break")
            {
                TPType = "Edge Break Machining";
            }
            else if (TPType == "flowline")
            {
                TPType = "Flowline Machining";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "parametric_spiral")
            {
                TPType = "Parametric Spiral Machining";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "adaptive_area_clear")
            {
                TPType = "Vortex";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "rib")
            {
                TPType = "Rib Machining";
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
            }
            else if (TPType == "blade")
            {
                TPType = "Blade Machining";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "feature_face")
            {
                TPType = "Feature Face Machining";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
            }
            else if (TPType == "feature_finishing")
            {
                TPType = "Feature Finish";
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
            }
            else if (TPType == "feature_chamfer")
            {
                TPType = "Feature Chamfer Machining";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
                DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
            }
            else if (TPType == "turn_roughing")
            {
                TPType = "Turn Roughing";
            }
            else if (TPType == "turn_finishing")
            {
                TPType = "Turn Finishing";
            }
            else if (TPType == "turn_face_roughing")
            {
                TPType = "Turn Face Roughing";
            }
            else if (TPType == "turn_face_finishing")
            {
                TPType = "Turn Face Finishing";
            }
            else if (TPType == "turn_groove_roughing")
            {
                TPType = "Turn Groove Roughing";
            }
            else if (TPType == "turn_groove_finishing")
            {
                TPType = "Turn Groove Finishing";
            }
            else if (TPType == "turn_bore_roughing")
            {
                TPType = "Turn Bore Roughing";
            }
            else if (TPType == "turn_bore_finishing")
            {
                TPType = "Turn Bore Finishing";
            }
            else if (TPType == "pocket_area_clear")
            {
                TPType = "Pocket Area Clearance";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "pocket_profile")
            {
                TPType = "Pocket Profile Machining";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "feature_external_thread")
            {
                TPType = "Feature External Thread Machining";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "slot")
            {
                TPType = "Feature Slot Machining";
                Other_Param = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").FloorOnly");
                if (Other_Param != "1")
                {
                    DOC = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").Stepdown"));
                }
            }
            else if (TPType == "feature_finishing")
            {
                TPType = "Feature Finishing";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "feature_top_fillet")
            {
                TPType = "Feature Top Fillet Machining";
            }
            else if (TPType == "surface_coating")
            {
                TPType = "Surface Coating";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "feature_construction")
            {
                TPType = "Feature Construction";
                Stepover = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + TPName + Appos + ").stepover"));
            }
            else if (TPType == "surface_inspection")
            {
                TPType = "Surface Inspection";
            }
            else if (TPType == "hole_inspection")
            {
                TPType = "Hole Inspection";
            }
            else if (TPType == "turn_pattern")
            {
                TPType = "Turn Pattern";
            }


            //Localization to output different toolpath names in local lnaguage
            //if (Thread.CurrentThread.CurrentCulture.Name.Equals("fr-CA") || Thread.CurrentThread.CurrentCulture.Name.Equals("fr-FR"))
            //{
            //    if (TPType == "Raster")
            //    {
            //        TPType = "Raster";
            //    }
            //    else if (TPType == "Radial")
            //    {
            //        TPType = "Radial";
            //    }
            //    else if (TPType == "Spiral")
            //    {
            //        TPType = "Spiral";
            //    }
            //    else if (TPType == "Pattern")
            //    {
            //        TPType = "Pattern";
            //    }
            //    else if (TPType == "Commited Pattern")
            //    {
            //        TPType = "Commited Pattern";
            //    }
            //    else if (TPType == "Commited Boundary")
            //    {
            //        TPType = "Commited Boundary";
            //    }
            //    else if (TPType == "Constant Z")
            //    {
            //        TPType = "Constant Z";
            //    }
            //    else if (TPType == "3D Offset")
            //    {
            //        TPType = "3D Offset";
            //    }
            //    else if (TPType == "Pencil Corner")
            //    {
            //        TPType = "Pencil Corner";
            //    }
            //    else if (TPType == "Automatic Corner")
            //    {
            //        TPType = "Automatic Corner";
            //    }
            //    else if (TPType == "Along Corner")
            //    {
            //        TPType = "Along Corner";
            //    }
            //    else if (TPType == "Multipencil Corner")
            //    {
            //        TPType = "Multipencil Corner";
            //    }
            //    else if (TPType == "4 Axis Rotary")
            //    {
            //        TPType = "4 Axis Rotary";
            //    }
            //    else if (TPType == "Point Projection")
            //    {
            //        TPType = "Point Projection";
            //    }
            //    else if (TPType == "Line Projection")
            //    {
            //        TPType = "Line Projection";
            //    }
            //    else if (TPType == "Plane Projection")
            //    {
            //        TPType = "Plane Projection";
            //    }
            //    else if (TPType == "Curve Projection")
            //    {
            //        TPType = "Curve Projection";

            //    }
            //    else if (TPType == "Surface Projection")
            //    {
            //        TPType = "Surface Projection";
            //    }
            //    else if (TPType == "Profile")
            //    {
            //        TPType = "Profile";
            //    }
            //    else if (TPType == "Optimised Constant Z")
            //    {
            //        TPType = "Optimised Constant Z";
            //    }
            //    else if (TPType == "Steep and Shallow")
            //    {
            //        TPType = "Steep and Shallow";
            //    }
            //    else if (TPType == "Swarf")
            //    {
            //        TPType = "Swarf";
            //    }
            //    else if (TPType == "Embedded Curves")
            //    {
            //        TPType = "Embedded Curves";
            //    }
            //    else if (TPType == "Raster Area Clearance")
            //    {
            //        TPType = "Raster Area Clearance";
            //    }
            //    else if (TPType == "Offset Area Clearance")
            //    {
            //        TPType = "Offset Area Clearance";
            //    }
            //    else if (TPType == "Profile Area Clearance")
            //    {
            //        TPType = "Profile Area Clearance";
            //    }
            //    else if (TPType == "Drilling")
            //    {
            //        TPType = "Drilling";
            //    }
            //    else if (TPType == "Wireframe Swarf")
            //    {
            //        TPType = "Wireframe Swarf";
            //    }
            //    else if (TPType == "Raster Flat")
            //    {
            //        TPType = "Raster Flat";
            //    }
            //    else if (TPType == "Offset Flat")
            //    {
            //        TPType = "Offset Flat";
            //    }
            //    else if (TPType == "Plunge Milling")
            //    {
            //        TPType = "Plunge Milling";
            //    }
            //    else if (TPType == "Parametric Offset")
            //    {
            //        TPType = "Parametric Offset";
            //    }
            //    else if (TPType == "Surface Machining")
            //    {
            //        TPType = "Surface Machining";
            //    }
            //    else if (TPType == "Port Area Clearance")
            //    {
            //        TPType = "Port Area Clearance";
            //    }
            //    else if (TPType == "Port Plunge Finishing")
            //    {
            //        TPType = "Port Plunge Finishing";
            //    }
            //    else if (TPType == "Port Spiral Finishing")
            //    {
            //        TPType = "Port Spiral Finishing";
            //    }
            //    else if (TPType == "Method")
            //    {
            //        TPType = "Method";
            //    }
            //    else if (TPType == "Blisk/Impellor Machining")
            //    {
            //        TPType = "Blisk/Impellor Machining";
            //    }
            //    else if (TPType == "Blisk/Impellor Hub Finishing")
            //    {
            //        TPType = "Blisk/Impellor Hub Finishing";
            //    }
            //    else if (TPType == "Blisk/Impellor Blade Finishing")
            //    {
            //        TPType = "Blisk/Impellor Blade Finishing";
            //    }
            //    else if (TPType == "Disc Profile Machining")
            //    {
            //        TPType = "Disc Profile Machining";
            //    }
            //    else if (TPType == "Curve Profile Machining")
            //    {
            //        TPType = "Curve Profile Machining";
            //    }
            //    else if (TPType == "Curve Area Clearance")
            //    {
            //        TPType = "Curve Area Clearance";
            //    }
            //    else if (TPType == "Face Machining")
            //    {
            //        TPType = "Face Machining";
            //    }
            //    else if (TPType == "Chamfer Machining")
            //    {
            //        TPType = "Chamfer Machining";
            //    }
            //    else if (TPType == "Wireframe Profile Machining")
            //    {
            //        TPType = "Wireframe Profile Machining";
            //    }
            //    else if (TPType == "Corner Clearance")
            //    {
            //        TPType = "Corner Clearance";
            //    }
            //    else if (TPType == "Edge Break Machining")
            //    {
            //        TPType = "Edge Break Machining";
            //    }
            //    else if (TPType == "Flowline Machining")
            //    {
            //        TPType = "Flowline Machining";
            //    }
            //    else if (TPType == "Parametric Spiral Machining")
            //    {
            //        TPType = "Parametric Spiral Machining";
            //    }
            //    else if (TPType == "Vortex")
            //    {
            //        TPType = "Vortex";
            //    }
            //    else if (TPType == "Rib Machining")
            //    {
            //        TPType = "Rib Machining";
            //    }
            //    else if (TPType == "Blade Machining")
            //    {
            //        TPType = "Blade Machining";
            //    }
            //    else if (TPType == "Feature Face Machining")
            //    {
            //        TPType = "Feature Face Machining";
            //    }
            //    else if (TPType == "Feature Finish")
            //    {
            //        TPType = "Feature Finish";
            //    }
            //    else if (TPType == "Feature Chamfer Machining")
            //    {
            //        TPType = "Feature Chamfer Machining";
            //    }
            //    else if (TPType == "Turn Roughing")
            //    {
            //        TPType = "Turn Roughing";
            //    }
            //    else if (TPType == "Turn Finishing")
            //    {
            //        TPType = "Turn Finishing";
            //    }
            //    else if (TPType == "Turn Face Roughing")
            //    {
            //        TPType = "Turn Face Roughing";
            //    }
            //    else if (TPType == "Turn Face Finishing")
            //    {
            //        TPType = "Turn Face Finishing";
            //    }
            //    else if (TPType == "Turn Groove Roughing")
            //    {
            //        TPType = "Turn Groove Roughing";
            //    }
            //    else if (TPType == "Turn Groove Finishing")
            //    {
            //        TPType = "Turn Groove Finishing";
            //    }
            //    else if (TPType == "Turn Bore Roughing")
            //    {
            //        TPType = "Turn Bore Roughing";
            //    }
            //    else if (TPType == "Turn Bore Finishing")
            //    {
            //        TPType = "Turn Bore Finishing";
            //    }
            //    else if (TPType == "Pocket Area Clearance")
            //    {
            //        TPType = "Pocket Area Clearance";
            //    }
            //    else if (TPType == "Pocket Profile Machining")
            //    {
            //        TPType = "Pocket Profile Machining";
            //    }
            //    else if (TPType == "Feature External Thread Machining")
            //    {
            //        TPType = "Feature External Thread Machining";
            //    }
            //    else if (TPType == "Feature Slot Machining")
            //    {
            //        TPType = "Feature Slot Machining";
            //    }
            //    else if (TPType == "Feature Finishing")
            //    {
            //        TPType = "Feature Finishing";
            //    }
            //    else if (TPType == "Feature Top Fillet Machining")
            //    {
            //        TPType = "Feature Top Fillet Machining";
            //    }
            //    else if (TPType == "Surface Coating")
            //    {
            //        TPType = "Surface Coating";
            //    }
            //    else if (TPType == "Feature Construction")
            //    {
            //        TPType = "Feature Construction";
            //    }
            //    else if (TPType == "Surface Inspection")
            //    {
            //        TPType = "Surface Inspection";
            //    }
            //    else if (TPType == "Hole Inspection")
            //    {
            //        TPType = "Hole Inspection";
            //    }
            //    else if (TPType == "Turn Pattern")
            //    {
            //        TPType = "Turn Pattern";
            //    }
            //}


            return TPType;
        }

        static string GetColumnName(int index)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            var value = "";

            if (index >= letters.Length)
                value += letters[index / letters.Length - 1];

            value += letters[index % letters.Length];

            return value;
        }

        public static Dictionary<string,string> ExtractTemplateData(_Worksheet Sheet)
        {
            int xlLastRow;
            int xlLastColumn;
            object[,] FullTable;
            Range xlLastCell;
            Range xlTableRange;
            Dictionary<string, string> VarsList = new Dictionary<string, string>();

            //xlLastRow = Sheet.Cells.Find("*", System.Reflection.Missing.Value,System.Reflection.Missing.Value, System.Reflection.Missing.Value, XlSearchOrder.xlByRows, XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Row;
            //xlLastColumn = Sheet.Cells.Find("*", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, XlSearchOrder.xlByColumns, XlSearchDirection.xlPrevious, false, System.Reflection.Missing.Value, System.Reflection.Missing.Value).Column;
            

            xlLastCell = Sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            xlLastColumn = xlLastCell.Column;
            xlLastRow = xlLastCell.Row;
            xlTableRange = Sheet.get_Range("A1", xlLastCell);
            
            FullTable = xlTableRange.Cells.Value;

            for (int i = 1; i <= xlLastRow; i++)
            {
                for (int j = 1; j <= xlLastColumn; j++)
                {
                   //Store specific items to a dictionnary with their adresses
                    int Instr = Convert.ToString(FullTable[i, j]).IndexOf("}");
                    if (Instr > 0)                    {
                        //Current_Column = GetColumnName(j-1);
                        VarsList.Add(Convert.ToString(j + "," + i), Convert.ToString(FullTable[i, j]) + "-" + j + "," + i);
                    }
                }
            }

            return VarsList;
        }

        public static string KeyByValue(Dictionary<string, string> dict, string val)
        {
            string key = null;
            foreach (KeyValuePair<string, string> pair in dict)
            {
                if (pair.Value == val)
                {
                    key = pair.Key;
                    break;
                }
            }
            return key;
        }

        private static void ExtractStockInfo(string Toolpath, out string StockType, out double StockX, out double StockY, out double StockZ, out double StockOD, out double StockID)
        {
            StockID = 0;
            StockOD = 0;
            StockX = 0;
            StockY = 0;
            StockZ = 0;
            StockType = "";
            char Appos = (char)34;

            StockType = PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + Toolpath + Appos + ").Block.Type");
            if (StockType == "box")
            {
                StockType = "Box";
                StockX = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + Toolpath + Appos + ").Block.Limits.XMax"))- double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + Toolpath + Appos + ").Block.Limits.XMin"));
                StockY = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + Toolpath + Appos + ").Block.Limits.YMax")) - double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + Toolpath + Appos + ").Block.Limits.YMin"));
                StockZ = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + Toolpath + Appos + ").Block.Limits.ZMax")) - double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + Toolpath + Appos + ").Block.Limits.ZMin"));
            }
            else if (StockType == "picture")
            {
                StockType = "picture";
            }
            else if (StockType == "triangles")
            {
                StockType = "3D Model";
            }
            else if (StockType == "boundary")
            {
                StockType = "Boundary";
            }
            else if (StockType == "cylinder" || StockType == "cylinder_sector")
            {
                if (StockType == "cylinder")
                {
                    StockOD = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + Toolpath + Appos + ").Block.Diameter"));
                    StockID = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + Toolpath + Appos + ").Block.InnerDiameter"));
                    StockZ = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + Toolpath + Appos + ").Block.Limits.ZMax")) - double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('toolpath';" + Appos + Toolpath + Appos + ").Block.Limits.ZMin"));
                }
                StockType = "Cylinder";
            }
            else if (StockType == "spun" || StockType == "spun_sector")
            {
                StockType = "Revolved Curve";
            }
        }

        private static void ExtractViseInfo(string ActiveWorkplane, out string ViseName, out string Para_dimensions)
        {
            List<string> Models = new List<string>();
            ViseName = "None";
            Para_dimensions = "None";

            Models = PowerMILLAutomation.GetListOf(PowerMILLAutomation.enumEntity.Models);

            foreach (string Model in Models)
            {
                int Workplane = Model.IndexOf(ActiveWorkplane);
                int Chuck_Base = Model.IndexOf("chuck_base");
                int Vise_Bottom = Model.IndexOf("fix_bottom");
                int Para_Fix = Model.IndexOf("para_fix");

                if (Workplane >= 0 && Chuck_Base > 0)
                {
                    ViseName = Model.Substring((ActiveWorkplane.Length + 1), (Chuck_Base - (ActiveWorkplane.Length + 2)));
                }
                else if (Workplane >= 0 && Vise_Bottom > 0)
                {
                    ViseName = Model.Substring((ActiveWorkplane.Length + 1), (Vise_Bottom - (ActiveWorkplane.Length + 2)));
                }
                else if (Workplane >= 0 && Para_Fix > 0)
                {
                    Para_dimensions = Model.Substring((ActiveWorkplane.Length + 1), (Para_Fix - (ActiveWorkplane.Length + 2)));
                }
            }

        }

        public static Image resizeImage(Image imgToResize, System.Drawing.Size size)
        {
            return (Image)(new Bitmap(imgToResize, size));
        }
        private static void ConvertToUsableName(string OriginalName, out string NewName)
        {
            NewName = "";
            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                NewName = OriginalName.Replace(c, '_');
            }
            NewName = NewName.Replace("''", "");
            NewName = NewName.Replace("*","");
            if (NewName.Length > 31)
            {
                NewName = NewName.Substring(NewName.Length - 31, 31);
            }
        }
        public static void ImportImage(_Worksheet oSheet, string PictureName, string Project_Path, int Row, int Column)
        {
            int Height = 0;
            int Width = 0;
            double ImageWidth = 0;
            double ImageHeight = 0;
            float Left = 0;
            float Top = 0;
            double TempVal = 0;
            double PicRatio = 0;
            double NewHeight = 0;
            double NewWidth = 0;
            double CenterOffset = 0;
            double ScaleFactor = 1.3;

            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                PictureName = PictureName.Replace(c, '_');
            }

            PictureName = Project_Path + "//Excel_Setupsheet//" + PictureName + ".png";
            if (File.Exists(PictureName))
            {
                TempVal = oSheet.Cells[Row, Column].MergeArea.Height() - 1;
                Height = (int)TempVal;

                TempVal = oSheet.Cells[Row, Column].MergeArea.Width() - 1;
                Width = (int)TempVal;

                TempVal = oSheet.Cells[Row, Column].MergeArea.Top() + 3;
                Top = (float)TempVal;

                TempVal = oSheet.Cells[Row, Column].MergeArea.Left() + 3;
                Left = (float)TempVal;

                Range oRange = (Range)oSheet.Cells[Row, Column];

                Image oImage = Image.FromFile(PictureName);
                ImageHeight = oImage.Height;
                ImageWidth = oImage.Width;

                PicRatio = ImageWidth / ImageHeight;
                NewWidth = (int)(Height * PicRatio);
                

                if (NewWidth  > Width)
                {
                    NewWidth = (Width) * ScaleFactor;
                    NewHeight = (int)(Width / PicRatio)* ScaleFactor;
                }
                else
                {
                    NewWidth = NewWidth * ScaleFactor;
                    NewHeight = (Height) * ScaleFactor;
                }
                CenterOffset = (Width - (NewWidth/ ScaleFactor)) / 2;

                oImage = resizeImage(oImage, new System.Drawing.Size((int)(NewWidth), (int)(NewHeight)));
                System.Windows.Forms.Clipboard.Clear();
                System.Windows.Forms.Clipboard.SetDataObject(oImage, true);
                System.Threading.Thread.Sleep(200);
                oSheet.Paste(oRange, PictureName);
                Pictures p = oSheet.Pictures(System.Reflection.Missing.Value) as Pictures;
                Picture pic = p.Item(p.Count) as Picture;
                pic.Left = pic.Left + CenterOffset;
                pic.Top = pic.Top + 3;

                //oRange.Select();
                
                System.Windows.Forms.Clipboard.Clear();
                //oSheet.Shapes.AddPicture(PictureName, MsoTriState.msoFalse, MsoTriState.msoCTrue, Left, Top, Width, Height);

                oSheet.Cells[Row, Column] = "";
            }
            else
            {
                oSheet.Cells[Row, Column] = "Picture Missing";
            }
        }

        private static void MatchParameter(string CurrentParameter, string Project_Path, string NCProg, ToolpathInfo Toolpath, ToolInfo Tool, ProjectInfo Project, List<ToolpathInfo> ToolpathData, bool Is_Hyperlink, _Worksheet oSheet, out string Value, out string HyperlinkText, out bool Needs_Hyperlink, string ActiveWorkplane)
        {
            Value = "";
            Needs_Hyperlink = false;
            HyperlinkText = "";
            string StockType = "";
            double StockX = 0;
            double StockY = 0;
            double StockZ = 0;
            double StockOD = 0;
            double StockID = 0;
            bool FirstTp = true;
            char Appos = (char)34;

            if (ActiveWorkplane == "#ERROR: Type error")
            {
                ActiveWorkplane = "world";
            }

            int Search_Toolpath = CurrentParameter.IndexOf("{toolpath.block.");
            int SearcNCProgram = CurrentParameter.IndexOf("{ncprogram.block.");
            if (Search_Toolpath >= 0)
            {
                ExtractStockInfo(Toolpath.Name, out StockType, out StockX, out StockY, out StockZ, out StockOD, out StockID);
            }
            else if (SearcNCProgram >= 0)
            {
                ExtractStockInfo(ToolpathData[0].Name, out StockType, out StockX, out StockY, out StockZ, out StockOD, out StockID);
            }

            if (CurrentParameter == "{project.path}")
            {
                Value = Project.Path;
            }
            else if (CurrentParameter == "{project.name}")
            {
                Value = Project.Name;
            }
            else if (CurrentParameter == "{project.ordernumber}")
            {
                Value = Project.OrderNumber;
            }
            else if (CurrentParameter == "{project.programmer}")
            {
                Value = Project.Programmer;
            }
            else if (CurrentParameter == "{project.partname}")
            {
                Value = Project.PartName;
            }
            else if (CurrentParameter == "{project.customer}")
            {
                Value = Project.Customer;
            }
            else if (CurrentParameter == "{project.date}")
            {
                Value = Project.Date;
            }
            else if (CurrentParameter == "{project.notes}")
            {
                Value = Project.Notes;
            }
            else if (CurrentParameter == "{project.totaltime}")
            {
                ConvertMinutesToTime(Project.TotalTime, out Value);
            }
            else if (CurrentParameter == "{project.machmodels.maxx}")
            {
                Value = Project.MachModelsMaxX;
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{project.machmodels.maxy}")
            {
                Value = Project.MachModelsMaxY;
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{project.machmodels.maxz}")
            {
                Value = Project.MachModelsMaxZ;
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{project.machmodels.minx}")
            {
                Value = Project.MachModelsMinX;
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{project.machmodels.miny}")
            {
                Value = Project.MachModelsMinY;
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{project.machmodels.minz}")
            {
                Value = Project.MachModelsMinZ;
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{project.modelslist}")
            {
                Value = Project.ModelsList;
            }
            else if (CurrentParameter == "{project.combinedlist}")
            {
                Value = Project.CombinedNCTPList;
                Needs_Hyperlink = true;
            }
            else if (CurrentParameter == "{project.backtoprojectsummaryfull}")
            {
                Value = "'Project_Summary_Full'!A1";
                Needs_Hyperlink = true;
                HyperlinkText = "Back to Project Summary Full";
            }
            else if (CurrentParameter == "{ncprogram.name}")
            {
                Value = PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").name");
            }
            else if (CurrentParameter == "{ncprogram.outputworkplane}")
            {
                if (PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").outputworkplane.name") == "#ERROR: Invalid name")
                {
                    Value = "world";
                }
                else
                {
                    Value = PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").outputworkplane.name");
                }
            }
            else if (CurrentParameter == "{ncprogram.totaltime}")
            {
                double TotalTime = double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").statistics.totaltime"));
                double Hours = Math.Round(TotalTime / 60, 0);
                if ((Hours * 60) > TotalTime)
                {
                    Hours = Hours - 1;
                }
                double Minutes = Math.Round(TotalTime - (Hours * 60));
                if (Minutes + (60 * Hours) > TotalTime)
                {
                    Minutes = Minutes - 1;
                }
                double Seconds = Math.Round(((TotalTime - (Hours * 60) - Minutes) * 60), 0);
                Value = Hours + ":" + Minutes + ":" + Seconds;
            }
            else if (CurrentParameter == "{ncprogram.cuttinglength}")
            {
                Value = (double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").Statistics.CuttingMoves.Lengths.Linear")) + double.Parse(PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").Statistics.CuttingMoves.Lengths.Arcs"))).ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{ncprogram.notes}")
            {
                Value = PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").notes");
            }
            else if (CurrentParameter == "{ncprogram.filename}")
            {
                Value = PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").filename");
            }
            else if (CurrentParameter == "{ncprogram.optionfile.name}")
            {
                string OptionFileName = PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").optionfile.path");
                OptionFileName = OptionFileName.Split('\\').Last();
                Value = OptionFileName;
            }
            else if (CurrentParameter == "{ncprogram.optionfile.fullname}")
            {
                Value = PowerMILLAutomation.ExecuteEx("print $entity('ncprogram';" + Appos + NCProg + Appos + ").optionfile.path");
            }
            else if (CurrentParameter == "{ncprogram.backtoprojectsummary}")
            {
                Value = "'Project_Summary'!A1";
                HyperlinkText = "Back to Project Summary";
                Needs_Hyperlink = true;
            }
            else if (CurrentParameter == "{ncprogram.block.type}" || CurrentParameter == "{toolpath.block.type}")
            {
                Value = StockType;
            }
            else if (CurrentParameter == "{ncprogram.block.xsize}" || CurrentParameter == "{toolpath.block.xsize}")
            {
                if (StockType == "Box")
                {
                    Value = Math.Round(StockX, 4).ToString();
                }
                else if (StockType == "Cylinder")
                {
                    Value = Math.Round(StockOD, 4).ToString();
                }
                Value = Value.Replace(",", ".");

            }
            else if (CurrentParameter == "{ncprogram.block.ysize}" || CurrentParameter == "{toolpath.block.ysize}")
            {
                if (StockType == "Box")
                {
                    Value = Math.Round(StockY, 4).ToString();
                }
                else if (StockType == "Cylinder")
                {
                    Value = Math.Round(StockID, 4).ToString();
                }
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{ncprogram.block.zsize}" || CurrentParameter == "{toolpath.block.zsize}")
            {
                Value = Math.Round(StockZ, 4).ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{ncprogram.toolpaths}")
            {
                foreach (ToolpathInfo Tp in ToolpathData)
                {
                    if (FirstTp)
                    {
                        Value = Tp.Name;
                        FirstTp = false;
                    }
                    else
                    {
                        Value = Value + Environment.NewLine + Tp.Name;
                    }
                    
                }
                Needs_Hyperlink = true;
            }
            else if (CurrentParameter == "{workplane.vise}")
            {
                ExtractViseInfo(ActiveWorkplane, out string ViseName, out string ParaDimensions);

                Value = ViseName;
            }
            else if (CurrentParameter == "{workplane.parrallels}")
            {
                ExtractViseInfo(ActiveWorkplane, out string ViseName, out string ParaDimensions);
                Value = ParaDimensions;
            }
            else if (CurrentParameter == "{toolpath.tool.number}")
            {
                Value = Toolpath.ToolNumber.ToString();
            }
            else if (CurrentParameter == "{toolpath.type}")
            {
                Value = Toolpath.ToolpathType;
            }
            else if (CurrentParameter == "{toolpath.tool.name}")
            {
                Value = Toolpath.ToolName;
            }
            else if (CurrentParameter == "{toolpath.name}")
            {
                if (Is_Hyperlink)
                {
                    string SheetName = NCProg + "_" + Toolpath.Name;
                    ConvertToUsableName(SheetName, out SheetName);

                    Value = "'" + SheetName + "'!A1";
                    Needs_Hyperlink = true;
                    HyperlinkText = Toolpath.Name;
                }
                else
                {
                    Value = Toolpath.Name;
                }

            }
            else if (CurrentParameter == "{toolpath.description}")
            {
                Value = Toolpath.Description;
            }
            else if (CurrentParameter == "{toolpath.notes}")
            {
                Value = Toolpath.Notes;
            }
            else if (CurrentParameter == "{toolpath.backtoncprogsummary}")
            {
                Value = "'" + oSheet.Name + "'!A1";
                Needs_Hyperlink = true;
                HyperlinkText = "Back to NCProg Summary";
            }
            else if (CurrentParameter == "{toolpath.tool.description}")
            {
                Value = Toolpath.ToolDescription;
            }
            else if (CurrentParameter == "{toolpath.sfm}")
            {
                Value = Toolpath.SFM.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.ipt}")
            {
                Value = Toolpath.IPT.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.plungefeed}")
            {
                Value = Toolpath.PlungeFeed.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.skimfeed}")
            {
                Value = Toolpath.SkimFeed.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.feed}")
            {
                Value = Toolpath.Feed.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.rpm}")
            {
                Value = Toolpath.Speed.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.coolant}")
            {
                Value = Toolpath.Coolant;
            }
            else if (CurrentParameter == "{toolpath.tool.diameter}")
            {
                Value = Toolpath.ToolDiameter.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.tool.length}")
            {
                Value = Toolpath.ToolCutterLength.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.tool.overhang}")
            {
                Value = Toolpath.ToolOverhang.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.tool.gaugelength}")
            {
                Value = Toolpath.ToolGaugeLength.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.tool.flutes}")
            {
                Value = Toolpath.ToolNumberOfFlutes.ToString();
            }
            else if (CurrentParameter == "{toolpath.tool.holder}")
            {
                Value = Toolpath.ToolHolderName;
            }
            else if (CurrentParameter == "{toolpath.tool.tipradius}")
            {
                Value = Toolpath.ToolTipRadius.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.tool.type}")
            {
                Value = Toolpath.ToolType;
            }
            else if (CurrentParameter == "{toolpath.AxialRadialthickness}")
            {
                Value = Toolpath.Thickness + "/" + Toolpath.AxialThickness;
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.Axialthickness}")
            {
                Value = Toolpath.AxialThickness;
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.Radialthickness}")
            {
                Value = Toolpath.Thickness;
                if (Value == "")
                {
                    Value = Toolpath.AxialThickness;
                }
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.tool.lengthoffset}")
            {
                Value = Toolpath.ToolLengthOffset.ToString();
            }
            else if (CurrentParameter == "{toolpath.tool.diameteroffset}")
            {
                Value = Toolpath.ToolRadOffset.ToString();
            }
            else if (CurrentParameter == "{toolpath.cuttercomp}")
            {
                Value = Toolpath.CutterComp;
            }
            else if (CurrentParameter == "{toolpath.totaltime}")
            {
                ConvertMinutesToTime(Toolpath.Statistic_Time, out Value);
            }
            else if (CurrentParameter == "{toolpath.cuttinglength}")
            {
                Value = Toolpath.CuttingDistance.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.stepover}")
            {
                Value = Toolpath.Stepover.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.stepdown}")
            {
                Value = Toolpath.DOC.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.workplane}")
            {
                Value = Toolpath.TPWorkplane;
            }
            else if (CurrentParameter == "{toolpath.axistype}")
            {
                Value = Toolpath.GeneralAxisType;
            }
            else if (CurrentParameter == "{toolpath.tolerance}")
            {
                Value = Toolpath.Tolerance.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.rapidheight}")
            {
                Value = Toolpath.RapidHeight.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.skimheight}")
            {
                Value = Toolpath.SkimHeight.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.minx}")
            {
                Value = Toolpath.ToolpathMinX.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.miny}")
            {
                Value = Toolpath.ToolpathMinY.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.minz}")
            {
                Value = Toolpath.ToolpathMinZ.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.maxx}")
            {
                Value = Toolpath.ToolpathMaxX.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.maxy}")
            {
                Value = Toolpath.ToolpathMaxY.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.maxz}")
            {
                Value = Toolpath.ToolpathMaxZ.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{toolpath.leads}")
            {
                Value = Toolpath.FirstLeadInType + "/" + Toolpath.FirstLeadOutType;
            }
            else if (CurrentParameter == "{toolpath.ncprogram}")
            {
                Value = Toolpath.NCProgName;
            }
            else if (CurrentParameter == "{tool.number}")
            {
                Value = Tool.Number.ToString();
            }
            else if (CurrentParameter == "{tool.name}")
            {
                Value = Tool.Name;
            }
            else if (CurrentParameter == "{tool.diameter}")
            {
                Value = Tool.Diameter.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{tool.length}")
            {
                Value = Tool.CutterLength.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{tool.overhang}")
            {
                Value = Tool.ToolOverhang.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{tool.gaugelength}")
            {
                Value = Tool.ToolGaugeLength.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{tool.flutes}")
            {
                Value = Tool.NumberOfFlutes.ToString();
            }
            else if (CurrentParameter == "{tool.tiprad}")
            {
                Value = Tool.ToolTipRadius.ToString();
                Value = Value.Replace(",", ".");
            }
            else if (CurrentParameter == "{tool.type}")
            {
                Value = Tool.ToolType;
            }
            else if (CurrentParameter == "{tool.holdername}")
            {
                Value = Tool.HolderName;
            }
            else if (CurrentParameter == "{tool.lengthoffsetnumber}")
            {
                Value = Tool.ToolLengthOffset.ToString();
            }
            else if (CurrentParameter == "{tool.radiusoffsetnumber}")
            {
                Value = Tool.ToolRadOffset.ToString();
            }
            else if (CurrentParameter == "{tool.description}")
            {
                Value = Tool.Description;
            }
            else if (CurrentParameter == "{tool.image}")
            {
                ConvertToUsableName(Tool.Name + ".png", out Value);
                if (File.Exists(Project_Path + "\\Excel_Setupsheet\\" + Value))
                {
                    PowerMILLAutomation.ExecuteEx("FORM RAISE TOOLASSPREVIEWFORM '" + Tool.Name + "'");
                    PowerMILLAutomation.ExecuteEx("WAIT 1");
                    PowerMILLAutomation.ExecuteEx("EXPORT TOOLASSEMBLYPREVIEW FILESAVE '" + Project_Path + "\\Excel_Setupsheet\\" + Value + "'");
                    PowerMILLAutomation.ExecuteEx("FORM CANCEL TOOLASSPREVIEWFORM");
                    PowerMILLAutomation.ExecuteEx("TOOL ACCEPT");
                }
            }
            else if (CurrentParameter == "{toolpath.tool.image}")
            {
                ConvertToUsableName(Toolpath.ToolName + ".png", out Value);
                if (!File.Exists(Project_Path + "\\Excel_Setupsheet\\" + Value))
                {
                    PowerMILLAutomation.ExecuteEx("FORM RAISE TOOLASSPREVIEWFORM '" + Toolpath.ToolName + "'");
                    PowerMILLAutomation.ExecuteEx("WAIT 1");
                    PowerMILLAutomation.ExecuteEx("EXPORT TOOLASSEMBLYPREVIEW FILESAVE '" + Project_Path + "\\Excel_Setupsheet\\" + Value + "'");
                    PowerMILLAutomation.ExecuteEx("FORM CANCEL TOOLASSPREVIEWFORM");
                    PowerMILLAutomation.ExecuteEx("TOOL ACCEPT");
                }
            }
        }

        private static void ConvertToHyperlink(string Parameters, _Worksheet oSheet, int Row, int Column, string HyperlinkText, string NCProg, bool Is_NcProgList)
        {
            double TempVal = 0;
            float Left = 0;
            float Top = 0;
            float Height = 0;
            float Width = 0;
            string SheetName = "";
            string NCProgName = "";

            String[] lines = Parameters.Split(new String[] { Environment.NewLine }, StringSplitOptions.None);
            if (lines.Count() > 1)
            {
                oSheet.Cells[Row, Column] = Parameters;
                TempVal = oSheet.Cells[Row, Column].MergeArea.Top();
                Top = (float)TempVal;

                TempVal = oSheet.Cells[Row, Column].MergeArea.Left();
                Left = (float)TempVal;

                TempVal = oSheet.Cells[Row, Column].Font.Size + 4;
                Height = (float)TempVal;

                TempVal = oSheet.Cells[Row, Column].MergeArea.Width() - 1;
                Width = (float)TempVal;

                foreach (String line in lines)
                {
                    var Shape = oSheet.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, Left, Top, Width, Height);
                    Shape.Fill.Visible = MsoTriState.msoFalse;
                    Shape.Line.Visible = MsoTriState.msoFalse;

                    SheetName = line;
                    ConvertToUsableName(SheetName, out SheetName);

                    if (Is_NcProgList)
                    {
                        SheetName = NCProg + "_" + SheetName;
                    }
                    else
                    {
                        int Search_Toolpaths = SheetName.IndexOf("    ");
                        if (Search_Toolpaths == -1)
                        {
                            NCProgName = SheetName;
                        }
                        else
                        {
                            SheetName = NCProgName + "_" + SheetName.Substring(4, SheetName.Length - 4);
                        }
                    }


                    if (SheetName.Length > 31)
                    {
                        SheetName = SheetName.Substring(SheetName.Length - 31, 31);
                    }

                    oSheet.Hyperlinks.Add(Shape, "", "'" + SheetName + "'!A1", SheetName, SheetName);
                    //Shape.Hyperlink.Address = "Project_Summary";

                    Top = Top + Height;
                }
            }
            else
            {
                oSheet.Hyperlinks.Add(oSheet.Cells[Row, Column], string.Empty, Parameters, string.Empty, HyperlinkText);
            }
        }
        public static void CreateExcelFile(string NCProg, List<ToolpathInfo> ToolpathData, List<ToolInfo> ToolData, ProjectInfo Project, _Worksheet oSheet, Dictionary<string, string> VarsList, _Workbook Workbook, bool NeedsSummaryHyperlink, string Project_Path, _Worksheet oNCProgSummary, string ProjectTPList)
        {
            string FirstTPLine;
            string TemplateLine;
            string CurrentParameter = "";
            int Row;
            int CurrentRow = 0;
            int Column;
            int i=0;
            bool FirstCell = true;
            string PictureName = "";
            _Worksheet OriginalSheet = oSheet;
            string SheetName = "";
            IEnumerable<string> fullMatchingValues = null;
            float Left = 0;
            float Top = 0;
            float Height = 0;
            float Width = 0;
            double TempVal = 0;
            string Value = "";
            string HyperlinkText = "";
            bool Needs_Hyperlink = false;
            ToolpathInfo ToolpathEmpty = null;
            ToolInfo ToolEmpty = null;
            string CurrentNCProgName = "";
            bool NCProgramTitle = false;


            string ActiveWorkplane = PowerMILLAutomation.ExecuteEx("print $entity('workplane';'').name");

            fullMatchingValues = VarsList.Values.Where(CurrentValue => CurrentValue.Contains("{"));
            foreach (string currentValue in fullMatchingValues)
            {
                TemplateLine = KeyByValue(VarsList, currentValue);
                Int32.TryParse(TemplateLine.Substring(0, TemplateLine.IndexOf(",")), out Column);
                Int32.TryParse(TemplateLine.Substring(TemplateLine.IndexOf(",") + 1, TemplateLine.Length - TemplateLine.IndexOf(",") - 1), out Row);
                CurrentParameter = currentValue.Split('-').First(); ;

               // string ActiveWorkplane = PowerMILLAutomation.ExecuteEx("print $entity('workplane';'').name");

                int Search_Project = CurrentParameter.IndexOf("{project.");
                int Search_Workplane = CurrentParameter.IndexOf("{workplane.");
                int Search_NCProgram = CurrentParameter.IndexOf("{ncprogram.");

                if (Search_Project >= 0 || Search_Workplane >= 0 || Search_NCProgram >= 0)
                {
                    if (CurrentParameter == "{ncprogram.picture1}")
                    {
                        PictureName = "NCP-" + NCProg + "-1";
                        ImportImage(oSheet, PictureName, Project_Path, Row, Column);
                    }
                    else if (CurrentParameter == "{ncprogram.picture2}")
                    {
                        PictureName = "NCP-" + NCProg + "-2";
                        ImportImage(oSheet, PictureName, Project_Path, Row, Column);
                    }
                    else if (CurrentParameter == "{ncprogram.picture3}")
                    {
                        PictureName = "NCP-" + NCProg + "-3";
                        ImportImage(oSheet, PictureName, Project_Path, Row, Column);
                    }
                    else if (CurrentParameter == "{ncprogram.picture4}")
                    {
                        PictureName = "NCP-" + NCProg + "-4";
                        ImportImage(oSheet, PictureName, Project_Path, Row, Column);
                    }
                    else if (CurrentParameter == "{ncprogram.picture5}")
                    {
                        PictureName = "NCP-" + NCProg + "-5";
                        ImportImage(oSheet, PictureName, Project_Path, Row, Column);
                    }
                    else if (CurrentParameter == "{project.picture1}")
                    {
                        PictureName = "PRJ-1";
                        ImportImage(oSheet, PictureName, Project_Path, Row, Column);
                    }
                    else if (CurrentParameter == "{project.picture2}")
                    {
                        PictureName = "PRJ-2";
                        ImportImage(oSheet, PictureName, Project_Path, Row, Column);
                    }
                    else if (CurrentParameter == "{project.picture3}")
                    {
                        PictureName = "PRJ-3";
                        ImportImage(oSheet, PictureName, Project_Path, Row, Column);
                    }
                    else if (CurrentParameter == "{project.picture4}")
                    {
                        PictureName = "PRJ-4";
                        ImportImage(oSheet, PictureName, Project_Path, Row, Column);
                    }
                    else if (CurrentParameter == "{project.picture5}")
                    {
                        PictureName = "PRJ-5";
                        ImportImage(oSheet, PictureName, Project_Path, Row, Column);
                    }
                    else
                    {
                        MatchParameter(CurrentParameter, Project_Path, NCProg, ToolpathEmpty, ToolEmpty, Project, ToolpathData, NeedsSummaryHyperlink, oNCProgSummary, out Value, out HyperlinkText, out Needs_Hyperlink, ActiveWorkplane);
                        if (Needs_Hyperlink)
                        {
                            if (Search_NCProgram >= 0)
                            {
                                ConvertToHyperlink(Value, oSheet, Row, Column, HyperlinkText, NCProg, true);
                            }
                            else
                            {
                                ConvertToHyperlink(Value, oSheet, Row, Column, HyperlinkText, NCProg, false);
                            }

                        }
                        else
                        {
                            oSheet.Cells[Row, Column] = Value;
                            
                        }
                    }
                }
            }

            if (OriginalSheet.Name != "Project_Summary")
            {
                foreach (ToolpathInfo Toolpath in ToolpathData)
                {
                    SheetName = NCProg + "_" + Toolpath.Name;
                    ConvertToUsableName(SheetName, out SheetName);

                    if (OriginalSheet.Name == "Toolpath_Details")
                    {
                        int Index = OriginalSheet.Index;
                        OriginalSheet.Copy(OriginalSheet, Type.Missing);
                        Workbook.Sheets[Index].Name = SheetName;
                        oSheet = Workbook.Sheets[Index];
                    }

                    fullMatchingValues = VarsList.Values.Where(CurrentValue => CurrentValue.Contains("{toolpath."));
                    
                    FirstTPLine = "";
                    foreach (string currentValue in fullMatchingValues)
                    {
                        FirstTPLine = KeyByValue(VarsList, currentValue);
                        break;
                    }

                    if (OriginalSheet.Name == "Toolpath_Details")
                    {
                        CurrentRow = 0;
                    }

                    FirstCell = true;
                    NCProgramTitle = false;
                    foreach (string currentValue in fullMatchingValues)
                    {
                        FirstTPLine = KeyByValue(VarsList, currentValue);
                        Int32.TryParse(FirstTPLine.Substring(0, FirstTPLine.IndexOf(",")), out Column);
                        Int32.TryParse(FirstTPLine.Substring(FirstTPLine.IndexOf(",") + 1, FirstTPLine.Length - FirstTPLine.IndexOf(",") - 1), out Row);
                        Row = CurrentRow + Row;

                        CurrentParameter = currentValue.Split('-').First();

                        if (OriginalSheet.Name != "Toolpath_Details")
                        {

                            //Copy current excel file line until it's the last toolpath
                            if (i < ToolpathData.Count - 1 && FirstCell)
                            {
                                Range newRange = null;
                                Range RowRange = null;
                                //Copy the ncprogram name AND toolpath detail line if it's a full project summary page
                                if (OriginalSheet.Name == "Project_Summary_Full")
                                {
                                    if (CurrentParameter == "{toolpath.ncprogram}")
                                    {
                                        newRange = oSheet.Rows[Row+2];
                                        RowRange = oSheet.Rows.Range[oSheet.Cells[Row, 1], oSheet.Cells[Row + 1, 100]];
                                        newRange.Insert(XlInsertShiftDirection.xlShiftDown);
                                        newRange.Insert(XlInsertShiftDirection.xlShiftDown);
                                        RowRange.Copy();
                                        newRange = (Range)oSheet.Rows[newRange.Row - 2];
                                        NCProgramTitle = true;
                                    }
                                    else
                                    {
                                        newRange = oSheet.Rows[Row];
                                    }
                                }
                                else
                                {
                                    newRange = oSheet.Rows[Row];
                                    RowRange = oSheet.Rows[Row];
                                    newRange.Insert(XlInsertShiftDirection.xlShiftDown);
                                    RowRange.Copy();
                                    newRange = (Range)oSheet.Rows[newRange.Row - 1];
                                }


                                newRange.PasteSpecial();
                                newRange.RowHeight = RowRange.RowHeight;
                                FirstCell = false;
                            }
                        }

                        if (CurrentParameter == "{toolpath.picture}")
                        {
                            PictureName = "TP-" + Toolpath.Name;
                            ImportImage(oSheet, PictureName, Project_Path, Row, Column);
                        }
                        else
                        {
                            //Match current parameter with something from the dictionnary
                            MatchParameter(CurrentParameter, Project_Path, NCProg, Toolpath, ToolEmpty, Project, ToolpathData, NeedsSummaryHyperlink, oNCProgSummary, out Value, out HyperlinkText, out Needs_Hyperlink, ActiveWorkplane);
                            if (OriginalSheet.Name == "Project_Summary_Full" && CurrentParameter == "{toolpath.ncprogram}")
                            {
                                NCProg = Value;
                            }
                            if (Needs_Hyperlink)
                            {
                                ConvertToHyperlink(Value, oSheet, Row, Column, HyperlinkText, NCProg, false);
                            }
                            else
                            {
                                if (OriginalSheet.Name == "Project_Summary_Full" && CurrentParameter == "{toolpath.ncprogram}")
                                {
                                    //Delete previous line if the ncprogram name is the same as before
                                    if (Toolpath.NCProgName != CurrentNCProgName)
                                    {
                                        oSheet.Cells[Row, Column] = Value;
                                        CurrentNCProgName = Toolpath.NCProgName;
                                    }
                                    else
                                    {
                                        Range newRange = oSheet.Rows[Row];
                                        newRange.Delete();
                                        CurrentRow = CurrentRow - 1;
                                    }
                                }
                                else
                                {
                                    if (CurrentParameter == "{toolpath.tool.image}" && Value.IndexOf(".png") > 0)
                                    {
                                        ImportImage(oSheet, Value.Replace(".png", ""), Project_Path, Row, Column);
                                    }
                                    else
                                    {
                                        oSheet.Cells[Row, Column] = Value;
                                    }
                                }
                            }
                        }
                    }
                    i = i + 1;

                    //Add an extra line for Ncprogram name for full project summary
                    if (NCProgramTitle)
                    {
                        CurrentRow = CurrentRow + 2;
                    }
                    else
                    {
                        CurrentRow = CurrentRow + 1;
                    }

                    if (OriginalSheet.Name != "Toolpath_Details")
                    {
                        //break;
                    }
                }


                //tools infos
                fullMatchingValues = VarsList.Values.Where(CurrentValue => CurrentValue.Contains("{tool."));

                FirstTPLine = "";
                foreach (string currentValue in fullMatchingValues)
                { 
                    FirstTPLine = KeyByValue(VarsList, currentValue);
                    break;
                }

                CurrentRow = 0;
                i = 0;
                foreach (var Tool in ToolData)
                { 
                    FirstCell = true;
                    foreach (string currentValue in fullMatchingValues)
                    {
                        FirstTPLine = KeyByValue(VarsList, currentValue);
                        Int32.TryParse(FirstTPLine.Substring(0, FirstTPLine.IndexOf(",")), out Column);
                        Int32.TryParse(FirstTPLine.Substring(FirstTPLine.IndexOf(",") + 1, FirstTPLine.Length - FirstTPLine.IndexOf(",") - 1), out Row);
                        Row = CurrentRow + Row;

                        if (i < ToolData.Count - 1 && FirstCell)
                        {
                            Range newRange = oSheet.Rows[Row];
                            Range RowRange = oSheet.Rows[Row];
                            newRange.Insert(XlInsertShiftDirection.xlShiftDown);
                            RowRange.Copy();
                            newRange = (Range)oSheet.Rows[newRange.Row - 1];
                            newRange.PasteSpecial();
                            newRange.RowHeight = RowRange.RowHeight;
                            FirstCell = false;
                        }
                        CurrentParameter = currentValue.Split('-').First(); ;

                        MatchParameter(CurrentParameter, Project_Path, NCProg, ToolpathEmpty, Tool, Project,ToolpathData, NeedsSummaryHyperlink, oNCProgSummary, out Value, out HyperlinkText, out Needs_Hyperlink, ActiveWorkplane);

                        if (CurrentParameter == "{toolpath.tool.image}" && Value.IndexOf(".png") > 0)
                        {
                            ImportImage(oSheet, Value.Replace(".png", ""), Project_Path, Row, Column);
                        }
                        else
                        {
                            oSheet.Cells[Row, Column] = Value;
                        }

                    }
                    i = i + 1;
                    CurrentRow = CurrentRow + 1;
                }
            }
            if (oSheet.Name == "Project_Summary_Full" || oSheet.Name == "Project_Summary")
            {
                oSheet.Activate();
                var range = oSheet.Range["A6"];
                range.Select();
            }
        }

        public static void ConvertMinutesToTime(double TotalTime, out string Time)
        {
            double Hours = Math.Round(TotalTime / 60, 0);
            if ((Hours * 60) > TotalTime)
            {
                Hours = Hours - 1;
            }
            double Minutes = Math.Round(TotalTime - (Hours * 60));
            if (Minutes + (60 * Hours) > TotalTime)
            {
                Minutes = Minutes - 1;
            }
            double Seconds = Math.Round(((TotalTime - (Hours * 60) - Minutes) * 60), 0);
            Time = Hours + ":" + Minutes + ":" + Seconds;
        }
    }
}
