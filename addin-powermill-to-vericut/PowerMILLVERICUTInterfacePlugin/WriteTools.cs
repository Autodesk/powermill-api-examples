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
using System.IO;
using Delcam.Plugins.Utils;

using PowerMILLExporter;
using PowerMILLExporter.Tools;


namespace PowerMILLVERICUTInterfacePlugin
{
    public static class WriteTools
    {
        /// <summary>
        /// Gets list of tools missing from the library
        /// </summary>
        /// <param name="tools"></param>
        /// <param name="library_tools"></param>
        /// <returns></returns>
        public static List<string> FindToolsMissingFromLibrary(Tools tools, List<string> library_tools)
        {
            List<string> missing_tools = new List<string>();

            if (tools.o2DTools != null)
            {
                foreach (_2DTool tool in tools.o2DTools)
                    if (!library_tools.Contains(tool.Name))
                        missing_tools.Add(tool.Name);
            }

            if (tools.oATP5MillTools != null)
            {
                foreach (ATP5Mill tool in tools.oATP5MillTools)
                    if (!library_tools.Contains(tool.Name))
                        missing_tools.Add(tool.Name);
            }

            if (tools.oDrillingTools != null)
            {
                foreach (DrillingTool tool in tools.oDrillingTools)
                    if (!library_tools.Contains(tool.Name))
                        missing_tools.Add(tool.Name);
            }

            if (tools.oMillingTools != null)
            {
                foreach (MillingTool tool in tools.oMillingTools)
                    if (!library_tools.Contains(tool.Name))
                        missing_tools.Add(tool.Name);
            }

            if (tools.oSpecialMILLTools != null)
            {
                foreach (SpecialMill tool in tools.oSpecialMILLTools)
                    if (!library_tools.Contains(tool.Name))
                        missing_tools.Add(tool.Name);
            }

            if (tools.oTaperSphericalTools != null)
            {
                foreach (TaperSpherical tool in tools.oTaperSphericalTools)
                    if (!library_tools.Contains(tool.Name))
                        missing_tools.Add(tool.Name);
            }

            if (tools.oTaperTippedTools != null)
            {
                foreach (TaperedTipped tool in tools.oTaperTippedTools)
                    if (!library_tools.Contains(tool.Name))
                        missing_tools.Add(tool.Name);
            }

            if (tools.oTapTools != null)
            {
                foreach (TapTool tool in tools.oTapTools)
                    if (!library_tools.Contains(tool.Name))
                        missing_tools.Add(tool.Name);
            }

            if (tools.oTopAndSideMillTools != null)
            {
                foreach (TopAndSideMill tool in tools.oTopAndSideMillTools)
                    if (!library_tools.Contains(tool.Name))
                        missing_tools.Add(tool.Name);
            }

            return missing_tools;
        }

        ///<summary>
        ///Append tools to library
        ///</summary>
        ///<param name="oStreamWriter"></param>
        ///<remarks></remarks>
        public static void AppendToolsToXML(PowerMILLExporter.Tools.Tools tools, string output_dir, string fname,
                                           List<string> missing_tools)
        {
            string fpath = Path.Combine(output_dir, fname);
            StringBuilder oStringBuilder = new StringBuilder();

            // Write MillingTools
            if (tools.oMillingTools != null)
            {
                foreach (MillingTool tool in tools.oMillingTools)
                    if (missing_tools.Contains(tool.Name))
                        WriteMillingTool(tool, oStringBuilder, output_dir, false, true);
            }


            // Write DrillingTools
            if (tools.oDrillingTools != null)
            {
                foreach (DrillingTool tool in tools.oDrillingTools)
                    if (missing_tools.Contains(tool.Name))
                        WriteDrillingTool(tool, oStringBuilder, output_dir, false, true);
            }

            // Write Tapered Tipped Tools
            if (tools.oTaperTippedTools != null)
            {
                foreach (TaperedTipped tool in tools.oTaperTippedTools)
                    if (missing_tools.Contains(tool.Name))
                        WriteTaperedTippedTool(tool, oStringBuilder, output_dir, false, true);
            }

            // Write Tapered Ball Tools
            if (tools.oTaperSphericalTools != null)
            {
                foreach (TaperSpherical tool in tools.oTaperSphericalTools)
                    if (missing_tools.Contains(tool.Name))
                        WriteTaperSphericalTool(tool, oStringBuilder, output_dir, false, true);
            }

            // Write TopAndSideMillTools
            if (tools.oTopAndSideMillTools != null)
            {
                foreach (TopAndSideMill tool in tools.oTopAndSideMillTools)
                    if (missing_tools.Contains(tool.Name))
                        WriteTopAndSideMillTool(tool, oStringBuilder, output_dir, false, true);
            }

            string fcontent = File.ReadAllText(fpath);
            string fcontent_start, fcontent_end;
            fcontent_start = fcontent.Substring(0, fcontent.IndexOf("</Tools>")).TrimEnd();
            fcontent_end = fcontent.Substring(fcontent.IndexOf("</Tools>"));
            File.WriteAllText(fpath,
                                fcontent_start + Environment.NewLine +
                                oStringBuilder.ToString() + 
                                fcontent_end);
        }

        ///<summary>
        ///Write tools
        ///</summary>
        ///<param name="oStreamWriter"></param>
        ///<remarks></remarks>
        public static void WriteToolsToXML(PowerMILLExporter.Tools.Tools tools, string output_dir, string fname, 
                                           bool tool_id_use_num, bool tool_id_use_name)
        {
            string fpath = Path.Combine(output_dir, fname);
            StringBuilder oStringBuilder = new StringBuilder();
            oStringBuilder.AppendLine("<?xml version=\"1.0\"?>");
            oStringBuilder.AppendLine("<CGTechToolLibrary>");
            oStringBuilder.AppendLine(Indent(1) + "<Tools>");

            // Write MillingTools
            if (tools.oMillingTools != null)
            {
                foreach (MillingTool tool in tools.oMillingTools)
                    WriteMillingTool(tool, oStringBuilder, output_dir, tool_id_use_num, tool_id_use_name);
            }


            // Write DrillingTools
            if (tools.oDrillingTools != null)
            {
                foreach (DrillingTool tool in tools.oDrillingTools)
                    WriteDrillingTool(tool, oStringBuilder, output_dir, tool_id_use_num, tool_id_use_name);
            }

            // Write Tapper Tipped tools
            if (tools.oTaperTippedTools != null)
            {
                foreach (TaperedTipped tool in tools.oTaperTippedTools)
                    WriteTaperedTippedTool(tool, oStringBuilder, output_dir, tool_id_use_num, tool_id_use_name);
            }

            // Write Tapper Spherical tools
            if (tools.oTaperSphericalTools != null)
            {
                foreach (TaperSpherical tool in tools.oTaperSphericalTools)
                    WriteTaperSphericalTool(tool, oStringBuilder, output_dir, tool_id_use_num, tool_id_use_name);
            }

            //    // Write TapTools

            if (tools.oTapTools != null)
            {
                foreach (TapTool tool in tools.oTapTools)
                    WriteTapTool(tool, oStringBuilder, output_dir, tool_id_use_num, tool_id_use_name);
            }

            if (tools.oSpecialMILLTools != null)
            {
                foreach (SpecialMill tool in tools.oSpecialMILLTools)
                    WriteSpecialMILLTool(tool, oStringBuilder, output_dir, tool_id_use_num, tool_id_use_name);
            }


            //    // Write ATP5MillTools
            //    if (MainModule.oATP5MillTools != null)
            //    {
            //        for (int i = 0; i <= MainModule.oATP5MillTools.Count - 1; i++)
            //        {
            //            oStreamWriter.WriteLine("BEGIN_TOOL");
            //            oStreamWriter.WriteLine("  |TOOL|" + MainModule.oATP5MillTools[i].Type + "|" + MainModule.oATP5MillTools[i].Name.Replace(Convert.ToChar(216), 'D').Replace(Convert.ToChar(248), 'D') + "| |" + MainModule.oATP5MillTools[i].ToolNumber + "|");
            //            oStreamWriter.WriteLine("  |MATRICE_OXY|0.0,0.0,0.0,1.0,0.0,0.0,0.0,1.0,0.0|");
            //            oStreamWriter.WriteLine("  |UNITS|" + MainModule.PmillUnitsSystem + "|");
            //            oStreamWriter.WriteLine("  |GAUGE|" + MainModule.oATP5MillTools[i].Gauge + "|");
            //            oStreamWriter.WriteLine("  |PARAMETER|" + MainModule.oATP5MillTools[i].Diameter + "|" + MainModule.oATP5MillTools[i].Length + "|" + MainModule.oATP5MillTools[i].CornerRadius + "|" + MainModule.oATP5MillTools[i].OverLength + "|" + MainModule.oATP5MillTools[i].CuttingLength + "|" + "|" + MainModule.oATP5MillTools[i].NoseAngle + "|" + MainModule.oATP5MillTools[i].SideAngle + "|");
            //            oStreamWriter.WriteLine("  |CUTTING_PARAMETER|" + MainModule.oATP5MillTools[i].TeethNumber + "| | |");

            //            // Write ToolShank Part
            //            if (MainModule.oATP5MillTools[i].ToolShankPath != null)
            //            {
            //                oStreamWriter.WriteLine("BEGIN_EXTENSION");
            //                oStreamWriter.WriteLine("  |2DEF|" + System.IO.Path.GetFileName(MainModule.oATP5MillTools[i].ToolShankPath) + "|Y+|0|");
            //                oStreamWriter.WriteLine("END");
            //            }

            //            // Write ToolHolder Part
            //            if (MainModule.oATP5MillTools[i].ToolHolderPath != null)
            //            {
            //                oStreamWriter.WriteLine("BEGIN_TOOLHOLDERS");
            //                oStreamWriter.WriteLine("  |2DEF|" + System.IO.Path.GetFileName(MainModule.oATP5MillTools[i].ToolHolderPath) + "|Y+|0|");
            //                oStreamWriter.WriteLine("END");
            //            }
            //            oStreamWriter.WriteLine("END");
            //        }
            //    }

            // Write TopAndSideMillTools
            if (tools.oTopAndSideMillTools != null)
            {
                foreach (TopAndSideMill tool in tools.oTopAndSideMillTools)
                    WriteTopAndSideMillTool(tool, oStringBuilder, output_dir, tool_id_use_num, tool_id_use_name);
            }

            //    // Write 2DTools
            if (tools.o2DTools != null)
            {
                foreach (_2DTool tool in tools.o2DTools)
                    Write2DTool(tool, oStringBuilder, output_dir, tool_id_use_num, tool_id_use_name);
            }


            oStringBuilder.AppendLine(Indent(1) + "</Tools>");
            oStringBuilder.AppendLine("</CGTechToolLibrary>");
            File.WriteAllText(fpath, oStringBuilder.ToString());
        }

        /* Done */
        private static void WriteMillingTool(MillingTool oTool, StringBuilder oStringBuilder, string output_dir,
                                            bool tool_id_use_num, bool tool_id_use_name)
        {
            oStringBuilder.AppendLine(Indent(2) + String.Format("<Tool ID=\"{0}\" Units=\"{1}\">",
                                        ConstructToolName(oTool.ToolNumber, oTool.Name, tool_id_use_num, tool_id_use_name),
                                        (ProjectData.PmillUnitsIsInch ? "Inch" : "Millimeter")));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Description>{0}</Description>", oTool.Name));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Teeth>{0}</Teeth>", oTool.TeethNumber));
            oStringBuilder.AppendLine(Indent(3) + "<Type>Milling</Type>");
            oStringBuilder.AppendLine(Indent(3) + "<Cutter>");
            WriteShankSTL(ref oStringBuilder, oTool.ToolShankProfile, oTool.Name, false);
            string type = "";
            if (oTool.CornerRadius == 0)
                type = "FLAT END";
            else if (Math.Abs(oTool.CornerRadius - oTool.Diameter / 2) > 0.00001)
                type = "BULL NOSE";
            else
                type = "BALL END";
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Apt ID=\"{0}_Cutter\" ParentID=\"{0}_Holder\" Type=\"{1}\">",
                                        oTool.Name, type));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<D>{0}</D>", oTool.Diameter));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<R>{0}</R>", oTool.CornerRadius));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<E>{0}</E>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<F>{0}</F>", oTool.CornerRadius));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<A>{0}</A>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<B>{0}</B>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<H>{0}</H>", oTool.CuttingLength));
            //oStringBuilder.AppendLine(Indent(4) + String.Format("<StickoutLength>{0}</StickoutLength>", oTool.OverHang));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<R2>{0}</R2>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<E2>{0}</E2>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<F2>{0}</F2>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<SpindleDirection>CW</SpindleDirection>"));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<FluteLength>{0}</FluteLength>", oTool.CuttingLength));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<ShankDiameter>{0}</ShankDiameter>", oTool.Diameter));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Alternate>off</Alternate>"));
            oStringBuilder.AppendLine(Indent(4) + "</Apt>");
            oStringBuilder.AppendLine(Indent(3) + "</Cutter>");
					
            // Write ToolHolder Part
            WriteHolderSTL(ref oStringBuilder, oTool.ToolHolderProfile, oTool.Name);

	    oStringBuilder.AppendLine(Indent(4) + "<GagePoint>");
	    oStringBuilder.AppendLine(Indent(6) + "<X>0</X>");
	    oStringBuilder.AppendLine(Indent(6) + "<Y>0</Y>");
        oStringBuilder.AppendLine(Indent(6) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Z>{0}</Z>", oTool.Gauge));
	    oStringBuilder.AppendLine(Indent(4) + "</GagePoint>");
            oStringBuilder.AppendLine(Indent(2) + "</Tool>");
        }


        private static void Write2DTool(_2DTool oTool, StringBuilder oStringBuilder, string output_dir,
                                    bool tool_id_use_num, bool tool_id_use_name)
        {
            oStringBuilder.AppendLine(Indent(2) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Tool ID=\"{0}\" Units=\"{1}\">",
                                        ConstructToolName(oTool.ToolNumber, oTool.Name, tool_id_use_num, tool_id_use_name),
                                        (ProjectData.PmillUnitsIsInch ? "Inch" : "Millimeter")));
            oStringBuilder.AppendLine(Indent(3) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Description>{0}</Description>", oTool.Name));
            oStringBuilder.AppendLine(Indent(3) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Teeth>{0}</Teeth>", oTool.TeethNumber));
            oStringBuilder.AppendLine(Indent(3) + "<Type>Milling</Type>");
            oStringBuilder.AppendLine(Indent(3) + "<Cutter>");
            WriteShankSTL(ref oStringBuilder, oTool.ToolShankProfile, oTool.Name, false);
            string type = "";

            oStringBuilder.AppendLine(Indent(3) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Model ID=\"{0}_Cutter\" ParentID=\"{0}_Holder\" Type=\"STL\" Insert = \"yes\" DRILLMILL = \"FALSE\">", oTool.Name, type));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<FileName>{0}</FileName>", oTool.CADPath));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Normal>{0}</Normal>", "Outward"));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<NoSpin>{0}</NoSpin>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Alternate>{0}</Alternate>", "off"));
            oStringBuilder.AppendLine(Indent(4) + "</Model>");
            oStringBuilder.AppendLine(Indent(3) + "</Cutter>");

            // Write ToolHolder Part
            WriteHolderSTL(ref oStringBuilder, oTool.ToolHolderProfile, oTool.Name);

            oStringBuilder.AppendLine(Indent(4) + "<GagePoint>");
            oStringBuilder.AppendLine(Indent(6) + "<X>0</X>");
            oStringBuilder.AppendLine(Indent(6) + "<Y>0</Y>");
            oStringBuilder.AppendLine(Indent(6) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Z>{0}</Z>", oTool.Gauge));
            oStringBuilder.AppendLine(Indent(4) + "</GagePoint>");
            oStringBuilder.AppendLine(Indent(2) + "</Tool>");
        }


        private static void WriteSpecialMILLTool(SpecialMill oTool, StringBuilder oStringBuilder, string output_dir,
                            bool tool_id_use_num, bool tool_id_use_name)
        {
            oStringBuilder.AppendLine(Indent(2) + String.Format("<Tool ID=\"{0}\" Units=\"{1}\">",
                                       ConstructToolName(oTool.ToolNumber, oTool.Name, tool_id_use_num, tool_id_use_name),
                                       (ProjectData.PmillUnitsIsInch ? "Inch" : "Millimeter")));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Description>{0}</Description>", oTool.Name));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Teeth>{0}</Teeth>", oTool.TeethNumber));
            oStringBuilder.AppendLine(Indent(3) + "<Type>Milling</Type>");
            oStringBuilder.AppendLine(Indent(3) + "<Cutter>");
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Apt ID=\"{0}_Cutter\" ParentID=\"{0}_Holder\" Type=\"{1}\">",
                                        oTool.Name, "APT 7"));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<D>{0}</D>", oTool.Diameter));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<R>{0}</R>", oTool.TipRadius));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<E>{0}</E>", oTool.RadXOffset));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<F>{0}</F>", oTool.RadYOffset));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<A>{0}</A>", 0));//Tip Angle
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<B>{0}</B>", 0));//Side Angle
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<H>{0}</H>", oTool.CuttingLength));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<R2>{0}</R2>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<E2>{0}</E2>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<F2>{0}</F2>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<SpindleDirection>CW</SpindleDirection>"));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<FluteLength>{0}</FluteLength>", oTool.CuttingLength));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<ShankDiameter>{0}</ShankDiameter>", oTool.Diameter));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Alternate>off</Alternate>"));
            oStringBuilder.AppendLine(Indent(4) + "</Apt>");
            // Write Tool Shank Part
            WriteShankSTL(ref oStringBuilder, oTool.ToolShankProfile, oTool.Name, false);
            oStringBuilder.AppendLine(Indent(3) + "</Cutter>");

            // Write ToolHolder Part
            WriteHolderSTL(ref oStringBuilder, oTool.ToolHolderProfile, oTool.Name);

            oStringBuilder.AppendLine(Indent(4) + "<GagePoint>");
            oStringBuilder.AppendLine(Indent(6) + "<X>0</X>");
            oStringBuilder.AppendLine(Indent(6) + "<Y>0</Y>");
            oStringBuilder.AppendLine(Indent(6) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Z>{0}</Z>", oTool.Gauge));
            oStringBuilder.AppendLine(Indent(4) + "</GagePoint>");
            oStringBuilder.AppendLine(Indent(2) + "</Tool>");
        }


        private static void WriteDrillingTool(DrillingTool oTool, StringBuilder oStringBuilder, string output_dir,
                                            bool tool_id_use_num, bool tool_id_use_name)
        {
            oStringBuilder.AppendLine(Indent(2) + String.Format("<Tool ID=\"{0}\" Units=\"{1}\">",
                                        ConstructToolName(oTool.ToolNumber, oTool.Name, tool_id_use_num, tool_id_use_name),
                                        (ProjectData.PmillUnitsIsInch ? "Inch" : "Millimeter")));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Description>{0}</Description>", oTool.Name));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Teeth>{0}</Teeth>", oTool.TeethNumber));
            oStringBuilder.AppendLine(Indent(3) + "<Type>HoleMaking</Type>");
            oStringBuilder.AppendLine(Indent(3) + "<Cutter>");
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Apt ID=\"{0}_Cutter\" ParentID=\"{0}_Holder\" Type=\"{1}\">",
                                        oTool.Name, "DRILL"));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<D>{0}</D>", oTool.Diameter));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<R>{0}</R>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<E>{0}</E>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<F>{0}</F>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<A>{0}</A>", (90 - oTool.NoseAngle)));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<B>{0}</B>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<H>{0}</H>", oTool.CuttingLength));
            //oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<StickoutLength>{0}</StickoutLength>", oTool.OverHang));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<R2>{0}</R2>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<E2>{0}</E2>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<F2>{0}</F2>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format("<SpindleDirection>CW</SpindleDirection>"));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<FluteLength>{0}</FluteLength>", oTool.CuttingLength));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<ShankDiameter>{0}</ShankDiameter>", oTool.Diameter));
            oStringBuilder.AppendLine(Indent(4) + String.Format("<Alternate>off</Alternate>"));
            oStringBuilder.AppendLine(Indent(4) + "</Apt>");

            // Write Tool Shank Part
            WriteShankSTL(ref oStringBuilder, oTool.ToolShankProfile, oTool.Name, false);
            oStringBuilder.AppendLine(Indent(3) + "</Cutter>");

            // Write ToolHolder Part
            WriteHolderSTL(ref oStringBuilder, oTool.ToolHolderProfile, oTool.Name);

            oStringBuilder.AppendLine(Indent(4) + "<GagePoint>");
            oStringBuilder.AppendLine(Indent(6) + "<X>0</X>");
            oStringBuilder.AppendLine(Indent(6) + "<Y>0</Y>");
            oStringBuilder.AppendLine(Indent(6) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Z>{0}</Z>", oTool.Gauge));
            oStringBuilder.AppendLine(Indent(4) + "</GagePoint>");
            oStringBuilder.AppendLine(Indent(2) + "</Tool>");
        }


        private static void WriteTapTool(TapTool oTool, StringBuilder oStringBuilder, string output_dir,
                                    bool tool_id_use_num, bool tool_id_use_name)
        {
            oStringBuilder.AppendLine(Indent(2) + String.Format("<Tool ID=\"{0}\" Units=\"{1}\">",
                                        ConstructToolName(oTool.ToolNumber, oTool.Name, tool_id_use_num, tool_id_use_name),
                                        (ProjectData.PmillUnitsIsInch ? "Inch" : "Millimeter")));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Description>{0}</Description>", oTool.Name));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Teeth>{0}</Teeth>", oTool.TeethNumber));
            oStringBuilder.AppendLine(Indent(3) + "<Type>HoleMaking</Type>");
            oStringBuilder.AppendLine(Indent(3) + "<Cutter>");
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Apt ID=\"{0}_Cutter\" ParentID=\"{0}_Holder\" Type=\"{1}\">",
                                        oTool.Name, "TAP"));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<D>{0}</D>", oTool.Diameter));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<R>{0}</R>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<E>{0}</E>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<F>{0}</F>", 0));//Tip Angle Height
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<A>{0}</A>", 0));//Tip Angle
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<B>{0}</B>", 0));//Side Angle
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<H>{0}</H>", oTool.CuttingLength));
            //oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<StickoutLength>{0}</StickoutLength>", oTool.OverHang));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<R2>{0}</R2>", 1));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<E2>{0}</E2>", 2));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<F2>{0}</F2>", 3));
            oStringBuilder.AppendLine(Indent(4) + String.Format("<SpindleDirection>CW</SpindleDirection>"));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<FluteLength>{0}</FluteLength>", oTool.CuttingLength));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<ShankDiameter>{0}</ShankDiameter>", oTool.Diameter));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Alternate>off</Alternate>"));
            oStringBuilder.AppendLine(Indent(4) + "</Apt>");
            // Write Tool Shank Part
            WriteShankSTL(ref oStringBuilder, oTool.ToolShankProfile, oTool.Name, false);
            oStringBuilder.AppendLine(Indent(3) + "</Cutter>");

            // Write ToolHolder Part
            WriteHolderSTL(ref oStringBuilder, oTool.ToolHolderProfile, oTool.Name);

            oStringBuilder.AppendLine(Indent(4) + "<GagePoint>");
            oStringBuilder.AppendLine(Indent(6) + "<X>0</X>");
            oStringBuilder.AppendLine(Indent(6) + "<Y>0</Y>");
            oStringBuilder.AppendLine(Indent(6) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Z>{0}</Z>", oTool.Gauge));
            oStringBuilder.AppendLine(Indent(4) + "</GagePoint>");
            oStringBuilder.AppendLine(Indent(2) + "</Tool>");
        }




        /* Done */
        private static void WriteTaperedTippedTool(TaperedTipped oTool, StringBuilder oStringBuilder, string output_dir,
                                            bool tool_id_use_num, bool tool_id_use_name)
        {
            // Get the profile points
            double TaperAngleInRad = Utilities.DegreesToRadians(oTool.TaperAngle);
            double x3 = oTool.Diameter / 2;
            double y3 = oTool.TaperHeight;
            double R = oTool.TipRadius;
            double y2 = R * (1 - System.Math.Sin(TaperAngleInRad));
            double x2 = oTool.Diameter / 2 - (oTool.TaperHeight - y2) * System.Math.Tan(TaperAngleInRad);
            double x1 = x2 - R * System.Math.Cos(TaperAngleInRad);


            oStringBuilder.AppendLine(Indent(2) + String.Format("<Tool ID=\"{0}\" Units=\"{1}\">",
                                        ConstructToolName(oTool.ToolNumber, oTool.Name, tool_id_use_num, tool_id_use_name),
                                        (ProjectData.PmillUnitsIsInch ? "Inch" : "Millimeter")));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Description>{0}</Description>", oTool.Name));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Teeth>{0}</Teeth>", oTool.TeethNumber));
            oStringBuilder.AppendLine(Indent(3) + "<Type>Milling</Type>");
            oStringBuilder.AppendLine(Indent(3) + "<Cutter>");
            WriteShankSTL(ref oStringBuilder, oTool.ToolShankProfile, oTool.Name, false);
            oStringBuilder.AppendLine(Indent(4) + String.Format("<SOR ID=\"{0}_Cutter\" ParentID=\"{0}_Holder\" Type=\"{1}\">",
                            oTool.Name, "TAPERED BULL NOSE"));
            oStringBuilder.AppendLine(PointToXML(0, 0));
            oStringBuilder.AppendLine(PointToXML(x2 - R * System.Math.Cos(TaperAngleInRad), 0));
            oStringBuilder.AppendLine(ArcToXML(x2 - R * System.Math.Cos(TaperAngleInRad), oTool.TipRadius, oTool.TipRadius, true));
            oStringBuilder.AppendLine(PointToXML(x2, y2));
            oStringBuilder.AppendLine(PointToXML(oTool.Diameter / 2, oTool.TaperHeight));
            oStringBuilder.AppendLine(PointToXML(oTool.Diameter / 2, oTool.Length));
            oStringBuilder.AppendLine(PointToXML(0, oTool.Length));
            oStringBuilder.AppendLine(Indent(4) + "</SOR>");
            oStringBuilder.AppendLine(Indent(3) + "</Cutter>");

            // Write ToolHolder Part
            WriteHolderSTL(ref oStringBuilder, oTool.ToolHolderProfile, oTool.Name);

            oStringBuilder.AppendLine(Indent(4) + "<GagePoint>");
            oStringBuilder.AppendLine(Indent(6) + "<X>0</X>");
            oStringBuilder.AppendLine(Indent(6) + "<Y>0</Y>");
            oStringBuilder.AppendLine(Indent(6) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Z>{0}</Z>", oTool.Gauge));
            oStringBuilder.AppendLine(Indent(4) + "</GagePoint>");
            oStringBuilder.AppendLine(Indent(2) + "</Tool>");

        }


        /* Done */
        private static void WriteTaperSphericalTool(TaperSpherical oTool, StringBuilder oStringBuilder, string output_dir,
                                            bool tool_id_use_num, bool tool_id_use_name)
        {
            oStringBuilder.AppendLine(Indent(2) + String.Format("<Tool ID=\"{0}\" Units=\"{1}\">",
                                        ConstructToolName(oTool.ToolNumber, oTool.Name, tool_id_use_num, tool_id_use_name),
                                        (ProjectData.PmillUnitsIsInch ? "Inch" : "Millimeter")));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Description>{0}</Description>", oTool.Name));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Teeth>{0}</Teeth>", oTool.TeethNumber));
            oStringBuilder.AppendLine(Indent(3) + "<Type>Milling</Type>");
            oStringBuilder.AppendLine(Indent(3) + "<Cutter>");
            WriteShankSTL(ref oStringBuilder, oTool.ToolShankProfile, oTool.Name, false);
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Apt ID=\"{0}_Cutter\" ParentID=\"{0}_Holder\" Type=\"{1}\">",
                                        oTool.Name, "TAPERED BULL NOSE"));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<D>{0}</D>", oTool.Diameter));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<R>{0}</R>", oTool.CornerRadius));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<E>{0}</E>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<F>{0}</F>", oTool.CornerRadius));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<A>{0}</A>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<B>{0}</B>", oTool.NoseAngle));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<H>{0}</H>", oTool.CuttingLength));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<R2>{0}</R2>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<E2>{0}</E2>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<F2>{0}</F2>", 0));
            oStringBuilder.AppendLine(Indent(4) + String.Format("<SpindleDirection>CW</SpindleDirection>"));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<FluteLength>{0}</FluteLength>", oTool.CuttingLength));
            oStringBuilder.AppendLine(Indent(4) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<ShankDiameter>{0}</ShankDiameter>", oTool.Diameter));
            oStringBuilder.AppendLine(Indent(4) + String.Format("<Alternate>off</Alternate>"));
            oStringBuilder.AppendLine(Indent(4) + "</Apt>");
            oStringBuilder.AppendLine(Indent(3) + "</Cutter>");

            // Write ToolHolder Part
            WriteHolderSTL(ref oStringBuilder, oTool.ToolHolderProfile, oTool.Name);

            oStringBuilder.AppendLine(Indent(4) + "<GagePoint>");
            oStringBuilder.AppendLine(Indent(6) + "<X>0</X>");
            oStringBuilder.AppendLine(Indent(6) + "<Y>0</Y>");
            oStringBuilder.AppendLine(Indent(6) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Z>{0}</Z>", oTool.Gauge));
            oStringBuilder.AppendLine(Indent(4) + "</GagePoint>");
            oStringBuilder.AppendLine(Indent(2) + "</Tool>");
        }



        private static void WriteTopAndSideMillTool(TopAndSideMill oTool, StringBuilder oStringBuilder, string output_dir,
                                            bool tool_id_use_num, bool tool_id_use_name)
        {
            oStringBuilder.AppendLine(Indent(2) + String.Format("<Tool ID=\"{0}\" Units=\"{1}\">",
                                        ConstructToolName(oTool.ToolNumber, oTool.Name, tool_id_use_num, tool_id_use_name),
                                        (ProjectData.PmillUnitsIsInch ? "Inch" : "Millimeter")));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Description>{0}</Description>", oTool.Name));
            oStringBuilder.AppendLine(Indent(3) + String.Format("<Teeth>{0}</Teeth>", oTool.TeethNumber));
            oStringBuilder.AppendLine(Indent(3) + "<Type>Milling</Type>");
            oStringBuilder.AppendLine(Indent(3) + "<Cutter>");
            WriteShankSTL(ref oStringBuilder, oTool.ToolShankProfile, oTool.Name, false);
            oStringBuilder.AppendLine(Indent(4) + String.Format("<SOR ID=\"{0}_Cutter\" ParentID=\"{0}_Holder\" Type=\"{1}\">",
                oTool.Name, "TOP AND SIDE MILL"));
            oStringBuilder.AppendLine(PointToXML(0, 0));
            if (oTool.SideAngle == 0)
            {
                if (oTool.RadiusInf > 0)
                {
                    oStringBuilder.AppendLine(PointToXML(oTool.Diameter / 2 - oTool.RadiusInf, 0));
                    oStringBuilder.AppendLine(ArcToXML(oTool.Diameter / 2 - oTool.RadiusInf,
                                    oTool.RadiusInf,
                                    oTool.RadiusInf,
                                    true));
                    oStringBuilder.AppendLine(PointToXML(oTool.Diameter / 2, oTool.RadiusInf));
                }
                else
                {
                    oStringBuilder.AppendLine(PointToXML(oTool.Diameter / 2, 0));
                }
                if (oTool.RadiusSup > 0)
                {
                    oStringBuilder.AppendLine(PointToXML(oTool.Diameter / 2, oTool.Length - oTool.RadiusSup));
                    oStringBuilder.AppendLine(ArcToXML(oTool.Diameter / 2 - oTool.RadiusSup,
                                                     oTool.Length - oTool.RadiusSup,
                                                     oTool.RadiusSup,
                                                     true));
                    oStringBuilder.AppendLine(PointToXML(oTool.Diameter / 2 - oTool.RadiusSup, oTool.Length));
                }
                else
                {
                    oStringBuilder.AppendLine(PointToXML(oTool.Diameter / 2, oTool.Length));
                }
            }
            else
            {
                double X_lower_rad;
                double Z_lower_rad;
                if (oTool.RadiusInf > 0)
                {
                    oStringBuilder.AppendLine(PointToXML(oTool.Diameter / 2 - oTool.RadiusInf, 0));
                    oStringBuilder.AppendLine(ArcToXML(oTool.Diameter / 2 - oTool.RadiusInf,
                                    oTool.RadiusInf,
                                    oTool.RadiusInf,
                                    true));
                    X_lower_rad = (((oTool.Diameter / 2) - oTool.RadiusInf) + Math.Cos(Utilities.DegreesToRadians(oTool.SideAngle)) * oTool.RadiusInf);
                    Z_lower_rad = (oTool.RadiusInf + (Math.Sin(Utilities.DegreesToRadians(oTool.SideAngle)) * oTool.RadiusInf));
                }
                else
                {
                    X_lower_rad = oTool.Diameter / 2;
                    Z_lower_rad = 0;
                }
                oStringBuilder.AppendLine(PointToXML(X_lower_rad, Z_lower_rad));

                if (oTool.RadiusSup > 0)
                {
                    double X_no_Arc = X_lower_rad - ((oTool.Length - Z_lower_rad) * System.Math.Tan(Utilities.DegreesToRadians(oTool.SideAngle)));
                    double Z_Arc_Center = oTool.Length - oTool.RadiusSup;
                    double X_Arc_Center = X_no_Arc - ((Math.Tan(Utilities.DegreesToRadians((90 - oTool.SideAngle) / 2))) * oTool.RadiusSup);
                    double Z_Arc_Start = Z_Arc_Center + (Math.Sin(Utilities.DegreesToRadians(oTool.SideAngle))* oTool.RadiusSup);
                    double X_Arc_Start = X_Arc_Center + (Math.Cos(Utilities.DegreesToRadians(oTool.SideAngle)) * oTool.RadiusSup); ;

                    oStringBuilder.AppendLine(PointToXML(X_Arc_Start, Z_Arc_Start));
                    oStringBuilder.AppendLine(ArcToXML(X_Arc_Center,Z_Arc_Center,
                                                     oTool.RadiusSup,
                                                     true));
                    oStringBuilder.AppendLine(PointToXML(X_Arc_Center, oTool.Length));
                }
                else
                {
                    oStringBuilder.AppendLine(PointToXML(X_lower_rad - ((oTool.Length-Z_lower_rad) * System.Math.Tan(Utilities.DegreesToRadians(oTool.SideAngle))), oTool.Length));
                }
                                
            }
            oStringBuilder.AppendLine(PointToXML(0, oTool.Length));

            oStringBuilder.AppendLine(Indent(4) + "</SOR>");
            oStringBuilder.AppendLine(Indent(3) + "</Cutter>");

            // Write ToolHolder Part
            WriteHolderSTL(ref oStringBuilder, oTool.ToolHolderProfile, oTool.Name);

            oStringBuilder.AppendLine(Indent(4) + "<GagePoint>");
            oStringBuilder.AppendLine(Indent(6) + "<X>0</X>");
            oStringBuilder.AppendLine(Indent(6) + "<Y>0</Y>");
            oStringBuilder.AppendLine(Indent(6) + String.Format(System.Globalization.CultureInfo.InvariantCulture, "<Z>{0}</Z>", oTool.Gauge));
            oStringBuilder.AppendLine(Indent(4) + "</GagePoint>");
            oStringBuilder.AppendLine(Indent(2) + "</Tool>");
        }

        public static string ConstructToolName(int tool_num, string tool_name, bool tool_id_use_num, bool tool_id_use_name)
        {
            string name = "";
            if (tool_id_use_num)
                name = "" + tool_num;
            if (tool_id_use_name)
                name += ((name != "") ? "_" : "") + tool_name;

            return name;
        }

        private static string PointToXML(double x, double z)
        {
            return String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0}<Pt>\n{1}<X>{2}</X>\n{3}<Z>{4}</Z>\n{5}</Pt>",
                Indent(5),
                Indent(6),
                Utilities.ToInvariantCultureString(Math.Round(x, 4)),
                Indent(6),
                Utilities.ToInvariantCultureString(Math.Round(z, 4)),
                Indent(5));

        }

        private static string ArcToXML(double x, double z, double r, bool is_cw)
        {
            return String.Format(System.Globalization.CultureInfo.InvariantCulture, "{0}<Arc>\n{1}<X>{2}</X>\n{3}<Z>{4}</Z>\n{5}<R>{6}</R>\n{7}<Dir>{8}</Dir>\n{9}</Arc>",
                Indent(5),
                Indent(6),
                Utilities.ToInvariantCultureString(Math.Round(x, 4)),
                Indent(6),
                Utilities.ToInvariantCultureString(Math.Round(z, 4)),
                Indent(6),
                Utilities.ToInvariantCultureString(Math.Round(r, 4)),
                Indent(6),
                (is_cw ? "CW" : "CCW"),
                Indent(5));
        }

        /* Write tool shank part */
        private static void WriteShankSTL(ref StringBuilder oStringBuilder, string shank_stl_fpath, string tool_name, bool drillmill)
        {
            if (!String.IsNullOrEmpty(shank_stl_fpath))
            {
                oStringBuilder.AppendLine(Indent(4) + String.Format("<Model ID=\"{0}_Shank\" ParentID=\"{0}_Holder\" Type=\"STL\" Insert=\"no\" DRILLMILL=\"{1}\">",
                    tool_name, (drillmill ? "TRUE" : "FALSE")));
                oStringBuilder.AppendLine(Indent(5) + String.Format("<FileName>{0}</FileName>", shank_stl_fpath));
                oStringBuilder.AppendLine(Indent(5) + String.Format("<Normal>Outward</Normal>"));
                oStringBuilder.AppendLine(Indent(5) + String.Format("<SpindleDirection>CW</SpindleDirection>"));
                oStringBuilder.AppendLine(Indent(5) + String.Format("<NoSpin>0</NoSpin>"));
                oStringBuilder.AppendLine(Indent(5) + String.Format("<Alternate>off</Alternate>"));
                oStringBuilder.AppendLine(Indent(4) + String.Format("</Model>"));
            }
        }

        /* Write tool holder part */
        private static void WriteHolderSTL(ref StringBuilder oStringBuilder, string holder_stl_fpath, string tool_name)
        {
            if (!String.IsNullOrEmpty(holder_stl_fpath))
            {
                oStringBuilder.AppendLine(Indent(3) + "<Holder>");
                oStringBuilder.AppendLine(Indent(4) + String.Format("<Model ID=\"{0}_Holder\" Type=\"STL\" Insert=\"no\" DRILLMILL=\"FALSE\">",
                                            tool_name));
                oStringBuilder.AppendLine(Indent(5) + String.Format("<FileName>{0}</FileName>", holder_stl_fpath));
                oStringBuilder.AppendLine(Indent(5) + String.Format("<Normal>Outward</Normal>"));
                oStringBuilder.AppendLine(Indent(5) + String.Format("<NoSpin>0</NoSpin>"));
                oStringBuilder.AppendLine(Indent(5) + String.Format("<Alternate>off</Alternate>"));
                oStringBuilder.AppendLine(Indent(5) + String.Format("</Model>"));
                oStringBuilder.AppendLine(Indent(4) + "</Holder>");
            }
        }

        private static string Indent(int num)
        {
            if (num == 0)
                return "";
            else
                return string.Concat(Enumerable.Repeat("  ", num));
        }


        /// <summary>
        /// Match ToolNumber and NCProgToolNumber
        /// </summary>
        /// <param name="ToolName"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static int GetToolNumberAccordingToNCProgToolNumber(string ToolName, string ToolPath, List<NCProgramInfo> nc_progs)
        {
            if (nc_progs == null) return 0;
            foreach (NCProgramInfo oNCProg in nc_progs)
            {
                if (oNCProg.oTools != null)
                {
                    for (int i = 0; i <= oNCProg.oTools.Count - 1; i++)
                    {
                        if (ToolName == oNCProg.oTools[i].Name && ToolPath == oNCProg.oTools[i].Toolpath)
                        return oNCProg.oTools[i].Number;
                    }
                }
            }
            return 0;
        }
       

        /// <summary>
        /// Create the Tools
        /// </summary>
        /// <param name="ToolPathes"></param>
        /// <remarks></remarks>
        public static bool CreateListOfTools(List<string> ToolPathes, string OutputDir, List<NCProgramInfo> nc_progs,
                                             ref List<int> UsedToolNumbers, ref PowerMILLExporter.Tools.Tools oTools)
        {
            string sToolHolderPath = "",
                   sToolShankPath = "";
            if (ToolPathes == null) return true;
            if (ToolPathes.Count == 0) return true;

            for (int i = 0; i <= ToolPathes.Count - 1; i++)
            {
                try
                {
                    EventLogger.WriteToEvengLog("Get tool info for toolpath " + ToolPathes[i]);
                    PowerMILLAutomation.Execute("ACTIVATE Toolpath '" + ToolPathes[i] + "'");
                    string tmp = PowerMILLAutomation.GetParameterValue("tool.ShankSetValues").Trim();
                    string sToolType = PowerMILLAutomation.GetParameterValueTerse("tool.Type").Trim();// PowerMILLAutomation.ExecuteEx("print par terse 'tool.Type'").Trim();
                    string sToolName = PowerMILLAutomation.GetParameterValueTerse("tool.Name");//Trim was messing up tools with a trailing space -- PowerMILLAutomation.ExecuteEx("print par terse 'tool.Name'").Trim();
                    EventLogger.WriteToEvengLog("Get tool number from nc program dialog");
                    int iToolNumber = GetToolNumberAccordingToNCProgToolNumber(sToolName.Trim(), ToolPathes[i], nc_progs);
                    if (iToolNumber == 0)
                    {
                        if (Messages.ShowWarning(Properties.Resources.IDS_ToolNumberNotSet, sToolName) == System.Windows.MessageBoxResult.Yes)
                            return false;
                    }
                    double dHolderLength = 0;

                    // Check if the tool has already been added in a list
                    if (!MainModule.CheckExistingTool(sToolName, iToolNumber, sToolType, oTools))
                    {
                        if (UsedToolNumbers.Contains(iToolNumber))
                        {
                            Messages.ShowError(Properties.Resources.IDS_ToolNumberNotUnique, sToolName);
                            return false;
                        }
                        else
                            UsedToolNumbers.Add(iToolNumber);

                        EventLogger.WriteToEvengLog(String.Format("Tool {0} hasn't been added to the list yet. Add it.", sToolName));
                        if (Utilities.ConvertDecimalSeparator(PowerMILLAutomation.ExecuteEx("print = size(tool.HolderSetValues)")) > 0)
                        {
                            // Export Tool holder as .stl
                            EventLogger.WriteToEvengLog("Export tool holder");
                            sToolHolderPath = PowerMILLAutomation.PMILLExportToolSegment(sToolName, "TOOLHOLDER", OutputDir);
                        }
                        else
                        {
                            EventLogger.WriteToEvengLog(String.Format(Properties.Resources.IDS_ToolNoHolder, sToolName));
                            if (Messages.ShowWarning(Properties.Resources.IDS_ToolNoHolder, sToolName) == System.Windows.MessageBoxResult.No)
                                return false;
                        }

                        if (Utilities.ConvertDecimalSeparator(PowerMILLAutomation.ExecuteEx("print = size(tool.ShankSetValues)")) > 0)
                        {
                            // Export Tool shank as .stl
                            EventLogger.WriteToEvengLog("Export tool shank");
                            sToolShankPath = PowerMILLAutomation.PMILLExportToolSegment(sToolName, "TOOLSHANK", OutputDir);
                        }
                        else
                        {
                            EventLogger.WriteToEvengLog(String.Format(Properties.Resources.IDS_ToolNoShank, sToolName));
                            if (Messages.ShowWarning(Properties.Resources.IDS_ToolNoShank, sToolName) == System.Windows.MessageBoxResult.No)
                                return false;
                        }

                        GetToolHolderNShankProfilesData(sToolName, out dHolderLength);

                        // Get tools informations
                        double dOverHang = Math.Round(Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.Overhang")),4);

                        // Get tools informations
                        double dCuttingLength = Math.Round(Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.Length")),4);
                        //dHolderLength += dCuttingLength;
                        double dDiameter = Math.Round(Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.Diameter")),4);
                        int iNumberOFTeeth = Convert.ToInt32(PowerMILLAutomation.GetParameterValueTerse("tool.NumberOfFlutes"));
                        double dGauge = Math.Round(Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("gauge_length(Tool)")),4);

                        // Length = cutting length , OverLength (total Length) = Gauge + 10, shankDiameter = ToolDiameter
                        switch (sToolType)
                        {
                            case "end_mill":
                                {
                                    EventLogger.WriteToEvengLog("Create a milling tool object corresponding to the endmill tool and add it to the list of milling tools");
                                    MillingTool oEndMillTool = new MillingTool(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter,
                                                                        dCuttingLength, dHolderLength, dCuttingLength, sToolShankPath, iNumberOFTeeth, 0, "EndMill");
                                    //MillingTool oEndMillTool = new MillingTool(sToolName, sbHolderProfile.ToString(), iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength, dOverToolLength, dCuttingLength, sbShankProfile.ToString(), iNumberOFTeeth, 0);
                                    oTools.oMillingTools.Add(oEndMillTool);
                                    break;
                                }
                            case "tip_radiused":
                                {
                                    EventLogger.WriteToEvengLog("Create a milling tool object corresponding to the tip radiused tool and add it to the list of milling tools");
                                    double dCornerRadius = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.TipRadius"));
                                    MillingTool oTipRadiusedTool = new MillingTool(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter,
                                                                        dCuttingLength, dHolderLength, dCuttingLength, sToolShankPath, iNumberOFTeeth, dCornerRadius, "TipRadiused");
                                    //MillingTool oTipRadiusedTool = new MillingTool(sToolName, sbHolderProfile.ToString(), iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength, dOverToolLength, dCuttingLength, sbShankProfile.ToString(), iNumberOFTeeth, dCornerRadius);
                                    oTools.oMillingTools.Add(oTipRadiusedTool);
                                    break;
                                }
                            case "ball_nosed":
                                {
                                    EventLogger.WriteToEvengLog("Create a milling tool object corresponding to the ball nosed tool and add it to the list of milling tools");
                                    MillingTool oBallNosedTool = new MillingTool(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter,
                                                                        dCuttingLength, dHolderLength, dCuttingLength, sToolShankPath, iNumberOFTeeth, dDiameter / 2, "BallNosed");
                                    //MillingTool oBallNosedTool = new MillingTool(sToolName, sbHolderProfile.ToString(), iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength, dOverToolLength, dCuttingLength, sbShankProfile.ToString(), iNumberOFTeeth, dDiameter / 2);
                                    oTools.oMillingTools.Add(oBallNosedTool);
                                    break;
                                }
                            case "taper_spherical":
                                {
                                    EventLogger.WriteToEvengLog("Create a milling tool object corresponding to the taper spherical tool and add it to the list of taper spherical tools");
                                    double dCornerRadius = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.TipRadius"));
                                    double dNoseAngle = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.TaperAngle"));
                                    TaperSpherical oTaperSpherical = new TaperSpherical(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter,
                                                                        dCuttingLength, dHolderLength, dCuttingLength, sToolShankPath, iNumberOFTeeth, dCornerRadius, dNoseAngle, "TaperSpherical");
                                    //TaperSpherical oTaperSpherical = new TaperSpherical(sToolName, sbHolderProfile.ToString(), iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength, dOverToolLength, dCuttingLength, sbShankProfile.ToString(), iNumberOFTeeth, dCornerRadius, dNoseAngle);
                                    oTools.oTaperSphericalTools.Add(oTaperSpherical);
                                    break;
                                }
                            case "taper_tipped":
                                {
                                    EventLogger.WriteToEvengLog("Create a milling tool object corresponding to the taper tipped tool and add it to the list of taper tipped tools");
                                    double dTaperAngle = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.TaperAngle"));
                                    double dTaperHeight = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.TaperHeight"));
                                    double dTipRad = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.TipRadius"));
                                    double dTaperDia = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.TaperDiameter"));

                                    List<string> sProfileList = GetTapperTrippedProfile(dDiameter, dCuttingLength);
                                    TaperedTipped oTaperTrippedTool = new TaperedTipped(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter,
                                                                        dCuttingLength, dHolderLength, dCuttingLength, sToolShankPath, iNumberOFTeeth, dTaperAngle,
                                                                        dTaperDia, dTipRad, dTaperHeight, "TaperTipped");
                                    //TaperedTipped oTaperTrippedTool = new TaperedTipped(sToolName, sbHolderProfile.ToString(), iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength, dOverToolLength, dCuttingLength, sbShankProfile.ToString(), iNumberOFTeeth, dTaperAngle, dTaperDia, dTipRad, dTaperHeight);
                                    oTools.oTaperTippedTools.Add(oTaperTrippedTool);
                                    break;
                                }
                            case "off_centre_tip_rad":
                                {
                                    EventLogger.WriteToEvengLog("Create a off centre tip radius tool object corresponding to the endmill tool and add it to the list of special mill tools");
                                    double dTipRad = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.TipRadius"));
                                    double dX_Offset = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.TipRadiusCentre.X"));
                                    double dY_Offset = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.TipRadiusCentre.Y"));

                                    SpecialMill oOffCentreTipRad = new SpecialMill(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter,
                                                                        dCuttingLength, dHolderLength, dCuttingLength, dX_Offset, dY_Offset, dTipRad, sToolShankPath, iNumberOFTeeth,
                                                                        "OffCentreTipRad");
                                    oTools.oSpecialMILLTools.Add(oOffCentreTipRad);
                                    break;
                                }
                            case "barrel":
                                {
                                    EventLogger.WriteToEvengLog("Create a barrel tool object corresponding to the endmill tool and add it to the list of 2D tools");
                                    string sToolPath = PowerMILLAutomation.PMILLExportToolSegment(sToolName, "TOOL_PROFILE", OutputDir);
                                    _2DTool oFormTool = new _2DTool(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength,
                                                                    dHolderLength, dCuttingLength, sToolShankPath, iNumberOFTeeth, sToolPath, "Form");
                                    oTools.o2DTools.Add(oFormTool);
                                    break;
                                }
                            case "tipped_disc":
                                {
                                    EventLogger.WriteToEvengLog("Create a tipped disc tool object corresponding to the endmill tool and add it to the list of top and side mill tools");
                                    double dCornerRadius = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.TipRadius"));
                                    double dUpperCornerRadius = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.UpperTipRadius"));
                                    TopAndSideMill oTrippedDiscTool = new TopAndSideMill(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter,
                                                                        dCuttingLength, dHolderLength, dCuttingLength, sToolShankPath, iNumberOFTeeth, dCornerRadius,
                                                                        dUpperCornerRadius, 0, "TippedDisc");
                                    //TopAndSideMill oTrippedDiscTool = new TopAndSideMill(sToolName, sbHolderProfile.ToString(), iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength, dOverToolLength, dCuttingLength, sbShankProfile.ToString(), iNumberOFTeeth, dCornerRadius, dCornerRadius, 0);
                                    oTools.oTopAndSideMillTools.Add(oTrippedDiscTool);
                                    break;
                                }
                            case "drill":
                                {
                                    EventLogger.WriteToEvengLog("Create a drill tool object corresponding to the endmill tool and add it to the list of drilling tools");
                                    double dNoseAngle = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.DrillAngle"));
                                    double dLength = dCuttingLength - dDiameter / 2 / Math.Tan(Utilities.DegreesToRadians(dNoseAngle));
                                    DrillingTool oDrillTool = new DrillingTool(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter,
                                                                        dLength, dHolderLength, dCuttingLength, sToolShankPath, iNumberOFTeeth, dNoseAngle, "Drill");
                                    //DrillingTool oDrillTool = new DrillingTool(sToolName, sbHolderProfile.ToString(), iToolNumber, dOverHang, dGauge, dDiameter, dLength, dCuttingLength, dOverToolLength, sbShankProfile.ToString(), iNumberOFTeeth, dNoseAngle);
                                    oTools.oDrillingTools.Add(oDrillTool);
                                    break;
                                }
                            case "tap":
                                {
                                    EventLogger.WriteToEvengLog("Create a tap tool object corresponding to the endmill tool and add it to the list of tap tools");
                                    TapTool oTapTool = new TapTool(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength,
                                                                   dHolderLength, dCuttingLength, sToolShankPath, iNumberOFTeeth, 1, dDiameter * 0.7,
                                                                   dDiameter * 0.4, 0.12 * dCuttingLength, "Tap");
                                    //TapTool oTapTool = new TapTool(sToolName, sbHolderProfile.ToString(), iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength, dOverToolLength, dCuttingLength, sbShankProfile.ToString(), iNumberOFTeeth, 1, dDiameter * 0.7, dDiameter * 0.4, 0.12 * dCuttingLength);
                                    oTools.oTapTools.Add(oTapTool);
                                    break;
                                }
                            case "form":
                                {
                                    EventLogger.WriteToEvengLog("Create a form tool object corresponding to the endmill tool and add it to the list of 2D tools");
                                    string sToolPath = PowerMILLAutomation.PMILLExportToolSegment(sToolName, "TOOL_PROFILE", OutputDir);
                                    _2DTool oFormTool = new _2DTool(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength,
                                                                    dHolderLength, dCuttingLength, sToolShankPath, iNumberOFTeeth, sToolPath, "Form");
                                    //_2DTool oFormTool = new _2DTool(sToolName, sbHolderProfile.ToString(), iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength, dOverToolLength, dCuttingLength, sbShankProfile.ToString(), iNumberOFTeeth, sToolPath);
                                    oTools.o2DTools.Add(oFormTool);
                                    break;
                                }
                            case "routing":
                                {
                                    EventLogger.WriteToEvengLog("Create a routing tool object corresponding to the endmill tool and add it to the list of 2D tools");
                                    string sToolPath = PowerMILLAutomation.PMILLExportToolSegment(sToolName, "TOOL_PROFILE", OutputDir);
                                    _2DTool oRoutingTool = new _2DTool(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength,
                                                                       dHolderLength, dCuttingLength, sToolShankPath, iNumberOFTeeth, sToolPath, "Routing");
                                    //_2DTool oRoutingTool = new _2DTool(sToolName, sbHolderProfile.ToString(), iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength, dOverToolLength, dCuttingLength, sbShankProfile.ToString(), iNumberOFTeeth, sToolPath);
                                    oTools.o2DTools.Add(oRoutingTool);
                                    break;
                                }
                            case "dovetail":
                                {
                                    EventLogger.WriteToEvengLog("Create a dovetail tool object corresponding to the endmill tool and add it to the list of top and side mill tools");
                                    double dLowerTipRadius = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.TipRadius"));
                                    double dUpperTipRadius = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.UpperTipRadius"));
                                    double dTaperAngle = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.TaperAngle"));
                                    TopAndSideMill oTrippedDiscTool = new TopAndSideMill(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength, dCuttingLength, dCuttingLength, sToolShankPath, iNumberOFTeeth, dLowerTipRadius, dUpperTipRadius, dTaperAngle);
                                    oTools.oTopAndSideMillTools.Add(oTrippedDiscTool);
                                    break;
                                }
                            case "thread_mill":
                                {
                                    EventLogger.WriteToEvengLog("Create a thread milling tool object corresponding to the endmill tool and add it to the list of tap tools");
                                    double dToolStep = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.GetParameterValueTerse("tool.Pitch"));
                                    TapTool oThreadMillTool = new TapTool(sToolName, sToolHolderPath, iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength,
                                                                          dHolderLength, dCuttingLength, sToolShankPath, iNumberOFTeeth, dToolStep,
                                                                          dDiameter * 0.7, dDiameter * 0.4, 0.12 * dCuttingLength, "ThreadMill");
                                    //TapTool oThreadMillTool = new TapTool(sToolName, sbHolderProfile.ToString(), iToolNumber, dOverHang, dGauge, dDiameter, dCuttingLength, dOverToolLength, dCuttingLength, sbShankProfile.ToString(), iNumberOFTeeth, dToolStep, dDiameter * 0.7, dDiameter * 0.4, 0.12 * dCuttingLength);
                                    oTools.oTapTools.Add(oThreadMillTool);
                                    break;
                                }
                        }
                    }
                }
                catch 
                {
                    return false;
                }
            }
            return true;
        }

        private static void GetToolHolderNShankProfilesData(string tool_name, out double holder_length)
        {
            PowerMILL.Tool oTool = GetTool(tool_name);
            double dHolderLen = 0/*,
                   dShankLen = 0*/;

            holder_length = 0;

            if (oTool == null) return;

            //Holder. Temporarily comment out lines below
            //for (int i = oTool.HolderSetValues.Count - 1; i >= 0; i--)
            //{
            //    dHolderLen += oTool.HolderSetValues[i].Length;
            //}

            ////Shank
            //for (int i = oTool.ShankSetValues.Count - 1; i >= 0; i--)
            //{
            //    dShankLen += oTool.ShankSetValues[i].Length;
            //}

            holder_length = dHolderLen;
        }

        private static PowerMILL.Tool GetTool(string tool_name)
        {
            foreach (PowerMILL.Tool oTool in PowerMILLAutomation.oPServices.Project.Tools)
                if (oTool.Name == tool_name) return oTool;

            return null;
        }

        #region "SpecialTool Management (profile)"
        /// <summary>
        /// Create the TapperTippedProfile
        /// </summary>
        /// <param name="ToolDiameter">Tool diameter</param>
        /// <param name="ToolLength">Tool Length</param>
        /// <returns>The Profile</returns>
        /// <remarks></remarks>
        private static List<string> GetTapperTrippedProfile(double ToolDiameter, double ToolLength)
        {

            // Get the profile points
            double AlphaInRadius = (Utilities.ConvertDecimalSeparator(PowerMILLAutomation.ExecuteEx("print par terse'tool.TaperAngle'")) * System.Math.PI) / 180;
            double x3 = ToolDiameter / 2;
            double y3 = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.ExecuteEx("print par terse'tool.TaperHeight'"));
            double R = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.ExecuteEx("print par terse'tool.TipRadius'"));
            double y2 = R * (1 - System.Math.Sin(AlphaInRadius));
            double x2 = x3 - (y3 - y2) * System.Math.Tan(AlphaInRadius);
            double x1 = x2 - R * System.Math.Cos(AlphaInRadius);
            double y1 = 0;


            // Get the List of string for the profile
            List<string> sListOfString = new List<string>();
            sListOfString.Add("  |LINE|CUT|0,0," + x1 + "," + y1 + "|");
            sListOfString.Add("  |ARC|CUT|" + x1 + "," + y1 + "," + x2 + "," + y2 + "," + x1 + "," + y1 + R + "|CCW|");
            sListOfString.Add("  |LINE|CUT|" + x2 + "," + y2 + "," + x3 + "," + y3 + "|");
            sListOfString.Add("  |LINE|CUT|" + x3 + "," + y3 + "," + x3 + "," + (ToolLength - y3) + "|");
            sListOfString.Add("  |LINE|CUT|" + x3 + "," + (ToolLength - y3) + "," + 0 + "," + (ToolLength - y3) + "|");

            return sListOfString;
        }

        /// <summary>
        /// Create the Off Centre Tip Rad Profile
        /// </summary>
        /// <param name="ToolDiameter">Tool Diameter</param>
        /// <param name="ToolLength">Tool Length</param>
        /// <returns>The Profile</returns>
        /// <remarks></remarks>
        private static List<string> GetOff_Centre_Tip_Rad_Profile(double ToolDiameter, double ToolLength)
        {
            double R = Utilities.ConvertDecimalSeparator(PowerMILLAutomation.ExecuteEx("print par terse'tool.TipRadius'"));

            // Get the TipRadius centre
            string[] slist = Regex.Split(PowerMILLAutomation.ExecuteEx("print par 'tool.TipRadiusCentre'").Replace("(ARRAY)", ""), "[L]");
            for (int k = 0; k <= slist.Length - 1; k++)
            {
                slist[k] = slist[k].Remove(0, 10).Replace("[L]", "").Trim();
            }
            double I = Utilities.ConvertDecimalSeparator(slist[0]);
            double J = Utilities.ConvertDecimalSeparator(slist[1]);

            // Get the profile points
            double x1 = System.Math.Sqrt((R * R) - (J * J)) + I;
            double Y1 = 0;
            double x2 = ToolDiameter / 2;
            double y2 = System.Math.Sqrt((R * R) - System.Math.Pow((x2 - I), 2)) + J;

            // Get the List of string for the profile
            List<string> sListOfString = new List<string>();
            sListOfString.Add("  |LINE|CUT|0,0," + x1 + "," + Y1 + "|");
            sListOfString.Add("  |ARC|CUT|" + x1 + "," + Y1 + "," + x2 + "," + y2 + "," + I + "," + J + "|CCW|");
            sListOfString.Add("  |LINE|CUT|" + x2 + "," + y2 + "," + x2 + "," + (ToolLength - y2) + "|");
            sListOfString.Add("  |LINE|CUT|" + x2 + "," + (ToolLength - y2) + "," + 0 + "," + (ToolLength - y2) + "|");

            return sListOfString;
        }
        #endregion

    }
}
