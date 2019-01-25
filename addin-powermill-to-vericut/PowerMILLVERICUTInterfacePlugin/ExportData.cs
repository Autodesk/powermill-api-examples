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
using System.Globalization;
using PowerMILLExporter;
using System.Windows;
using System.Xml;
using System.Xml.Linq;

namespace PowerMILLVERICUTInterfacePlugin
{
    class ExportData
    {

        public static string GetToolLibraryFilePath(string template_fpath)
        {
            string tls_fpath = "";

            XDocument doc = XDocument.Load(template_fpath);
            XElement element = doc.Descendants("Setup").First();
            if (element.Descendants("ToolMan").Count() > 0)
            {
                element = element.Descendants("ToolMan").First();
                if (element.Descendants("Library").Count() > 0)
                    tls_fpath = element.Descendants("Library").First().Value;
            }

            return tls_fpath;
        }
        /// <summary>
        /// Create the different Output Files
        /// </summary>
        /// <remarks>nxf File,stl, iso, ...</remarks>
        public static string CreateOutputFiles(ProjectData proj_data, string output_dir, PluginSettings plugin_settings)
        {
            string project_fpath = Path.Combine(output_dir, ProjectData.ProjectName + ".vcproject");

            XDocument main_doc;
            try
            {
                main_doc = XDocument.Load(proj_data.proj_template);
            }
            catch
            {
                Messages.ShowError(Properties.Resources.IDS_VericutTemplateReadFailure, proj_data.proj_template);
                return "*";
            }

            main_doc.Descendants("Setup").Remove();
            if (main_doc.Descendants("InfoUserName").Count() > 0)
                main_doc.Descendants("InfoUserName").First().Value =
                    String.Format("PowerMill to Vericut plugin v{0}.{1}.{2}",
                                    System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major,
                                    System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor,
                                    System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build);

            for (int i = 0; i < proj_data.setup_infos.Count; i++)
            {
                XDocument doc;
                XElement setup;
                SetupInfo setup_info = proj_data.setup_infos[i];

                PowerMILLAutomation.CreateTransformWorkplane("Transform_Vericut");
                PowerMILLAutomation.ActivateWorkplane(setup_info.attach_workplane);
                //WorkPlaneOrigin model_origin = MainModule.GetWorkplaneCoords(PowerMILLAutomation.oToken, setup_info.model_workplane, false, Properties.Resources.IDS_WorkplaneFailedToGet);
                WorkPlaneOrigin model_origin = MainModule.GetWorkplaneCoords(PowerMILLAutomation.oToken, "Transform_Vericut", true, Properties.Resources.IDS_WorkplaneFailedToGet);//Changed relation to true and look for the Transform position from the model location to align correctly in Vericut
                PowerMILLAutomation.DeleteTransformWorkplane("Transform_Vericut");


                string template = null;
                if (String.IsNullOrEmpty(setup_info.template) || !File.Exists(setup_info.template))
                    template = proj_data.proj_template;
                else
                    template = setup_info.template;
                doc = XDocument.Load(template);

                // Add path to items from the template
                IEnumerable<XElement> rows = doc.Descendants("File");
                foreach (XElement file in rows.ToList())
                {
                    string file_name = (file.FirstNode != null) ? file.FirstNode.ToString() : null;
                    if (!String.IsNullOrEmpty(file_name) && !String.IsNullOrEmpty(Path.GetExtension(file_name))
                        && String.IsNullOrEmpty(Path.GetDirectoryName(file_name)))
                    {
                        string full_path = Path.Combine(Path.GetDirectoryName(template), file_name);
                        string new_element = file.ToString().Replace(file_name, full_path);

                        file.ReplaceWith(XElement.Parse(new_element));
                    }
                }

                setup = doc.Descendants("Setup").First();
                setup.Attribute("Name").Value = setup_info.name;

                PowerMILLExporter.Tools.Tools setup_tools;
                GetListOfTools(setup_info, plugin_settings, proj_data.output_dir, out setup_tools);
                setup_info.oTools = setup_tools;

                if (setup_info.tools_to_use == 0)
                {
                    EventLogger.WriteToEvengLog("Write tool info to tools.xml");
                    WriteTools.WriteToolsToXML(setup_info.oTools, output_dir, setup_info.name + "_tools.tls", plugin_settings.tool_id_use_num, plugin_settings.tool_id_use_name);

                    //output_dir, setup_info.name + "_tools.tls"
                    SetTools(ref setup, Path.Combine(output_dir, setup_info.name + "_tools.tls"));
                }
                else 
                {
                    List<string> lib_tools = GetListOfToolsFromLibrary(setup_info.tls_fpath);
                    List<string> missing_tools = WriteTools.FindToolsMissingFromLibrary(setup_info.oTools, lib_tools);
                    if (missing_tools.Count > 0)
                    {
                        if (setup_info.tools_to_use == 1)
                        {
                            Messages.ShowWarning(String.Format("Setup {0}: the following tools are missing from the library:\r\n {1}.", setup_info.name, String.Join(Environment.NewLine, missing_tools.ToArray())));
                        }
                        else if (setup_info.tools_to_use == 2)
                        {
                            EventLogger.WriteToEvengLog("Write tool info to tools.xml");
                            WriteTools.AppendToolsToXML(setup_info.oTools, output_dir, setup_info.tls_fpath, missing_tools);

                            //output_dir, setup_info.name + "_tools.tls"
                            SetTools(ref setup, Path.Combine(output_dir, setup_info.tls_fpath));
                        }
                    }
                }

                SetWorkOffsets(ref setup, setup_info);

                if (setup_info.tools_to_use == 1 || setup_info.tools_to_use == 2)
                    SetNCPrograms(ref setup, setup_info, false, true);
                else
                    SetNCPrograms(ref setup, setup_info, plugin_settings.tool_id_use_num, plugin_settings.tool_id_use_name);

                SetCSysInfo(ref setup, setup_info, proj_data.cut_stock_csys);

                /* Don't set block info for 2nd and other setups if using cut stock transition */
                if (i == 0 || String.IsNullOrEmpty(proj_data.cut_stock_csys)) 
                    SetBlockInfo(ref setup, output_dir, setup_info, model_origin, plugin_settings.export_options.block_tol);

                SetFixturesInfo(ref setup, output_dir, setup_info, model_origin, plugin_settings.export_options.fixture_tol);

                /* Don't set part info for 2nd and other setups if using cut stock transition */
                if (i == 0 || String.IsNullOrEmpty(proj_data.cut_stock_csys))
                    SetPartsInfo(ref setup, output_dir, setup_info, model_origin, plugin_settings.export_options.model_tol);

                if (i == 0)
                {
                    main_doc.Elements("VcProject").First().Add(setup);
                }
                else
                {
                    if (main_doc.Descendants("Setup").Count() > 0)
                        main_doc.Descendants("Setup").Last().AddAfterSelf(setup);
                }                    

            }

            main_doc.Save(project_fpath);


            return project_fpath;
        }

        private static List<string> GetListOfToolsFromLibrary(string tls_fpath)
        {
            List<string> names = new List<string>();
            try
            {
                if (File.Exists(tls_fpath))
                {
                    XDocument doc = XDocument.Load(tls_fpath);
                    names = new List<string>((from e in doc.Descendants("Description") select e.Value).Distinct());
                }
            }
            catch 
            {}
            return names;
        }

        private static void SetWorkOffsets(ref XElement x_setup, SetupInfo setup_info)
        {
            XElement base_work_offsets = null,
                     program_zeros = null,
                     work_offsets = null,
                     temp = null;
            

            if (x_setup.Descendants("Table").Where(e => e.Attribute("Name").Value.ToUpper() == "BASE WORK OFFSET").Count() > 0)
            {
                base_work_offsets = x_setup.Descendants("Table").Where(e => e.Attribute("Name").Value.ToUpper() == "BASE WORK OFFSET").First();
                //base_work_offsets.Descendants("Row").Remove();
            }
            if (x_setup.Descendants("Table").Where(e => e.Attribute("Name").Value.ToUpper() == "PROGRAM ZERO").Count() > 0)
            {
                program_zeros = x_setup.Descendants("Table").Where(e => e.Attribute("Name").Value.ToUpper() == "PROGRAM ZERO").First();
                //program_zeros.Descendants("Row").Remove();
            }
            if (x_setup.Descendants("Table").Where(e => e.Attribute("Name").Value.ToUpper() == "WORK OFFSETS").Count() > 0)
            {
                List<XElement> els = (List<XElement>)x_setup.Descendants("Table").Where(e => e.Attribute("Name").Value.ToUpper() == "WORK OFFSETS").ToList();
                work_offsets = x_setup.Descendants("Table").Where(e => e.Attribute("Name").Value.ToUpper() == "WORK OFFSETS").First();
                //work_offsets.Descendants("Row").Remove();
            }

            if (setup_info.offsets == null) return;

            foreach (WorkOffset offset in setup_info.offsets)
            {
                temp = new XElement("Row",
                                        new XAttribute("Index", offset.register),
                                        new XAttribute("SubIndex", "1"),
                                        new XAttribute("Auto", "auto"),
                                        new XElement("System", offset.subsystem),
                                        new XElement("From",
                                            new XAttribute("CSystem", "off"),
                                            new XAttribute("X", 0),
                                            new XAttribute("Y", 0),
                                            new XAttribute("Z", 0),
                                            offset.from_component),
                                        new XElement("To",
                                            new XAttribute("CSystem", "on"),
                                            new XAttribute("X", 0),
                                            new XAttribute("Y", 0),
                                            new XAttribute("Z", 0),
                                            offset.to_ucs));
                if (offset.name.Equals("Program zero", StringComparison.OrdinalIgnoreCase))
                {
                    if (program_zeros != null)
                        program_zeros.Add(temp);
                    else
                    {
                        if (x_setup.Descendants("GCode") != null)
                        {
                            XElement gcode = x_setup.Descendants("GCode").First();
                            gcode.AddFirst(new XElement("Table",
                                            new XAttribute("Name", "Program zero"),
                                            temp));
                        }
                    }

                }
                else if (offset.name.Equals("Work Offsets", StringComparison.OrdinalIgnoreCase))
                {
                    if (work_offsets != null)
                        work_offsets.Add(temp);
                    else
                    {
                        if (x_setup.Descendants("GCode") != null)
                        {
                            XElement gcode = x_setup.Descendants("GCode").First();
                            gcode.AddFirst(new XElement("Table",
                                            new XAttribute("Name", "Work Offsets"),
                                            temp));
                        }
                    }
                }
            }
       }

        /// <summary>
        /// Create the NXF File
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        private static bool GetListOfTools(SetupInfo setup_info, PluginSettings plugin_settings, string output_dir, out PowerMILLExporter.Tools.Tools setup_tools)
        {
            List<PowerMILLExporter.Tools.NCProgramTool> oTools = new List<PowerMILLExporter.Tools.NCProgramTool>();
            int iInitToolNum = -1;
            bool bReturn = true;
            List<int> oUsedToolNumbers = new List<int>();
            PowerMILLAutomation.ExecuteEx(String.Format(System.Globalization.CultureInfo.InvariantCulture,
                "$Powermill.Export.TriangleTolerance = \"{0}\"", plugin_settings.export_options.tool_tol));

            setup_tools = new PowerMILLExporter.Tools.Tools();
            // For each NC prog, create and add tools in list of tools
            for (int nc_i = 0; nc_i < setup_info.nc_progs.Count; nc_i++)
            {
                if (String.IsNullOrEmpty(setup_info.nc_progs[nc_i].sName)) continue;
                EventLogger.WriteToEvengLog("Get all toolpaths for nc program " + setup_info.nc_progs[nc_i].sName);

                setup_info.toolpaths = PowerMILLAutomation.GetNCProgToolpathes(setup_info.nc_progs[nc_i].sName);
                EventLogger.WriteToEvengLog("Get all tools for nc program " + setup_info.nc_progs[nc_i].sName);

                //OTools get's out of the function with the first toolpath missing
                MainModule.FillNCProgramTools(setup_info.nc_progs[nc_i].sName, setup_info.toolpaths, nc_i, out oTools, ref iInitToolNum);

                setup_info.nc_progs[nc_i].oTools = oTools;


                EventLogger.WriteToEvengLog("Get tool info for nc program " + setup_info.nc_progs[nc_i].sName);

                bReturn &= WriteTools.CreateListOfTools(setup_info.toolpaths, output_dir, setup_info.nc_progs, ref oUsedToolNumbers, ref setup_tools);
            }

            return bReturn;
        }
        

        private static void SetPartsInfo(ref XElement x_setup, string output_dir, SetupInfo setup_info, WorkPlaneOrigin model_origin, 
                            double model_tol)
        {
            if (setup_info.part_models != null)
                if (setup_info.part_models.Count > 0)
                {
                    PowerMILLAutomation.SetExportTolerance(model_tol);
                    /* Export fixtures first */
                    setup_info.part_stl_fpaths = new List<string>();
                    foreach (string part_name in setup_info.part_models)
                        setup_info.part_stl_fpaths.Add(PowerMILLAutomation.ExportModel(part_name, output_dir));
                }
            /* Set part data in the .xml file */
            if (setup_info.part_stl_fpaths != null)
                if (setup_info.part_stl_fpaths.Count > 0)
                {
                    /* Set fixture data in the .xml file */
                    XElement target = x_setup.Descendants("Component").Where(e => e.Attribute("Name").Value.ToUpper() == "DESIGN" && e.Attribute("Type") != null).First();

                    /* Remove template stock from new project */
                    remove_components_from_xml(target.Elements(), Properties.Resources.IDS_RemoveDesign);

                    target.Element("Attach").Value = setup_info.fixture_attach_to;
                    foreach (string stl_fpath in setup_info.part_stl_fpaths)
                    {
                        if (String.IsNullOrEmpty(stl_fpath)) continue;
                        target.Add(new XElement("STL",
					new XAttribute("Unit", (ProjectData.PmillUnitsIsInch ? "inch" : "Millimeter")),
                                        new XAttribute("Normal", "Computed"),
                                        new XAttribute("Visible", "on"),
                                        new XAttribute("XRGB", "0x0033CC00"),
                                        new XAttribute("Mirror", "off"),
                                        new XAttribute("MirrorCsys", "0"),
                                        //new XAttribute("Hue", 1),
                                        new XElement("Position",
                                            new XAttribute("X", model_origin.dX),
                                            new XAttribute("Y", model_origin.dY),
                                            new XAttribute("Z", model_origin.dZ)),
                                        new XElement("Rotation",
                                            new XAttribute("I", model_origin.dXangle),
                                            new XAttribute("J", model_origin.dYangle),
                                            new XAttribute("K", model_origin.dZangle)),
                                        new XElement("File", stl_fpath)
                                        ));
                    }
                }
        }

        private static void SetFixturesInfo(ref XElement x_setup, string output_dir, SetupInfo setup_info, WorkPlaneOrigin model_origin, 
                                            double fixture_tol)
        {
            if (setup_info.fixture_models != null)
                if (setup_info.fixture_models.Count > 0)
                {
                    PowerMILLAutomation.SetExportTolerance(fixture_tol);
                    /* Export fixtures first */
                    setup_info.fixture_stl_fpaths = new List<string>();
                    foreach (string fixture_name in setup_info.fixture_models)
                        setup_info.fixture_stl_fpaths.Add(PowerMILLAutomation.ExportModel(fixture_name, output_dir));
                }

            /* Set fixture data in the .xml file */
            if (setup_info.fixture_stl_fpaths != null)
                if (setup_info.fixture_stl_fpaths.Count > 0)
                {
                    XElement target = x_setup.Descendants("Component").Where(e => e.Attribute("Name").Value.ToUpper() == "FIXTURE" && e.Attribute("Type") != null).First();

                    /* Remove fixture from new project*/
                    remove_components_from_xml(target.Elements(), Properties.Resources.IDS_RemoveFixture);

                    target.Element("Attach").Value = setup_info.fixture_attach_to;
                    foreach (string stl_fpath in setup_info.fixture_stl_fpaths)
                    {
                        if (String.IsNullOrEmpty(stl_fpath)) continue;
                        target.Add(new XElement("STL",
					new XAttribute("Unit", (ProjectData.PmillUnitsIsInch ? "inch" : "Millimeter")),
                                        new XAttribute("Normal", "Computed"),
                                        new XAttribute("Visible", "on"),
                                        new XAttribute("XRGB", "0x00A0A0A0"),
                                        new XAttribute("Mirror", "off"),
                                        new XAttribute("MirrorCsys", "0"),
                                        //new XAttribute("Hue", 1),
                                        new XElement("Position",
                                            new XAttribute("X", model_origin.dX),
                                            new XAttribute("Y", model_origin.dY),
                                            new XAttribute("Z", model_origin.dZ)),
                                        new XElement("Rotation",
                                            new XAttribute("I", model_origin.dXangle),
                                            new XAttribute("J", model_origin.dYangle),
                                            new XAttribute("K", model_origin.dZangle)),
                                        new XElement("File", stl_fpath)
                                        ));
                    }
                }
        }

        /// <summary>
        /// Optionally removes components from the output xml file that aren't "Attach", "Position", & "Rotation".
        /// If these components exist, a message will be shown allowing the user to choose to remove them or not.
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        private static void remove_components_from_xml(IEnumerable<XElement> content, string msg)
        {
            if (content != null && content.Any() && Messages.ShowWarning(msg) == System.Windows.MessageBoxResult.Yes)
                foreach (XElement item in content.ToList())
                {
                    if (!(item.Name.ToString().Equals("Attach") || item.Name.ToString().Equals("Position")
                        || item.Name.ToString().Equals("Rotation")))
                        item.Remove();
                }
        }

        private static void SetBlockInfo(ref XElement x_setup, string output_dir, SetupInfo setup_info, WorkPlaneOrigin model_origin, 
                                         double block_tol)
        {
            /* Depending on what's selected in UI, export block associated with toolpath or a stock model */
            if (!String.IsNullOrEmpty(setup_info.block_toolpath))
            {
                PowerMILLAutomation.SetExportTolerance(block_tol);
                setup_info.block_stl_fpath = PowerMILLAutomation.ExportBlock(setup_info.block_toolpath, output_dir);
            }
            if (!String.IsNullOrEmpty(setup_info.block_stockmodel))
            {
                PowerMILLAutomation.SetExportTolerance(block_tol);
                setup_info.block_stl_fpath = PowerMILLAutomation.ExportStockModel(setup_info.block_stockmodel, output_dir);
            }

            /* Set block data in the .xml file */
            if (!String.IsNullOrEmpty(setup_info.block_stl_fpath))
            {
                XElement target = x_setup.Descendants("Component").Where(e => e.Attribute("Name").Value.ToUpper() == "STOCK" && e.Attribute("Type") != null).First();

                /* Remove template stock from new project */
                remove_components_from_xml(target.Elements(), "The template file may contain a stock, would you like to remove it from the output project?");

                target.Element("Attach").Value = setup_info.block_attach_to;

                target.Add(new XElement("STL",
				new XAttribute("Unit", (ProjectData.PmillUnitsIsInch ? "inch" : "Millimeter")),
                                new XAttribute("Normal", "outward"),
                                new XAttribute("Visible", "on"),
                                new XAttribute("XRGB", "0x000000CC"),
                                new XAttribute("Mirror", "off"),
                                new XAttribute("MirrorCsys", "0"),
                                //new XAttribute("Hue", 1),
                                new XElement("Position",
                                    new XAttribute("X", model_origin.dX),
                                    new XAttribute("Y", model_origin.dY),
                                    new XAttribute("Z", model_origin.dZ)),
                                new XElement("Rotation",
                                    new XAttribute("I", model_origin.dXangle),
                                    new XAttribute("J", model_origin.dYangle),
                                    new XAttribute("K", model_origin.dZangle)),
                                new XElement("File", setup_info.block_stl_fpath)
                                ));
            }
        }

        public static void SetCSysInfo(ref XElement x_setup, SetupInfo setup_info, string cut_stock_csys)
        {
            IEnumerable<XElement> rows = x_setup.Descendants("CSystem");
            XElement lastrow;

            if (rows.Count() == 0)
            {
                IEnumerable<XElement> csystems = x_setup.Descendants("CSystems");
                if (csystems.Count() == 0)
                {
                    x_setup.Add(
                        new XElement("CSystems"),
                        new XAttribute("Simulation", "off"),
                        new XElement("Machine")
                        );
                    csystems = x_setup.Descendants("CSystems");
                }
                csystems.Last().Add(
                    new XElement("CSystem",
                    new XAttribute("Name", "Global"),
                    new XAttribute("Type", "component"),
                    new XAttribute("Visible", "all"),
                    new XElement("Attach", "Attach"),
                    new XElement("Position",
                        new XAttribute("X", 0),
                        new XAttribute("Y", 0),
                        new XAttribute("Z", 0)),
                    new XElement("Rotation",
                        new XAttribute("I", 0),
                        new XAttribute("J", 0),
                        new XAttribute("K", 0))        
                        ));
                rows = x_setup.Descendants("CSystem");
            }
            lastrow = rows.Last();
            lastrow.AddAfterSelf(
                new XElement("CSystem", 
                    new XAttribute("Name", setup_info.attach_workplane),
                    new XAttribute("Type", "component"),
                    new XAttribute("Visible", "all"),
                    new XElement("Attach", "Attach"),
                    new XElement("Position",
                        new XAttribute("X", 0),
                        new XAttribute("Y", 0),
                        new XAttribute("Z", 0)),
                    new XElement("Rotation",
                        new XAttribute("I", 0),
                        new XAttribute("J", 0),
                        new XAttribute("K", 0))
                        ));
            PowerMILLAutomation.ActivateWorkplane(setup_info.attach_workplane);
            foreach (string wp in setup_info.workplanes_to_export)
            {
                if (wp == setup_info.attach_workplane) continue;
                WorkPlaneOrigin orig = MainModule.GetWorkplaneCoords(PowerMILLAutomation.oToken, wp, true, Properties.Resources.IDS_WorkplaneFailedToGet);
                lastrow.AddAfterSelf(
                    new XElement("CSystem",
                        new XAttribute("Name", wp),
                        new XAttribute("Type", "component"),
                        new XAttribute("Visible", "all"),
                        new XAttribute("Transition", (wp.Equals(cut_stock_csys, StringComparison.OrdinalIgnoreCase) ? "on" : "off")),
                        new XElement("Attach", (!wp.Equals(cut_stock_csys) ? setup_info.attach_workplane_to : setup_info.block_attach_to)),
                        new XElement("Position",
                            new XAttribute("X", orig.dX),
                            new XAttribute("Y", orig.dY),
                            new XAttribute("Z", orig.dZ)),
                        new XElement("Rotation",
                            new XAttribute("I", orig.dXangle),
                            new XAttribute("J", orig.dYangle),
                            new XAttribute("K", orig.dZangle))
                            ));
            }
        }

        private static void SetTools(ref XElement x_setup, string output_fpath)
        {
            /* Set block data in the .xml file */
            if (!String.IsNullOrEmpty(output_fpath))
            {
                //IEnumerable<XElement> rows = doc.Descendants("Components");

                XElement target = x_setup.Descendants("ToolMan").First();
                target.Element("Library").Value = output_fpath;
            }
        }

        /// <summary>
        /// Fill the NCProg List
        /// </summary>
        /// <returns></returns>
        /// <remarks></remarks>
        private static bool SetNCPrograms(ref XElement x_setup, SetupInfo setup_info, bool tool_id_use_num, bool tool_id_use_name)
        {
            bool bReturn = true;

            XElement target = x_setup.Descendants("NCPrograms").First();
            target.Attribute("Type").Value = "gcode";
            //Ask to remove toolpaths existing in the template
            if (target.Descendants("NCProgram") != null &&
                Messages.ShowWarning(Properties.Resources.IDS_RemoveNCCode) == System.Windows.MessageBoxResult.Yes)
                target.Descendants("NCProgram").Remove();

	    if (target.Attribute("Change") == null)
		target.Add(new XAttribute("Change", "tool_num"));
	    else
		target.Attribute("Change").Value = "tool_num";

        foreach (NCProgramInfo nc_prog in setup_info.nc_progs)
            {
                if (String.IsNullOrEmpty(nc_prog.sPath)) continue;
                target.Add(new XElement("NCProgram",
                                new XAttribute("Use", "on"),
                                new XAttribute("Filter", "off"),
                                new XElement("File", nc_prog.sPath),
                                new XElement("Orient", "None"),
                                new XElement("Ident", "")));
            }

            target = x_setup.Descendants("ToolChange").First();
            if (target.Attribute("List") == null)
                target.Add(new XAttribute("List", "tool_num"));
            else
                target.Attribute("List").Value = "tool_num";
            target.RemoveAll();
            for (int nc_i = 0; nc_i < setup_info.nc_progs.Count; nc_i++)
            {
                if (setup_info.nc_progs[nc_i].oTools == null) continue;
                foreach (PowerMILLExporter.Tools.NCProgramTool tool in setup_info.nc_progs[nc_i].oTools)
                {
                    target.Add(new XElement("Event",
                                new XAttribute("NCProgram", nc_i + 1),
                                new XAttribute("Filter", "on"),
                                new XAttribute("Init", 1),
                                new XElement("Cutter",
                                    new XAttribute("Ident", WriteTools.ConstructToolName(tool.Number, tool.Name, tool_id_use_num, tool_id_use_name)),
                                    "1:" + tool.Number),
                                new XElement("Holder",
                                    new XAttribute("Ident", WriteTools.ConstructToolName(tool.Number, tool.Name, tool_id_use_num, tool_id_use_name))),
                                new XElement("Tool",
                                    new XAttribute("Use", "other"),
                                    new XAttribute("Teeth", 0),
                                    new XAttribute("File", setup_info.nc_progs[nc_i].sPath),
                                    new XAttribute("Line", 0))
                                    ));
                }

                if (x_setup.Descendants("Tool").Where(e => e.Attribute("CutterVisible") != null) == null)
                {
                    x_setup.Add(new XElement("Tool",
                                new XAttribute("Change", "tool_num"),
                                new XAttribute("CutterVisible", "all"),
                                new XAttribute("HolderVisible", "all"),
                                new XAttribute("Draw", "shaded"),
                                new XAttribute("Control", "tip"),
                                new XAttribute("HolderStockCollision", "on"),
                                new XAttribute("HolderFixtureCollision", "on"),
                                new XAttribute("CutterFixtureCollision", "on"),
                                new XAttribute("HolderStockNearMiss", 0),
                                new XAttribute("HolderFixtureNearMiss", 0),
                                new XAttribute("CutterFixtureNearMiss", 0),
                                new XAttribute("Gouge", "off"),
                                new XAttribute("HolderStockCollision", "on")));
                }
                else
                {
                    IEnumerable<XElement> cutters = x_setup.Descendants("Tool").Where(e => e.Attribute("CutterVisible") != null);
                    if (cutters.Count() != 0) {
                        target = cutters.First();
		                if (tool_id_use_name)
		                {
			                if (target.Attribute("Change") == null)
                                target.Add(new XAttribute("Change", "tool_num"));
			                else
                                target.Attribute("Change").Value = "tool_num";
		                }
                    }
                }

                //<Event NCProgram="1" Filter="on" Init="1"> //LINE DONE
                //  <Cutter Ident="Main_Upper_01">1:1</Cutter>
                //  <Holder Ident="Main_Upper_01"/>
                //  <Tool Use="other" Teeth="0" File="C:\VERICUT\33B3262_R_V2_NTJX\33B3262_R_V2_NTJX.mcd" Line="0"></Tool>
                //</Event>
            }
            return bReturn;
        }

    }
}
