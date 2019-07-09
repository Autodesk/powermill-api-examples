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
using System.Windows;
using System.Windows.Controls;
using System.IO;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace DrillingAutomation
{
    public class HoleInfos
    {
        public int HoleNumber { get; set; }
        public double Diameter { get; set; }
        public double Depth { get; set; }
        public double Draft { get; set; }
        public int ComponentID { get; set; }
    }

    public class MachineCells
    {
        public string Name { get; set; }
        public string CSVPath { get; set; }
        public string ToolpathsPath { get; set; }
        public double Tolerance { get; set; }
        public bool UseMethod { get; set; }
        public bool DepthFromTop { get; set; }
        public bool AllowTagDuplicate { get; set; }
    }

    public class NewHoleInfos
    {
        public int HoleNumber { get; set; }
        public double UpperDiameter { get; set; }
        public double LowerDiameter { get; set; }
        public double Depth { get; set; }
        public double RColor { get; set; }
        public double GColor { get; set; }
        public double BColor { get; set; }
        public int ComponentID { get; set; }
    }

    public class DBHoleInfos
    {
        public string HoleTag { get; set; }
        public string HoleDescription { get; set; }
        public string Family { get; set; }
        public double UpperDiameter { get; set; }
        public double LowerDiameter { get; set; }
        public double MaximumDepth { get; set; }
        public double MinimumDepth { get; set; }
        public double HoleDepth { get; set; }
        public double RColor { get; set; }
        public double GColor { get; set; }
        public double BColor { get; set; }
        public bool LastComponent { get; set; }
    }

    public class RecognizedHoles
    {
        public int HoleNumber { get; set; }
        public string HoleTag { get; set; }
        public string HoleDescription { get; set; }
    }

    public class TaggedHoles
    {
        public string HoleName { get; set; }
        public string HoleTag { get; set; }
    }

    public class DBHoleExport
    {
        public string Tag { get; set; }
        public string Description { get; set; }
        public string Family { get; set; }
        public double UpperDiameter { get; set; }
        public double LowerDiameter { get; set; }
        public double MinDepth { get; set; }
        public double MaxDepth { get; set; }
        public double Depth { get; set; }
        public double RColor { get; set; }
        public double GColor { get; set; }
        public double BColor { get; set; }
        public int ComponentID { get; set; }
    }
    class TagHoles
    {
        public static void GetINIPath(out string INIPath)
        {
            INIPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\DrillingAutomation\\Settings.ini";
            if (!File.Exists(INIPath))
            {
                if (!Directory.Exists(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\DrillingAutomation\\"))
                {
                    Directory.CreateDirectory(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) + "\\DrillingAutomation\\");
                }
                MessageBox.Show("Make sure you select your database file and set your toolpath folder first", "Error");
                return;
            }
        }
        public static void GetActiveMachineCellInfos(string CurrentMachine, List<MachineCells> MachineCells, out string CSVPath, out string TemplateFolder, out double Tolerance, out bool UseMethod, out bool DepthFromTop, out bool AllowTagDuplicate)
        {
            CSVPath = "";
            TemplateFolder = "";
            Tolerance = 0;
            UseMethod = false;
            DepthFromTop = false;
            AllowTagDuplicate = true;
            foreach (MachineCells Machine in MachineCells)
            {
                if (Machine.Name == CurrentMachine)
                {
                    TemplateFolder = Machine.ToolpathsPath;
                    CSVPath = Machine.CSVPath;
                    Tolerance = Machine.Tolerance;
                    if (Machine.AllowTagDuplicate)
                    {
                        AllowTagDuplicate = true;
                    }
                    else
                    {
                        AllowTagDuplicate = false;
                    }
                    if (Machine.UseMethod)
                    {
                        UseMethod = true;
                    }
                    else
                    {
                        UseMethod = false;
                    }
                    if (Machine.DepthFromTop)
                    {
                        DepthFromTop = true;
                    }
                    else
                    {
                        DepthFromTop = false;
                    }
                }
            }
        }
        public static void ExtractINIData(out List<MachineCells> MachineCells)
        {
            string CSVPath = "";
            string TemplateFolder = "";
            double Tolerance = 0;
            string MachineCell = "";
            bool UseMethod = false;
            bool FoundMethod = false;
            bool AllowTagDuplicate = true;
            MachineCells = new List<MachineCells>();

            GetINIPath(out string Path);

            if (File.Exists(Path))
            {
                const Int32 BufferSize = 128;
                using (var fileStream = File.OpenRead(Path))
                using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
                {
                    String line;
                    while ((line = streamReader.ReadLine()) != null)
                    {
                        if (line.IndexOf("Database=") >= 0)
                        {
                            CSVPath = line.Substring(9, line.Length - 9);
                        }
                        else if (line.IndexOf("ToolpathFolder=") >= 0)
                        {
                            TemplateFolder = line.Substring(15, line.Length - 15);
                        }
                        else if (line.IndexOf("Tolerance=") >= 0)
                        {
                            Tolerance = double.Parse(line.Substring(10, line.Length - 10));
                        }
                        else if (line.IndexOf("AllowTagDuplicate=True") >= 0)
                        {
                            AllowTagDuplicate = true;
                        }
                        else if (line.IndexOf("AllowTagDuplicate=False") >= 0)
                        {
                            AllowTagDuplicate = false;
                        }
                        else if (line.IndexOf("ToolpathORMethods=Methods") >= 0)
                        {
                            UseMethod = true;
                            FoundMethod = true;
                        }
                        else if (line.IndexOf("ToolpathORMethods=Toolpaths") >= 0)
                        {
                            UseMethod = false;
                            FoundMethod = true;
                        }
                        else if (line.IndexOf("DepthReference=TopHole") >= 0)
                        {
                            MachineCells.Add(new MachineCells
                            {
                                Name = MachineCell,
                                Tolerance = Tolerance,
                                ToolpathsPath = TemplateFolder,
                                CSVPath = CSVPath,
                                UseMethod = UseMethod,
                                AllowTagDuplicate = AllowTagDuplicate,
                                DepthFromTop = true
                            });
                            FoundMethod = false;
                        }
                        else if (line.IndexOf("DepthReference=TopComponent") >= 0)
                        {
                            MachineCells.Add(new MachineCells
                            {
                                Name = MachineCell,
                                Tolerance = Tolerance,
                                ToolpathsPath = TemplateFolder,
                                CSVPath = CSVPath,
                                UseMethod = UseMethod,
                                AllowTagDuplicate = AllowTagDuplicate,
                                DepthFromTop = false
                            });
                            FoundMethod = false;
                        }
                        else if (line.IndexOf("[") == 0)
                        {
                            MachineCell = line.Substring(1, line.Length - 2);
                        }
                        if (line.IndexOf("DepthReference=TopComponent") >= 0)
                        {
                            MachineCells.Add(new MachineCells
                            {
                                Name = MachineCell,
                                Tolerance = Tolerance,
                                ToolpathsPath = TemplateFolder,
                                CSVPath = CSVPath,
                                UseMethod = UseMethod,
                                AllowTagDuplicate = AllowTagDuplicate,
                                DepthFromTop = false
                            });
                            FoundMethod = false;
                        }
                    }
                }
                if (FoundMethod && MachineCells.Count == 0)
                {
                    MachineCells.Add(new MachineCells
                    {
                        Name = MachineCell,
                        Tolerance = Tolerance,
                        ToolpathsPath = TemplateFolder,
                        CSVPath = CSVPath,
                        UseMethod = UseMethod,
                        AllowTagDuplicate = AllowTagDuplicate,
                        DepthFromTop = false
                    });
                    FoundMethod = false;
                }
            }

        }
        public static void ExtractFeatureSetData(string CSVDatabase, bool DepthFromTop, out List<NewHoleInfos> HoleData, out List<DBHoleInfos> DBHoleData)
        {

            PowerMILLAutomation.ExecuteEx("DIALOGS MESSAGE OFF");
            PowerMILLAutomation.ExecuteEx("DIALOGS ERROR OFF");
            PowerMILLAutomation.ExecuteEx("STRING PMILLVAR = \"\"");
            PowerMILLAutomation.ExecuteEx("$PMILLVAR = \"extract(components(this), 'diameter')\"");
            string Diameters = PowerMILLAutomation.ExecuteEx("print par \"to_xml(apply(components(entity('featureset', '')), PMILLVAR))\"");
            string[] DiameterList = Regex.Split(Diameters, Environment.NewLine);

            PowerMILLAutomation.ExecuteEx("$PMILLVAR = \"extract(components(this), 'depth')\"");
            string Depths = PowerMILLAutomation.ExecuteEx("print par \"to_xml(apply(components(entity('featureset', '')), PMILLVAR))\"");
            string[] DepthList = Regex.Split(Depths, Environment.NewLine);

            PowerMILLAutomation.ExecuteEx("$PMILLVAR = \"extract(components(this), 'draft')\"");
            string Drafts = PowerMILLAutomation.ExecuteEx("print par \"to_xml(apply(components(entity('featureset', '')), PMILLVAR))\"");
            string[] DraftList = Regex.Split(Drafts, Environment.NewLine);

            PowerMILLAutomation.ExecuteEx("$PMILLVAR = \"extract(components(this), 'colour')\"");
            string Colors = PowerMILLAutomation.ExecuteEx("print par \"to_xml(apply(components(entity('featureset', '')), PMILLVAR))\"");
            string[] ColorList = Regex.Split(Colors, Environment.NewLine);

            //ExtractXMLValues(DiameterList,DepthList,DraftList, out HoleData);
            //ExtractDatabaseData(CSVDatabase, out DBHoleData);
            ExtractXMLValuesNew(DiameterList, DepthList, DraftList, ColorList, DepthFromTop, out HoleData);
            ExtractNewDatabaseData(CSVDatabase, out DBHoleData);

            PowerMILLAutomation.ExecuteEx("DIALOGS MESSAGE ON");
            PowerMILLAutomation.ExecuteEx("DIALOGS ERROR ON");
        }

        public static void ExtractNewDatabaseData(string DatabaseFile, out List<DBHoleInfos> DBHoleData)
        {
            DBHoleData = new List<DBHoleInfos>();
            if (File.Exists(DatabaseFile))
            {
                const Int32 BufferSize = 128;
                using (var fileStream = File.OpenRead(DatabaseFile))
                using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
                {
                    String Line;
                    while ((Line = streamReader.ReadLine()) != null)
                    {
                        if (Line != "")
                        {
                            if (Line.Substring(0, 1) != "*" && Line.IndexOf("Family=") < 0)
                            {
                                // Split the Line
                                string[] CurrentHole = Regex.Split(Line, ";");
                                string[] CurrentHoleList = Regex.Split(CurrentHole[0], ",");
                                double MinDepth = 0;
                                double MaxDepth = 0;
                                double HoleDepth = 0;
                                double UpperDiameter = 0;
                                double LowerDiameter = 0;
                                double ColorR = -1;
                                double ColorG = -1;
                                double ColorB = -1;
                                int ColorStartIndex = 0;
                                int ColorLength = 0;
                                string TempLine = "";
                                string temp = "";
                                string Family = "";

                                // Add NC prog to the List
                                for (int i = 0; i <= CurrentHoleList.Length - 1; i++)
                                {
                                    TempLine = CurrentHoleList[i];
                                    UpperDiameter = double.Parse(TempLine.Substring(TempLine.IndexOf("UD=") + 3, TempLine.IndexOf(" ") - 3));

                                    TempLine = TempLine.Substring(TempLine.IndexOf(" ")+1);
                                    LowerDiameter = double.Parse(TempLine.Substring(TempLine.IndexOf("LD=") + 3, TempLine.IndexOf(" ") - 3));

                                    TempLine = TempLine.Substring(TempLine.IndexOf(" ")+1);
                                    if (TempLine.LastIndexOf(" ") >= 0)
                                    {
                                        HoleDepth = double.Parse(TempLine.Substring(TempLine.IndexOf("Depth=") + 6, TempLine.IndexOf(" ") - 6));
                                    }
                                    else
                                    {
                                        HoleDepth = double.Parse(TempLine.Substring(TempLine.IndexOf("Depth=") + 6, TempLine.Length - 6));
                                    }
                                    
                                    if (TempLine.IndexOf("MinDepth=") >= 0 && TempLine.IndexOf("MaxDepth=") >= 0)
                                    {
                                        TempLine = TempLine.Substring(TempLine.IndexOf(" ")+1);
                                        MinDepth = double.Parse(TempLine.Substring(TempLine.IndexOf("MinDepth=") + 9, TempLine.IndexOf(" ") - 9));

                                        TempLine = TempLine.Substring(TempLine.IndexOf(" ")+1);

                                        if (TempLine.LastIndexOf(" ") >= 0)
                                        {
                                            temp = TempLine.Substring(TempLine.IndexOf("MaxDepth=") + 9, TempLine.IndexOf(" ") - 9);
                                            MaxDepth = double.Parse(TempLine.Substring(TempLine.IndexOf("MaxDepth=") + 9, TempLine.IndexOf(" ") - 9));
                                        }
                                        else
                                        {
                                            temp = TempLine.Substring(TempLine.IndexOf("MaxDepth=") + 9, TempLine.Length - 9);
                                            MaxDepth = double.Parse(TempLine.Substring(TempLine.IndexOf("MaxDepth=") + 9, TempLine.Length - 9));
                                        }
                                    }
                                    else if (TempLine.IndexOf("MinDepth=") >= 0 || TempLine.IndexOf("MaxDepth=") >= 0)
                                    {
                                        MessageBox.Show("Missing Min or Max depth on the entry", "Error");
                                        MaxDepth = -1;
                                        MinDepth = -1;
                                    }
                                    else
                                    {
                                        MinDepth = 0;
                                        MaxDepth = 0;
                                    }
                                    if (TempLine.IndexOf("Color=R") >= 0)
                                    {
                                        ColorStartIndex = TempLine.IndexOf("Color=R")+7;
                                        ColorLength = TempLine.IndexOf("G", ColorStartIndex)- ColorStartIndex;
                                        ColorR = double.Parse(TempLine.Substring(ColorStartIndex, ColorLength));

                                        TempLine = TempLine.Substring(ColorStartIndex + ColorLength+1);
                                        ColorStartIndex = 0;
                                        ColorLength = TempLine.IndexOf("B", ColorStartIndex);
                                        ColorG = double.Parse(TempLine.Substring(ColorStartIndex, ColorLength));
                                        ColorB = double.Parse(TempLine.Substring(ColorStartIndex + ColorLength+1));
                                    }

                                        if (CurrentHole.Count() == 4)
                                    {
                                        if (CurrentHole[3] != "")
                                        {
                                            Family = CurrentHole[3];
                                        }
                                        else
                                        {
                                            Family = "Unknown";
                                        }
                                    }
                                    else
                                    {
                                        Family = "Unknown";
                                    }
                                    if (i == CurrentHoleList.Length - 1)
                                    {
                                        DBHoleData.Add(new DBHoleInfos
                                        {
                                            HoleTag = CurrentHole[1],
                                            HoleDescription = CurrentHole[2],
                                            Family = Family,
                                            UpperDiameter = UpperDiameter,
                                            LowerDiameter = LowerDiameter,
                                            MaximumDepth = MaxDepth,
                                            MinimumDepth = MinDepth,
                                            HoleDepth = HoleDepth,
                                            RColor = ColorR,
                                            GColor = ColorG,
                                            BColor = ColorB,
                                            LastComponent = true
                                        });
                                    }
                                    else
                                    {
                                        DBHoleData.Add(new DBHoleInfos
                                        {
                                            HoleTag = CurrentHole[1],
                                            HoleDescription = CurrentHole[2],
                                            Family = Family,
                                            UpperDiameter = UpperDiameter,
                                            LowerDiameter = LowerDiameter,
                                            MaximumDepth = MaxDepth,
                                            MinimumDepth = MinDepth,
                                            HoleDepth = HoleDepth,
                                            RColor = ColorR,
                                            GColor = ColorG,
                                            BColor = ColorB,
                                            LastComponent = false
                                        });
                                    }

                                }
                            }
                        }
                    }
                }
                //Add a dummy component to the database to avoid single components holes to be left aside if they are the last hole in the csv file.
                DBHoleData.Add(new DBHoleInfos
                {
                    HoleTag = "",
                    HoleDescription = "",
                    Family = "",
                    UpperDiameter = -1,
                    LowerDiameter = -1,
                    MaximumDepth = -1,
                    MinimumDepth = -1,
                    HoleDepth = -1,
                    RColor = -1,
                    GColor = -1,
                    BColor = -1,
                    LastComponent = true
                });
            }
        }

        public static void MatchAndRenameHolesNew(string Featureset, List<NewHoleInfos> HoleData, List<DBHoleInfos> DBHoleData, double Tolerance, out List<RecognizedHoles> RecognizedHolesList, out List<TaggedHoles> UnRecognizedHolesNameList)
        {
            double PreviousHole = 0;
            List<DBHoleInfos> CurrentDBHoleData = new List<DBHoleInfos>();
            List<NewHoleInfos> CurrentHoleData = new List<NewHoleInfos>();
            RecognizedHolesList = new List<RecognizedHoles>();
            List<RecognizedHoles> UnRecognizedHolesList = new List<RecognizedHoles>();
            UnRecognizedHolesNameList = new List<TaggedHoles>();
            bool FirstHole = true;
            int PMHoleID = 0;
            int DBHoleID = 0;
            bool HoleMatched = false;

            PowerMILLAutomation.ExecuteEx("GRAPHICS LOCK");

            foreach (NewHoleInfos PMHole in HoleData)
            {
                PMHoleID = PMHoleID + 1;
                if (FirstHole)
                {
                    CurrentHoleData.Add(new NewHoleInfos
                    {
                        UpperDiameter = PMHole.UpperDiameter,
                        LowerDiameter = PMHole.LowerDiameter,
                        Depth = PMHole.Depth,
                        RColor = PMHole.RColor,
                        GColor = PMHole.GColor,
                        BColor = PMHole.BColor,
                        HoleNumber = PMHole.HoleNumber
                    });
                    FirstHole = false;
                    PreviousHole = PMHole.HoleNumber;
                }
                else
                {
                    //Loop here until the last component of the hole is find to build a list for one hole only
                    if (PMHole.HoleNumber == PreviousHole)
                    {
                        CurrentHoleData.Add(new NewHoleInfos
                        {
                            UpperDiameter = PMHole.UpperDiameter,
                            LowerDiameter = PMHole.LowerDiameter,
                            Depth = PMHole.Depth,
                            RColor = PMHole.RColor,
                            GColor = PMHole.GColor,
                            BColor = PMHole.BColor,
                            HoleNumber = PMHole.HoleNumber
                        });
                        PreviousHole = PMHole.HoleNumber;
                    }
                    if (PMHole.HoleNumber != PreviousHole || PMHoleID == HoleData.Count)
                    {
                        //Go here once the next hole is found so we can see if it matches with the current csv file hole
                        DBHoleID = 0;
                        PreviousHole = 0;
                        CurrentDBHoleData.Clear();
                        foreach (DBHoleInfos DBHole in DBHoleData)
                        {
                            DBHoleID = DBHoleID + 1;

                            CurrentDBHoleData.Add(new DBHoleInfos
                            {
                                HoleTag = DBHole.HoleTag,
                                HoleDescription = DBHole.HoleDescription,
                                UpperDiameter = DBHole.UpperDiameter,
                                LowerDiameter = DBHole.LowerDiameter,
                                MaximumDepth = DBHole.MaximumDepth,
                                MinimumDepth = DBHole.MinimumDepth,
                                RColor = DBHole.RColor,
                                GColor = DBHole.GColor,
                                BColor = DBHole.BColor,
                                HoleDepth = DBHole.HoleDepth,
                                LastComponent = DBHole.LastComponent
                            });
                            if (DBHole.LastComponent)
                            {
                                if (CurrentDBHoleData.Count == CurrentHoleData.Count)
                                {
                                    for (int i = 0; i <= CurrentDBHoleData.Count - 1; i++)
                                    {
                                        if (CurrentDBHoleData[i].MaximumDepth > 0)
                                        {
                                            if ((CurrentHoleData[i].UpperDiameter + Tolerance) > CurrentDBHoleData[i].UpperDiameter && (CurrentHoleData[i].UpperDiameter - Tolerance) < CurrentDBHoleData[i].UpperDiameter && (CurrentHoleData[i].LowerDiameter + Tolerance) > CurrentDBHoleData[i].LowerDiameter && (CurrentHoleData[i].LowerDiameter - Tolerance) < CurrentDBHoleData[i].LowerDiameter && CurrentHoleData[i].Depth <= CurrentDBHoleData[i].MaximumDepth && CurrentHoleData[i].Depth >= CurrentDBHoleData[i].MinimumDepth)
                                            {
                                                if (CurrentDBHoleData[i].RColor < 0)
                                                {
                                                    if (i == CurrentDBHoleData.Count - 1)
                                                    {
                                                        RecognizedHolesList.Add(new RecognizedHoles
                                                        {
                                                            HoleNumber = CurrentHoleData[i].HoleNumber,
                                                            HoleTag = CurrentDBHoleData[i].HoleTag,
                                                            HoleDescription = CurrentDBHoleData[i].HoleDescription
                                                        });
                                                        HoleMatched = true;
                                                    }
                                                    //else
                                                    //{
                                                    //    break;
                                                    //}
                                                }
                                                else
                                                {
                                                    if (CurrentHoleData[i].RColor == CurrentDBHoleData[i].RColor && CurrentHoleData[i].GColor == CurrentDBHoleData[i].GColor && CurrentHoleData[i].BColor == CurrentDBHoleData[i].BColor)
                                                    {
                                                        if (i == CurrentDBHoleData.Count - 1)
                                                        {
                                                            RecognizedHolesList.Add(new RecognizedHoles
                                                            {
                                                                HoleNumber = CurrentHoleData[i].HoleNumber,
                                                                HoleTag = CurrentDBHoleData[i].HoleTag,
                                                                HoleDescription = CurrentDBHoleData[i].HoleDescription
                                                            });
                                                            HoleMatched = true;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        break;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                        else
                                        {
                                            if ((CurrentHoleData[i].UpperDiameter + Tolerance) > CurrentDBHoleData[i].UpperDiameter && (CurrentHoleData[i].UpperDiameter - Tolerance) < CurrentDBHoleData[i].UpperDiameter && (CurrentHoleData[i].LowerDiameter + Tolerance) > CurrentDBHoleData[i].LowerDiameter && (CurrentHoleData[i].LowerDiameter - Tolerance) < CurrentDBHoleData[i].LowerDiameter)
                                            {
                                                if (CurrentDBHoleData[i].RColor < 0)
                                                {
                                                    if (i == CurrentDBHoleData.Count - 1)
                                                    {
                                                        RecognizedHolesList.Add(new RecognizedHoles
                                                        {
                                                            HoleNumber = CurrentHoleData[i].HoleNumber,
                                                            HoleTag = CurrentDBHoleData[i].HoleTag,
                                                            HoleDescription = CurrentDBHoleData[i].HoleDescription
                                                        });
                                                        HoleMatched = true;
                                                    }
                                                }
                                                else
                                                {
                                                    if (CurrentHoleData[i].RColor == CurrentDBHoleData[i].RColor && CurrentHoleData[i].GColor == CurrentDBHoleData[i].GColor && CurrentHoleData[i].BColor == CurrentDBHoleData[i].BColor)
                                                    {
                                                        if (i == CurrentDBHoleData.Count - 1)
                                                        {
                                                            RecognizedHolesList.Add(new RecognizedHoles
                                                            {
                                                                HoleNumber = CurrentHoleData[i].HoleNumber,
                                                                HoleTag = CurrentDBHoleData[i].HoleTag,
                                                                HoleDescription = CurrentDBHoleData[i].HoleDescription
                                                            });
                                                            HoleMatched = true;
                                                        }
                                                    }
                                                    else
                                                    {
                                                        break;
                                                    }
                                                }

                                            }
                                            else
                                            {
                                                break;
                                            }
                                        }
                                    }
                                }
                                CurrentDBHoleData.Clear();
                            }
                            if (HoleMatched)
                            {
                                break;
                            }
                        }

                        if (!HoleMatched)
                        {
                            UnRecognizedHolesList.Add(new RecognizedHoles
                            {
                                HoleNumber = CurrentHoleData[CurrentHoleData.Count - 1].HoleNumber,
                                HoleTag = "Unrecognized",
                                HoleDescription = "Unrecognized"
                            });
                        }
                        //Start a small list for the next sets of components
                        CurrentHoleData.Clear();
                        CurrentHoleData.Add(new NewHoleInfos
                        {
                            UpperDiameter = PMHole.UpperDiameter,
                            LowerDiameter = PMHole.LowerDiameter,
                            Depth = PMHole.Depth,
                            RColor = PMHole.RColor,
                            GColor = PMHole.GColor,
                            BColor = PMHole.BColor,
                            HoleNumber = PMHole.HoleNumber
                        });
                        PreviousHole = PMHole.HoleNumber;
                        HoleMatched = false;
                    }
                }
            }
            PMHoleID = 1;
            foreach (RecognizedHoles Hole in RecognizedHolesList)
            {
                PowerMILLAutomation.ExecuteEx("EDIT FEATURESET '" + Featureset + "' SELECT 'T" + Hole.HoleNumber.ToString() + "'");
                PowerMILLAutomation.ExecuteEx("EDIT FEATURESET '" + Featureset + "' HOLES DESCRIPTION USER '" + Hole.HoleDescription + "'");
                PowerMILLAutomation.ExecuteEx("EDIT FEATURESET '" + Featureset + "' HOLES DESCRIPTION TYPE '" + Hole.HoleTag + "'");
                PowerMILLAutomation.ExecuteEx("EDIT FEATURESET '" + Featureset + "' RENAME 'T" + Hole.HoleNumber.ToString() + "' '" + Hole.HoleTag + "_" + PMHoleID.ToString() + "'");
                PowerMILLAutomation.ExecuteEx("EDIT FEATURESET '" + Featureset + "' DESELECT '" + Hole.HoleTag + "_" + PMHoleID.ToString() + "'");
                PMHoleID = PMHoleID + 1;
            }

            PMHoleID = 1;
            foreach (RecognizedHoles Hole in UnRecognizedHolesList)
            {
                PowerMILLAutomation.ExecuteEx("EDIT FEATURESET '" + Featureset + "' RENAME 'T" + Hole.HoleNumber.ToString() + "' '" + Hole.HoleTag + "_" + PMHoleID.ToString() + "'");
                UnRecognizedHolesNameList.Add(new TaggedHoles
                {
                    HoleName = Hole.HoleTag + "_" + PMHoleID.ToString(),
                    HoleTag = Hole.HoleNumber.ToString()
                });
                PMHoleID = PMHoleID + 1;
            }
            PowerMILLAutomation.ExecuteEx("GRAPHICS UNLOCK");
        }
        public static void ExtractDiameter(string CurrentHole, bool ExtractDepth, out double Dia, out double MinDepth, out double MaxDepth)
        {
            MinDepth = 0;
            MaxDepth = 10000;
            Dia = 0;
            string TempLine = "";
            if (CurrentHole.IndexOf("|") == CurrentHole.LastIndexOf("|") && CurrentHole.IndexOf("|") >= 0)
            {
                MinDepth = -1;
                MaxDepth = -1;
            } else
            {
                if (CurrentHole.IndexOf("|") > 0)
                {
                    Dia = double.Parse(CurrentHole.Substring(0, CurrentHole.IndexOf("|")));
                    if (ExtractDepth)
                    {
                        TempLine = CurrentHole.Substring(CurrentHole.IndexOf("|") + 1);
                        MinDepth = double.Parse(TempLine.Substring(0, TempLine.IndexOf("|")));
                        MaxDepth = double.Parse(TempLine.Substring(TempLine.IndexOf("|") + 1));
                    }
                }
                else
                {
                    Dia = double.Parse(CurrentHole);
                }
            }

        }

        public static void ExtractColor(string XMLColorLine, out double Color)
        {
            Color = double.Parse(XMLColorLine.Substring(6, XMLColorLine.IndexOf("</") - 6));
            Color = Color * 255;
            Color = Math.Round(Color, 0);
        }
        public static void ExtractXMLValuesNew(string[] DiameterList, string[] DepthList, string[] DraftList, string[] ColorList, bool DepthFromTop, out List<NewHoleInfos> HoleData)
        {
            int HoleID = 0;
            int ComponentID = -1;
            bool Beginning = false;
            int LineIndex = 0;
            int ColorLineIndex = -8;
            string Line2 = "";
            string Line3 = "";
            string Line4 = "";
            double Dpt = 0;
            double UD = 0;
            double LD = 0;
            double Draft = 0;
            double RColor = -1;
            double GColor = -1;
            double BColor = -1;
            double TotalDepth = 0;
            HoleData = new List<NewHoleInfos>();

            foreach (string Line in DiameterList)
            {
                if (Line == "<array>")
                {
                    Beginning = true;
                    ComponentID = -1;
                    ColorLineIndex = ColorLineIndex + 10;
                    TotalDepth = 0;
                }
                else if (Beginning && (Line.IndexOf("<Diameter>") >= 0))
                {
                    HoleID = HoleID + 1;
                    Beginning = false;
                    ComponentID = ComponentID + 1;
                    Draft = 0;

                    Line3 = DraftList[LineIndex];
                    if (Line3.IndexOf("@Range") < 0)
                    {
                        Draft = double.Parse(Line3.Substring(7, Line3.IndexOf("</") - 7));
                    }
                    UD = double.Parse(Line.Substring(10, Line.IndexOf("</") - 10));
                    Line2 = DepthList[LineIndex];
                    Dpt = double.Parse(Line2.Substring(7, Line2.IndexOf("</") - 7));
                    GetLowerDiameter(Draft, UD, Dpt, out LD);

                    if ((ColorLineIndex + 2) < ColorList.Count())
                    {
                        if (ColorList[ColorLineIndex + 2].IndexOf("<real>") >= 0)
                        {
                            ExtractColor(ColorList[ColorLineIndex + 2], out RColor);
                            ExtractColor(ColorList[ColorLineIndex + 3], out GColor);
                            ExtractColor(ColorList[ColorLineIndex + 4], out BColor);
                        }
                    }
                    HoleData.Add(new NewHoleInfos
                    {
                        HoleNumber = HoleID,
                        UpperDiameter = Math.Round(UD, 4),
                        LowerDiameter = Math.Round(LD, 4),
                        Depth = Math.Round(Dpt, 4),
                        RColor = RColor,
                        GColor = GColor,
                        BColor = BColor,
                        ComponentID = ComponentID
                    });

                }
                else if (Line.IndexOf("<Diameter>") >= 0)
                {
                    ComponentID = ComponentID + 1;
                    ColorLineIndex = ColorLineIndex + 8;
                    Line3 = DraftList[LineIndex];
                    Draft = 0;

                    if (Line3.IndexOf("@Range") < 0)
                    {
                        Draft = double.Parse(Line3.Substring(7, Line3.IndexOf("</") - 7));
                    }
                    UD = double.Parse(Line.Substring(10, Line.IndexOf("</Diameter>") - 10));
                    Line2 = DepthList[LineIndex];
                    Line4 = DepthList[LineIndex - 1];
                    Dpt = double.Parse(Line2.Substring(7, Line2.IndexOf("</Depth>") - 7)) - double.Parse(Line4.Substring(7, Line4.IndexOf("</Depth>") - 7));
                    GetLowerDiameter(Draft, UD, Dpt, out LD);

                    if (DepthFromTop)
                    {
                        Dpt = double.Parse(Line2.Substring(7, Line2.IndexOf("</Depth>") - 7));
                    }

                    if ((ColorLineIndex + 2) < ColorList.Count())
                    {
                        if (ColorList[ColorLineIndex + 2].IndexOf("<real>") >= 0)
                        {
                            ExtractColor(ColorList[ColorLineIndex + 2], out RColor);
                            ExtractColor(ColorList[ColorLineIndex + 3], out GColor);
                            ExtractColor(ColorList[ColorLineIndex + 4], out BColor);
                        }
                    }

                    HoleData.Add(new NewHoleInfos
                    {
                        HoleNumber = HoleID,
                        UpperDiameter = Math.Round(UD, 4),
                        LowerDiameter = Math.Round(LD, 4),
                        Depth = Math.Round(Dpt, 4),
                        RColor = RColor,
                        GColor = GColor,
                        BColor = BColor,
                        ComponentID = ComponentID
                    });
                }
                LineIndex = LineIndex + 1;
                
            }

            //Add a dummy hoel to the list so that if the recognized hole only has one component, it's still analyzed in MatchAndRenameHolesNew method
            HoleData.Add(new NewHoleInfos
            {
                HoleNumber = -1,
                UpperDiameter = 0,
                LowerDiameter = 0,
                Depth = 0,
                RColor = -1,
                GColor = -1,
                BColor = -1,
                ComponentID = 0
            });


        }
        public static void GetLowerDiameter(double Draft, double UpperDiameter, double Depth, out double LowerDiameter)
        {
            LowerDiameter = 0;
            double Diff = Depth * Math.Tan((Draft*Math.PI)/180);
            LowerDiameter = UpperDiameter - (2 * Diff);
        }
        public static void GenerateToolpaths(string MacroPath, string TemplateFolder, bool UseMethod)
        {
            string Project_Path = PowerMILLAutomation.ExecuteEx("print $project_pathname(0)");
            string HoleNamesFile = Project_Path + "\\HoleNames.txt";
            string TPTemplate = "";
            string TPName = "";
            bool Missing_Toolpaths = false;
            string MissingToolpathsList = "";

            if (Project_Path == "")
            {
                MessageBox.Show("Project must be saved before importing toolpaths", "Warning");
            }
            else
            {
                PowerMILLAutomation.ExecuteEx("DIALOGS MESSAGE OFF");
                PowerMILLAutomation.ExecuteEx("DIALOGS ERROR OFF");

                PowerMILLAutomation.ExecuteEx("MACRO '" + MacroPath + "\\CreateList.mac" + "'");
                GetListOfHoles(HoleNamesFile, out List<TaggedHoles> TaggedHolesList, out List<TaggedHoles> FullList);

                PowerMILLAutomation.ExecuteEx("CREATE FOLDER 'Toolpath' Created_Holes");

                if (UseMethod)
                {
                    PowerMILLAutomation.ExecuteEx("CREATE FOLDER 'Toolpath' Temp_Methods");
                    foreach (TaggedHoles Hole in TaggedHolesList)
                    {
                        TPTemplate = TemplateFolder + "\\" + Hole.HoleTag + ".ptf";
                        PowerMILLAutomation.ExecuteEx("ACTIVATE FOLDER 'Toolpath\\Temp_Methods'");
                        PowerMILLAutomation.ExecuteEx("IMPORT TEMPLATE ENTITY TOOLPATH '" + TPTemplate + "'");
                        PowerMILLAutomation.ExecuteEx("CREATE FOLDER 'Toolpath\\Created_Holes' " + Hole.HoleTag);
                        PowerMILLAutomation.ExecuteEx("ACTIVATE FOLDER #");
                        PowerMILLAutomation.ExecuteEx("MACRO '" + MacroPath + "\\ImportMethod.mac" + "'");
                    }
                    PowerMILLAutomation.ExecuteEx("DELETE TOOLPATH FOLDER 'Toolpath\\Temp_Methods'");
                }
                else
                {
                    foreach (TaggedHoles Tags in TaggedHolesList)
                    {
                        PowerMILLAutomation.ExecuteEx("EDIT FEATURESET ; DESELECT ALL");
                        foreach (TaggedHoles Hole in FullList)
                        {
                            if (Tags.HoleTag == Hole.HoleTag)
                            {
                                PowerMILLAutomation.ExecuteEx("EDIT FEATURESET ; SELECT '" + Hole.HoleName + "'");
                            }
                        }
                        //Import toolpaths
                        if (!Directory.Exists(TemplateFolder + @"\" + Tags.HoleTag))
                        {
                            MissingToolpathsList = MissingToolpathsList + Environment.NewLine + Tags.HoleTag;
                            Missing_Toolpaths = true;
                        }
                        else
                        {
                            string[] Files = Directory.GetFiles(TemplateFolder + @"\" + Tags.HoleTag, "*.ptf", SearchOption.TopDirectoryOnly);
                            PowerMILLAutomation.ExecuteEx("CREATE FOLDER 'Toolpath\\Created_Holes' " + Tags.HoleTag);
                            PowerMILLAutomation.ExecuteEx("ACTIVATE FOLDER #");

                            foreach (string file in Files)
                            {
                                TPName = file.Substring(file.LastIndexOf("\\") + 1, file.Length - file.LastIndexOf("\\") - 5);
                                PowerMILLAutomation.ExecuteEx("IMPORT TEMPLATE ENTITY TOOLPATH '" + file + "'");
                                PowerMILLAutomation.ExecuteEx("RENAME Toolpath # '" + TPName + "'");
                                PowerMILLAutomation.ExecuteEx("ACTIVATE Toolpath '" + TPName + "'");
                                PowerMILLAutomation.ExecuteEx("MACRO '" + MacroPath + "\\ImportToolpath.mac" + "'");
                                PowerMILLAutomation.ExecuteEx("EDIT TOOLPATH ; CALCULATE");
                            }
                        }
                    }
                }
                if (Missing_Toolpaths)
                {
                    MessageBox.Show("No toolpath folder found for:" + Environment.NewLine + MissingToolpathsList, "Warning!");
                }
                PowerMILLAutomation.ExecuteEx("DIALOGS MESSAGE ON");
                PowerMILLAutomation.ExecuteEx("DIALOGS ERROR ON");
            }
        }

        public static void GetListOfHoles(string FileName, out List<TaggedHoles> TaggedHolesList, out List<TaggedHoles> FullList)
        {
            string HoleTag = "";
            TaggedHolesList = new List<TaggedHoles>();
            FullList = new List<TaggedHoles>();
            if (File.Exists(FileName))
            {
                const Int32 BufferSize = 128;
                using (var fileStream = File.OpenRead(FileName))
                using (var streamReader = new StreamReader(fileStream, Encoding.UTF8, true, BufferSize))
                {
                    String Line;
                    while ((Line = streamReader.ReadLine()) != null)
                    {
                        if (Line != "" && Line.IndexOf("Unrecognized") < 0)
                        {
                            HoleTag = Line.Substring(0, Line.LastIndexOf("_"));
                            if (!TaggedHolesList.Any(n => n.HoleTag == HoleTag))
                            {
                                TaggedHolesList.Add(new TaggedHoles
                                {
                                    HoleName = Line,
                                    HoleTag = HoleTag
                                });
                            }
                            FullList.Add(new TaggedHoles
                            {
                                HoleName = Line,
                                HoleTag = HoleTag
                            });
                        }
                    }
                }
            }
        }
    }
}
