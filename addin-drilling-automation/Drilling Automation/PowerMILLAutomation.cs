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
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using Delcam.Plugins.Framework;



namespace DrillingAutomation
{
    public static class PowerMILLAutomation
    {
        // variables
        public static PowerMILL.PluginServices oPServices;
        public static string oToken;

        public static void SetVariables(IPluginCommunicationsInterface comms)
        {
            oPServices = comms.Services;
            oToken = comms.Token;
        }

        public static void Execute(string Command)
        {
            oPServices.QueueCommand(oToken, Command);
        }


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

        public static void DialogsOFF()
        {
            Execute("DIALOGS MESSAGE OFF");
        }

        /// <summary>
        /// Turn on dialogs
        /// </summary>
        public static void DialogsON()
        {
            Execute("DIALOGS MESSAGE ON");
        }

        public enum enumEntity
        {
            MachineTools, NCPrograms, Toolpaths, Tools, Boundaries, Patterns, FeatureSets, Workplanes, Models, MachinableModels, StockModels
        }

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
                case enumEntity.MachinableModels:
                    {
                        foreach (PowerMILL.Model oEntity in oPServices.Project.Models)
                        {
                            string isReference = PowerMILLAutomation.ExecuteEx("print $entity('model';'" + oEntity.Name + "').IsReferenceModel");
                            if (isReference == "0")
                            {
                                oList.Add(oEntity.Name);
                            }
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

        public static string GetParameterValueTerse(string Parameter)
        {
            return oPServices.GetParameterValueTerse(oToken, Parameter);
        }

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
            for (int i = 4; i <= sTabCNInfos.Length - 3; i++)
            {
                if (!string.IsNullOrEmpty(sTabCNInfos[i].Trim()))
                    sToolpathesList.Add(sTabCNInfos[i]);
            }

            return sToolpathesList;
        }

        public static void GetToolpathLimits(string Toolpath, out double MinX, out double MinY, out double MinZ, out double MaxX, out double MaxY, out double MaxZ)
        {
            // Get NC prog Infos
            PowerMILLAutomation.Execute("DIALOGS MESSAGE OFF");
            string sSize = PowerMILLAutomation.ExecuteEx("SIZE TOOLPATH '" + Toolpath + "'");
            PowerMILLAutomation.Execute("DIALOGS MESSAGE ON");

            // Split the Line
            string[] sSizeInfos = Regex.Split(sSize, Environment.NewLine);// sProgCN.Split(Convert.ToChar("\n\r"));

            string[] sMinInfos = sSizeInfos[1].Split(new Char[] { ' '});
            string[] sMaxInfos = sSizeInfos[2].Split(new Char[] { ' '});

            MinX = 0;
            MinY = 0;
            MinZ = 0;

            for (int i = 0; i <= sMinInfos.Length-1; i++)
            {
                if (!string.IsNullOrEmpty(sMinInfos[i].Trim()) && sMinInfos[i].Trim() != "Minimum:")
                {
                    if (MinX == 0)
                    {
                        MinX = double.Parse(sMinInfos[i].Trim());
                    }
                    else if (MinY == 0)
                    {
                        MinY = double.Parse(sMinInfos[i].Trim());
                    }
                    else if (MinZ == 0)
                    {
                        MinZ = double.Parse(sMinInfos[i].Trim());
                    }
                }

            }

            MaxX = 0;
            MaxY = 0;
            MaxZ = 0;

            for (int i = 0; i <= sMaxInfos.Length-1; i++)
            {
                if (!string.IsNullOrEmpty(sMaxInfos[i].Trim()) && sMaxInfos[i].Trim() != "Maximum:")
                {
                    if (MaxX == 0)
                    {
                        MaxX = double.Parse(sMaxInfos[i].Trim());
                    }
                    else if (MaxY == 0)
                    {
                        MaxY = double.Parse(sMaxInfos[i].Trim());
                    }
                    else if (MaxZ == 0)
                    {
                        MaxZ = double.Parse(sMaxInfos[i].Trim());
                    }
                }

            }
           
        }

        public static void GetModelsLimits(List<string> ModelList, out double MinX, out double MinY, out double MinZ, out double MaxX, out double MaxY, out double MaxZ)
        {
            bool FoundX = false;
            bool FoundY = false;

            // Get NC prog Infos
            PowerMILLAutomation.Execute("DIALOGS MESSAGE OFF");

            MinX = 10000;
            MinY = 10000;
            MinZ = 10000;
            MaxX = -10000;
            MaxY = -10000;
            MaxZ = -10000;

            foreach (string Model in ModelList)
            {
                string sSize = PowerMILLAutomation.ExecuteEx("SIZE MODEL '" + Model + "'");

                // Split the Line
                string[] sSizeInfos = Regex.Split(sSize, Environment.NewLine);// sProgCN.Split(Convert.ToChar("\n\r"));

                string[] sMinInfos = sSizeInfos[2].Split(new Char[] { ' ' });
                string[] sMaxInfos = sSizeInfos[3].Split(new Char[] { ' ' });
                FoundX = false;
                FoundY = false;

                for (int i = 0; i <= sMinInfos.Length - 1; i++)
                {
                    if (!string.IsNullOrEmpty(sMinInfos[i].Trim()) && sMinInfos[i].Trim() != "Minimum:")
                    {
                        if (double.Parse(sMinInfos[i].Trim()) < MinX && FoundX == false)
                        {
                            MinX = double.Parse(sMinInfos[i].Trim());
                            FoundX = true;
                        }
                        else if (double.Parse(sMinInfos[i].Trim()) < MinY && FoundX && FoundY == false)
                        {
                            MinY = double.Parse(sMinInfos[i].Trim());
                            FoundY = true;
                        }
                        else if (double.Parse(sMinInfos[i].Trim()) < MinZ && FoundX && FoundY)
                        {
                            MinZ = double.Parse(sMinInfos[i].Trim());
                        }
                    }

                }

                FoundX = false;
                FoundY = false;
                for (int i = 0; i <= sMaxInfos.Length - 1; i++)
                {
                    if (!string.IsNullOrEmpty(sMaxInfos[i].Trim()) && sMaxInfos[i].Trim() != "Maximum:")
                    {
                        if (double.Parse(sMaxInfos[i].Trim()) > MaxX && FoundX == false)
                        {
                            MaxX = double.Parse(sMaxInfos[i].Trim());
                            FoundX = true;
                        }
                        else if (double.Parse(sMaxInfos[i].Trim()) > MaxY && FoundX && FoundY == false)
                        {
                            MaxY = double.Parse(sMaxInfos[i].Trim());
                            FoundY = true;
                        }
                        else if (double.Parse(sMaxInfos[i].Trim()) > MaxZ && FoundX && FoundY)
                        {
                            MaxZ = double.Parse(sMaxInfos[i].Trim());
                        }
                    }

                }

            }



            PowerMILLAutomation.Execute("DIALOGS MESSAGE ON");

        }

    }
}
