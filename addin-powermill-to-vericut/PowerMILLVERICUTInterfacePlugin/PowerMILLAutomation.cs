// -----------------------------------------------------------------------
// Copyright 2018 Autodesk, Inc. All rights reserved.
// 
// Use of this software is subject to the terms of the Autodesk license
// agreement provided at the time of installation or download, or which
// otherwise accompanies this software in either electronic or hard copy form.
// -----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using Delcam.Plugins.Framework;

namespace PowerMILLVericutInterfacePlugin
{
   public class PowerMILLAutomation
    {
        // variables
        private PowerMILL.PluginServices oPServices;
        private string oToken;
        
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
        public PowerMILLAutomation(IPluginCommunicationsInterface comms)
        {
	        oPServices = comms.Services;
	        oToken = comms.Token;
        }  

       /// <summary>
       /// This execute a PowerMILL command
       /// </summary>
       /// <param name="Command"></param>
        public void Execute(string Command)
        {
           oPServices.QueueCommand(oToken, Command);
        }

       /// <summary>
       /// This execute a PowerMILL command and return the answer in a string
       /// </summary>
       /// <param name="Command"></param>
       /// <returns></returns>
        public string ExecuteEx(string Command)
        {
            object sResponse;
            int error = oPServices.DoCommandEx(oToken, Command,out sResponse);
            if (error != 0)
            {
                throw new Exception("Failed to obtain command value: " + Command);
            }
            return sResponse as string;
        }

        /// <summary>
        /// Turn off dialogs
        /// </summary>
        public void DialogsOFF()
        {
            Execute("DIALOGS MESSAGE OFF");
            //Execute("DIALOGS ERROR OFF");
        }

        /// <summary>
        /// Turn on dialogs
        /// </summary>
        public void DialogsON()
        {
            Execute("DIALOGS MESSAGE ON");
            //Execute("DIALOGS ERROR ON");
        }

        /// <summary>
        /// Returns a list of PM entity name
        /// </summary>
        public List<string> GetListOf(enumEntity Entity)
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
        public string GetParameterValue(string Parameter)
        {
            return oPServices.GetParameterValue(oToken, Parameter);
        }

        /// <summary>
        /// Returns a parameter (terse)
        /// </summary>
        public string GetParameterValueTerse(string Parameter)
        {
            return oPServices.GetParameterValueTerse(oToken, Parameter);
        }

        /// <summary>
        /// return the PM units
        /// </summary>
        public enumUnit Units
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

    }
}
