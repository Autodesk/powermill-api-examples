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
using System.Windows.Forms;

namespace PowerMILLExporter
{
    public class Utilities
    {
        /// <summary>
        /// Check if Checked NCProg has been modified or if the OutputFile has been delete
        /// </summary>
        /// <param name="NCProgramName"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static string CheckNCProgram(string NCProgramName, out string sStatus)
        {
            string sCNFileName = null;

            // Get the NC prog Status
            //string sStatus = pmill.ExecuteEx("print par \"entity('ncprogram';'" + NCProgramName + "').Status\"").Replace("(ENUM)", "").Trim();
            sStatus = PowerMILLAutomation.GetParameterValueTerse("entity('ncprogram';'" + NCProgramName + "').Status").Trim();
            string sWorkPlane = PowerMILLAutomation.GetParameterValueTerse("entity('ncprogram';'" + NCProgramName + "').OutputWorkplane").Trim();

            if (sStatus.ToLower() == "written")
            {
                // Get NC prog Infos
                PowerMILLAutomation.Execute("DIALOGS MESSAGE OFF");
                string sProgCN = PowerMILLAutomation.ExecuteEx("EDIT NCPROGRAM '" + NCProgramName + "' LIST");
                PowerMILLAutomation.Execute("DIALOGS MESSAGE ON");

                string[] sTabCNInfos = sProgCN.Split((char)13);
                    int iSlashIndex = sTabCNInfos[2].IndexOf('/');
                    if (iSlashIndex < 0) iSlashIndex = 0;
                    int iSpaceIndex = sTabCNInfos[2].LastIndexOf((char)32, iSlashIndex);

                    // Get the NCProg File Path
                    sCNFileName = sTabCNInfos[2].Remove(0, iSpaceIndex).Trim();
                return sCNFileName;
            }
            else
            {
                return null;
            }
        }

        /// <summary>
        /// Put the number in the right format according to the decimal separator
        /// </summary>
        /// <param name="Number">Number to convert in string</param>
        /// <returns></returns>
        /// <remarks></remarks>
        public static double ConvertDecimalSeparator(string Number)
        {
            double dNumber = 0;
            if (Number.Contains(","))
            {
                Number = Number.Replace(",", System.Globalization.CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);
            }
            double.TryParse(Number, out dNumber);
            return dNumber;
        }

        public static double DegreesToRadians(double alpha_deg)
        {
            return Math.PI * alpha_deg / 180.0;
        }

        public static string ToInvariantCultureString(double value)
        {
            return value.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }

        public static double ToInvariantCultureDouble(string value)
        {
            return Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}
