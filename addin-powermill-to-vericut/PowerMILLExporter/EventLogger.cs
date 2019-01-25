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
using System.IO;

namespace PowerMILLExporter
{
    public static class EventLogger
    {
        private static string sLogFPath = "";

        public static void InitializeEventLogger()
        {
            if (!Directory.Exists(Path.Combine(Path.GetTempPath(), "PowerMILLVericutInterface")))
                Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), "PowerMILLVericutInterface"));
            sLogFPath = Path.Combine(Path.GetTempPath(), "PowerMILLVericutInterface", DateTime.Now.ToString("MM_d_yy__HH_mm") + ".txt");
            File.WriteAllText(sLogFPath, DateTime.Now.ToString("HH:mm:ss") + ": Plugin initialized" + Environment.NewLine);
        }

        public static void WriteToEvengLog(string sEntry)
        {
            if (sLogFPath != "")
                File.AppendAllText(sLogFPath, DateTime.Now.ToString("HH:mm:ss") + ": " + sEntry + Environment.NewLine);
        }

        public static void WriteExceptionToEvengLog(string sEntry, string sMethod)
        {
            if (sLogFPath != "")
                File.AppendAllText(sLogFPath, DateTime.Now.ToString("HH:mm:ss") + ": Exception occured in method " + sMethod + Environment.NewLine +
                    sEntry + Environment.NewLine);
        }
    }
}
