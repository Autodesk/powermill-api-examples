// -----------------------------------------------------------------------
// Copyright 2018 Autodesk, Inc. All rights reserved.
//
// Use of this software is subject to the terms of the Autodesk license
// agreement provided at the time of installation or download, or which
// otherwise accompanies this software in either electronic or hard copy form.
// -----------------------------------------------------------------------



using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Diagnostics;

namespace DrillingAutomation
{
    [RunInstaller(true)]
    public partial class DrillingAutomationInstaller : System.Configuration.Install.Installer
    {
        public DrillingAutomationInstaller()
        {
            InitializeComponent();
        }
        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);
            // Register the .NET assembly
            RegAsm("/codebase");

            // Register the plugin component category
            RegisterCOMCategory(true);
        }
        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);
            
            // Unregister the plugin component category
            RegisterCOMCategory(false);

            // Unregister the .NET assembly
            RegAsm("/u");
        }
        private static void RegAsm(string parameters)
        {
            // RuntimeEnvironment.GetRuntimeDirectory() returns something like
            // C:\Windows\Microsoft.NET\Framework64\v2.0.50727\
            // But we only want the "C:\Windows\Microsoft.NET" part
            string net_base = Path.GetFullPath(Path.Combine(
            RuntimeEnvironment.GetRuntimeDirectory(), @"..\.."));
            // Create paths to 32bit and 64bit versions of regasm.exe
            string[] to_exec = new[]
            {
            string.Concat(net_base, "\\Framework\\",
            RuntimeEnvironment.GetSystemVersion(), "\\regasm.exe"),
            string.Concat(net_base, "\\Framework64\\",
            RuntimeEnvironment.GetSystemVersion(), "\\regasm.exe")
            };
            // Get the path to the plugin's location
            var dll_path = Assembly.GetExecutingAssembly().Location;
            foreach (string path in to_exec)
            {
                // Skip the executables that do not exist; This most likely happens on
                // a 32bit machine when processing the path to 64bit regasm
                if (!File.Exists(path)) continue;
                // Build the argument string
                string args = string.Format("\"{0}\" {1}", dll_path, parameters);
                // Launch the regasm.exe process
                LaunchProcess(path, args);
            }
        }

        private static void LaunchProcess(string path, string arguments)
        {
            // Create a new process object, and setup its startup structure
            Process process = new Process
            {
                StartInfo =
                {
                CreateNoWindow = true,
                ErrorDialog = false,
                UseShellExecute = false,
                FileName = path,
                Arguments = arguments
                }
            };
            // Start the process, and wait for it to terminate
            using (process)
            {
                process.Start();
                process.WaitForExit();
            }
        }

        static string plugin_guid = "{7893D42C-0A4D-44E4-9B06-8C9E2E8F2BE2}";
        static string plugin_comp_category = "{311B0135-1826-4A8C-98DE-F313289F815E}";
        private static void RegisterCOMCategory(bool register)
        {
            // Build the registry key we wish to add
            string reg_key = "HKCR\\CLSID\\" + plugin_guid +
            "\\Implemented Categories\\" + plugin_comp_category;
            // Build the path
            string path = Environment.SystemDirectory + "\\reg.exe";
            // Loop over the platforms we wish to add the key to
            string[] platforms = new[] { "32", "64" };
            foreach (string platform in platforms)
            {
                // Build the arguments
                string args = (register) ? "ADD" : "DELETE";
                args += " \"" + reg_key + "\" /reg:" + platform + " /f";
                // Launch the reg.exe process
                LaunchProcess(path, args);
            }
        }
    }
}
