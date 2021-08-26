// -----------------------------------------------------------------------
// Copyright 2021 Autodesk, Inc. All rights reserved.
//
// Use of this software is subject to the terms of the Autodesk license
// agreement provided at the time of installation or download, or which
// otherwise accompanies this software in either electronic or hard copy form.
// -----------------------------------------------------------------------

using System;
using System.Collections;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;

namespace VaultForPowerMillInstall
{
    [RunInstaller(true)]
    public partial class RegisterDLL : System.Configuration.Install.Installer
    {
        public RegisterDLL()
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
            // RuntimeEnvironment.GetRuntimeDirectory() returns something like C:\Windows\Microsoft.NET\Framework64\v2.0.50727\
            // But we only want the "C:\Windows\Microsoft.NET" part
            string netBase = Path.GetFullPath(Path.Combine(RuntimeEnvironment.GetRuntimeDirectory(), @"..\.."));

            // Create paths to 32bit and 64bit versions of regasm.exe
            string[] toExecute =
            {
                //Path.Combine(netBase, "Framework", RuntimeEnvironment.GetSystemVersion(), "regasm.exe"),
                Path.Combine(netBase, "Framework64", RuntimeEnvironment.GetSystemVersion(), "regasm.exe")
            };

            // Get the path to the plugin's location
            var dllPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "ExamplePowerMillPlugin.dll");
            foreach (string path in toExecute)
            {
                // Skip the executables that do not exist; This most likely happens on a 32bit machine when processing the path to 64bit regasm
                if (File.Exists(path))
                {
                    // Build the argument string
                    string args = "\"" + dllPath + "\" " + parameters;

                    // Launch the regasm.exe process
                    LaunchProcess(path, args);
                }
            }
        }

        private static void RegisterCOMCategory(bool register)
        {
            // Build the registry key we wish to add
//TODO: Change this Guid to match your plugin Guid
            string pluginGuid = "{0D229F75-EA25-4B51-8EA5-991232FFCCE6}";
            string pluginComponentCategory = "{311B0135-1826-4A8C-98DE-F313289F815E}";
            string regKey = "HKCR\\CLSID\\" + pluginGuid + "\\Implemented Categories\\" + pluginComponentCategory;

            // Build the path
            string path = Environment.SystemDirectory + "\\reg.exe";

            // Loop over the platforms we wish to add the key to
            string[] platforms = {"32", "64"};
            foreach (string platform in platforms)
            {
                // Build the arguments
                string args = register ? "ADD" : "DELETE";
                args += " \"" + regKey + "\" /reg:" + platform + " /f";

                // Launch the reg.exe process
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
    }
}