// -----------------------------------------------------------------------
// Copyright 2018 Autodesk, Inc. All rights reserved.
// 
// Use of this software is subject to the terms of the Autodesk license
// agreement provided at the time of installation or download, or which
// otherwise accompanies this software in either electronic or hard copy form.
// -----------------------------------------------------------------------

//=============================================================================
//
// PowerMILL must not be running when the plugin is rebuilt, as it the DLL will be locked.
//
// To register:
// --------------
//
//  x64:
//  ----
//
//    .NET 3.5 =>    cd C:\WINDOWS\Microsoft.NET\Framework64\v2.0.50727\
//    .NET 4   =>    cd C:\WINDOWS\Microsoft.NET\Framework64\v4.0.30319\
//
//    regasm.exe "C:\SVN\PowerMILL_Addins\PowerMILLVERICUTInterfacePlugin\PowerMILLVERICUTInterfacePlugin\bin\Debug\PowerMILLVERICUTInterfacePlugin.dll" /register /codebase
//
//  x86:
//  ----
//
//    .NET 3.5 =>    cd C:\WINDOWS\Microsoft.NET\Framework\v2.0.50727\
//    .NET 4   =>    cd C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\
//
//    regasm.exe "C:\SVN\PowerMILL_Addins\PowerMILLVERICUTInterfacePlugin\PowerMILLVERICUTInterfacePlugin\bin\Debug\PowerMILLVERICUTInterfacePlugin.dll" /register /codebase
//
// To un-register:
// ---------------
//    regasm.exe "C:\SVN\PowerMILL_Addins\PowerMILLVERICUTInterfacePlugin\PowerMILLVERICUTInterfacePlugin\bin\Debug\PowerMILLVERICUTInterfacePlugin.dll" /u
//
//
// Registry Link to PowerMILL (Do not change "311b0135-1826-4a8c-98de-f313289f815e" as this is PowerMILL Guid):
// ------------------------------------------------------------------------------------------------------------
//
//    REG ADD "HKCR\CLSID\{AFBB4C14-BD94-4692-98DE-645609C1B033}\Implemented Categories\{311b0135-1826-4a8c-98de-f313289f815e}" /reg:32 /f
//    REG ADD "HKCR\CLSID\{AFBB4C14-BD94-4692-98DE-645609C1B033}\Implemented Categories\{311b0135-1826-4a8c-98de-f313289f815e}" /reg:64 /f
//
// to delete: REG DELETE ...
//
//
//    REG DELETE "HKCR\CLSID\{AFBB4C14-BD94-4692-98DE-645609C1B033}\Implemented Categories\{311b0135-1826-4a8c-98de-f313289f815e}" /reg:32 /f
//    REG DELETE "HKCR\CLSID\{AFBB4C14-BD94-4692-98DE-645609C1B033}\Implemented Categories\{311b0135-1826-4a8c-98de-f313289f815e}" /reg:64 /f

using System;
using System.Runtime.InteropServices; // For Guid and ComVisible attributes
using System.Windows.Forms;           // For Message boxes and WinForms
using System.Windows.Interop;         // For WindowInteropHelper
using Delcam.Plugins.Framework;
using Delcam.Plugins.Localisation;

namespace PowerMILLVERICUTInterfacePlugin
{
// Attributes that allow the SamplePlugin class to be created as a COM object
// Note: regenerate this guid for each new plugin you create - it must be unique!
//       This plugin GUID is declared is many placed of the project so look for all of them!
[Guid("AFBB4C14-BD94-4692-98DE-645609C1B033")]
[ClassInterface(ClassInterfaceType.None)]
[ComVisible(true)]
public class PowerMILLPlugin :  PluginFrameworkWithPanes, 
                                PowerMILL.IPowerMILLPluginLicensable,
                                Delcam.Plugins.Framework.IPluginCommunicationsInterface
{
    // Licensing
    string PowerMILL.IPowerMILLPluginLicensable.RequiredLicenses()
    {
        return "";
    }

    // plugin version
    public static int VersionMajor { get { return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Major; } }
    public static int VersionMinor { get { return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Minor; } }
    public static int VersionBuild { get { return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version.Build; } }

    // The main pane
    private PowerMILLVERICUTInterfacePlugin.Forms.VericutPaneWPF oVERICUTPaneWPF;

#region PluginFrameworkWithPanes


protected override void register_panes()
{
    // create the plugin pane
    oVERICUTPaneWPF = new PowerMILLVERICUTInterfacePlugin.Forms.VericutPaneWPF(this);
    register_pane(new PaneDefinition(oVERICUTPaneWPF, 780, 300, "VERICUT Interface", "Images/Logo24_1.png"));
}

public override string PluginName
{
    get
    {
        return "VERICUT Interface";
    }
}
public override string PluginAuthor
{
    get
    {
        return "Autodesk";
    }
}

public override string PluginDescription
{
    get
    {
        return "Interface for VERICUT";
    }
}
public override string PluginIconPath
{
    get
    {
        return "Images/Logo32_1.png";
    }
}

public override Version PluginVersion
{
    get
    {
        return new Version(VersionMajor, VersionMinor, VersionBuild);
    }
}

public override Version PowerMILLVersion
{
    get
    {
        return new Version(15, 0, 14);
    }
}

public override bool PluginHasOptions
{
    get
    {
        return false;
    }
}

public override string PluginAssemblyName
{
    get
    {
        return System.Reflection.Assembly.GetExecutingAssembly().GetName().Name;
    }
}

public override Guid PluginGuid
{
    get
    {
        return new Guid("AFBB4C14-BD94-4692-98DE-645609C1B033");
    }
}

#endregion

}
}
