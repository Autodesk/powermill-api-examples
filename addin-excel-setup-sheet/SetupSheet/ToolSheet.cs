// -----------------------------------------------------------------------
// Copyright 2019 Autodesk, Inc. All rights reserved.
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
using PowerMILL;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Interop;
using Delcam.Plugins.Framework;
using System.Text.RegularExpressions;
using Delcam.Plugins.Localisation;

namespace SetupSheet
{
    [Guid("8C96851C-7A01-4389-8FBF-22C3DC7B09FD")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]

    public class SetupSheet_Plugin : PluginFrameworkWithPanes,
                                PowerMILL.IPowerMILLPluginPane,
                                Delcam.Plugins.Framework.IPluginCommunicationsInterface

    {
        public override string PluginName { get { return "Excel SetupSheet"; } }
        public override string PluginAuthor { get { return "Autodesk"; } }
        public override string PluginDescription { get { return "Export toolpath and tool list to Excel"; } }
        public override string PluginIconPath { get { return "Icons/Icon_24.ico"; } }
        public override Version PluginVersion { get { return new Version(1, 0, 5); } }
        public override Version PowerMILLVersion { get { return new Version(2018, 0); } }
        public override bool PluginHasOptions { get { return true; } }

        public override void DisplayOptionsForm()
        {
            base.DisplayOptionsForm();
            Options OptionForm = new Options();
            OptionForm.Show();
        }
        public override string PluginAssemblyName { get { return System.Reflection.Assembly.GetExecutingAssembly().GetName().Name; } }
        public override Guid PluginGuid { get { return new Guid("8C96851C-7A01-4389-8FBF-22C3DC7B09FD"); } }

        private SetupSheet.ToolSheetPaneWPF oToolSheetPaneWPF;
        public override void setup_framework(string token, PluginServices services, int parent_window_hwnd)
        {
            base.setup_framework(token, services, parent_window_hwnd);
        }

        protected override void register_panes()
        {
            // create the plugin pane
            oToolSheetPaneWPF = new SetupSheet.ToolSheetPaneWPF(this);
            register_pane(new PaneDefinition(oToolSheetPaneWPF, 440, 300, "Excel SetupSheet", "Icons/Icon_24.ico"));
        }

 
    }
}
