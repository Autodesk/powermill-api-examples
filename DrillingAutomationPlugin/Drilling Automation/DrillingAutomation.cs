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
using PowerMILL;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Windows.Interop;
using Delcam.Plugins.Framework;
using System.Text.RegularExpressions;
using Delcam.Plugins.Localisation;

namespace DrillingAutomation
{
    [Guid("7893D42C-0A4D-44E4-9B06-8C9E2E8F2BE2")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]

    public class DrillingAutomation_Plugin : PluginFrameworkWithPanes,
                                PowerMILL.IPowerMILLPluginPane,
                                Delcam.Plugins.Framework.IPluginCommunicationsInterface

    {
        public override string PluginName { get { return "Drilling Automation"; } }
        public override string PluginAuthor { get { return "Autodesk"; } }
        public override string PluginDescription { get { return "Automate drilling"; } }
        public override string PluginIconPath { get { return "Icons/Icon.ico"; } }
        public override Version PluginVersion { get { return new Version(1, 0, 14); } }
        public override Version PowerMILLVersion { get { return new Version(2019, 0); } }
        public override bool PluginHasOptions { get { return true; } }

        public override void DisplayOptionsForm()
        {
            base.DisplayOptionsForm();
            Options OptionForm = new Options();
            OptionForm.Show();
        }
        public override string PluginAssemblyName { get { return System.Reflection.Assembly.GetExecutingAssembly().GetName().Name; } }
        public override Guid PluginGuid { get { return new Guid("7893D42C-0A4D-44E4-9B06-8C9E2E8F2BE2"); } }

        private DrillingAutomation.DrillingAutomationPaneWPF oDrillingAutomationPaneWPF;
        public override void setup_framework(string token, PluginServices services, int parent_window_hwnd)
        {
            base.setup_framework(token, services, parent_window_hwnd);
        }

        protected override void register_panes()
        {
            // create the plugin pane
            oDrillingAutomationPaneWPF = new DrillingAutomation.DrillingAutomationPaneWPF(this);
            register_pane(new PaneDefinition(oDrillingAutomationPaneWPF, 540, 295, "Drilling Automation", "Icons/Icon.ico"));
        }

 
    }
}
