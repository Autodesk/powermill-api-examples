// -----------------------------------------------------------------------
// Copyright 2021 Autodesk, Inc. All rights reserved.
//
// Use of this software is subject to the terms of the Autodesk license
// agreement provided at the time of installation or download, or which
// otherwise accompanies this software in either electronic or hard copy form.
// -----------------------------------------------------------------------

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Delcam.Plugins.Framework;

namespace ExamplePowerMillPlugin
{
//TODO: "Change this Guid to be something unique";
    [Guid("0D229F75-EA25-4B51-8EA5-991232FFCCE6")]
    [ClassInterface(ClassInterfaceType.None)]
    [ComVisible(true)]
    public class ExamplePowerMillPlugin : PluginFrameworkWithPanesAndTabs
    {
        public override string PluginAssemblyName => "ExamplePowerMillPlugin";
//TODO: "Change this Guid to match the one above that you uniquely generated";
        public override Guid PluginGuid => new Guid("0D229F75-EA25-4B51-8EA5-991232FFCCE6");
        public override string PluginName => "ExamplePowerMillPlugin";
        public override string PluginAuthor => "Your name here";
        public override string PluginDescription => "Example plugin";
        public override string PluginIconPath => null;
        public override Version PluginVersion => new Version(1, 0);
        public override Version PowerMILLVersion => new Version(2021, 0);
        public override bool PluginHasOptions => false;
        protected override void register_panes()
        {
            var pane = new VerticalPaneExample();
            register_pane(new PaneDefinition(pane, 900, 375, "Plugin for PowerMill", null));

        }

        protected override void register_tabs()
        {
            var tab = new HorizontalTabExample();
            register_tab(new TabDefinition(tab, 200, "Plugin for PowerMill", null));
        }
    }
}
