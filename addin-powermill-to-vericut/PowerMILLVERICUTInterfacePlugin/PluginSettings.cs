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
using System.Xml.Serialization;
using PowerMILLExporter;

namespace PowerMILLVERICUTInterfacePlugin
{
    [Serializable]
    public class PluginSettings
    {
        public ExportOptions export_options;
        public string vericut_fpath;
        public bool start_vericut;
        public bool tool_id_use_name;
        public bool tool_id_use_num;
    }
}
