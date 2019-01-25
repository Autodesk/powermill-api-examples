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

namespace PowerMILLExporter.Tools
{
    public class Tools
    {
        public List<MillingTool> oMillingTools { get; set; }
        public List<TaperSpherical> oTaperSphericalTools { get; set; }
        public List<DrillingTool> oDrillingTools { get; set; }
        public List<TapTool> oTapTools { get; set; }
        public List<ATP5Mill> oATP5MillTools { get; set; }
        public List<TopAndSideMill> oTopAndSideMillTools { get; set; }
        public List<TaperedTipped> oTaperTippedTools { get; set; }
        public List<SpecialMill> oSpecialMILLTools { get; set; }
        public List<_2DTool> o2DTools { get; set; }

        public Tools() 
        {
            oMillingTools = new List<MillingTool>();
            oTaperSphericalTools = new List<TaperSpherical>();
            oDrillingTools = new List<DrillingTool>();
            oTapTools = new List<TapTool>();
            oATP5MillTools = new List<ATP5Mill>();
            oTopAndSideMillTools = new List<TopAndSideMill>();
            oTaperTippedTools = new List<TaperedTipped>();
            oSpecialMILLTools = new List<SpecialMill>();
            o2DTools = new List<_2DTool>();
        }

    }
}
