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

//=======================================================================================================================================================================================================================
//=======================================================================================================================================================================================================================
//===== PowerMILLVericutInterface application's TaperSpherical ToolClass ============================================================================================================================================================================
//=======================================================================================================================================================================================================================
//=======================================================================================================================================================================================================================

namespace PowerMILLExporter.Tools
{
    public class TaperedTipped : Tool
    {


        #region "Properties"
        /// <summary>
        /// Tool TaperAngle
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double TaperAngle
        {
            get
            {
                return dTaperAngle;
            }
        }
        private double dTaperAngle = 0;

        /// <summary>
        /// Tool TaperDiameter
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double TaperDiameter
        {
            get
            {
                return dTaperDia;
            }
        }
        private double dTaperDia = 0;

        /// <summary>
        /// Tool TipRadius
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double TipRadius
        {
            get
            {
                return dTipRad;
            }
        }
        private double dTipRad = 0;

        /// <summary>
        /// Tool TaperHeight
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double TaperHeight
        {
            get
            {
                return dTaperHeight;
            }
        }
        private double dTaperHeight = 0;

        #endregion
        #region "Methods"
        /// <summary>
        /// Sub New
        /// </summary>
        /// <param name="Name">Name of the tool</param>
        /// <param name="ToolHolderPath">ToolHolder Path</param>
        /// <param name="ToolNumber">ToolNumber</param>
        /// <param name="OverHang">Tool OverHang</param>
        /// <param name="Gauge">Tool Gauge</param>
        /// <param name="Diameter">Tool Diameter</param>
        /// <param name="Length">Tool Length</param>
        /// <param name="OverLength">Tool Total length</param>
        /// <param name="CuttingLength">Tool Cutting Length</param>
        /// <param name="ToolShankPath">Shank path</param>
        /// <param name="TeethNumber">Tool Teeth Number</param>
        /// <param name="CornerRadius">Tool CornerRadius</param>
        /// <param name="NoseAngle">Tool Nose Angle</param>
        /// <remarks></remarks>
        public TaperedTipped(string Name, string ToolHolderPath, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length, double OverLength, double CuttingLength, string ToolShankPath, int TeethNumber, double TaperAngle, double TaperDiameter, double TipRadius, double TaperHeight)
            : base("12", Name, ToolHolderPath, ToolNumber, OverHang, Gauge, Diameter, Length, OverLength, CuttingLength, ToolShankPath, TeethNumber)
        {

            dTaperAngle = TaperAngle;
            dTaperDia = TaperDiameter;
            dTipRad = TipRadius;
            dTaperHeight = TaperHeight;
        }
        public TaperedTipped(string Name, string ToolHolderPath, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length, double OverLength, double CuttingLength, string ToolShankPath, int TeethNumber, double TaperAngle, double TaperDiameter, double TipRadius, double TaperHeight, string Type)
            : base("12", Name, ToolHolderPath, ToolNumber, OverHang, Gauge, Diameter, Length, OverLength, CuttingLength, ToolShankPath, TeethNumber, Type)
        {

            dTaperAngle = TaperAngle;
            dTaperDia = TaperDiameter;
            dTipRad = TipRadius;
            dTaperHeight = TaperHeight;
        }
        #endregion
    }
}

