// -----------------------------------------------------------------------
// Copyright 2018 Autodesk, Inc. All rights reserved.
// 
// Use of this software is subject to the terms of the Autodesk license
// agreement provided at the time of installation or download, or which
// otherwise accompanies this software in either electronic or hard copy form.
// -----------------------------------------------------------------------

using System;
//=======================================================================================================================================================================================================================
//=======================================================================================================================================================================================================================
//===== PowerMILLVericutInterface application's TopAndSideMill ToolClass ============================================================================================================================================================================
//=======================================================================================================================================================================================================================
//=======================================================================================================================================================================================================================


// ERROR: Not supported in C#: OptionDeclaration
namespace PowerMILLExporter.Tools
{

    public class TopAndSideMill : Tool
    {

        /// <summary>
        /// Radius Inferior
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double RadiusInf
        {
            get
            {
                return dRadiusInf;
            }
        }

        private double dRadiusInf = 0;

        /// <summary>
        /// Radius Superior
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double RadiusSup
        {
            get
            {
                return dRadiusSup;
            }
        }

        private double dRadiusSup = 0;
        /// <summary>
        /// Side Angle
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double SideAngle
        {
            get
            {
                return dSideAngle;
            }
        }

        private double dSideAngle = 0;
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
        /// <param name="RadiusInf">Tool Radius Inferior</param>
        /// <param name="RadiusSup">Tool Radius Superior</param>
        /// <param name="SideAngle">Tool Side Angle</param>
        /// <remarks>Cutting length = Thickness</remarks>
        public TopAndSideMill(string Name, string ToolHolderPath, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length, double OverLength, double CuttingLength, string ToolShankPath, int TeethNumber, double RadiusInf, double RadiusSup, double SideAngle)
            : base("7", Name, ToolHolderPath, ToolNumber, OverHang, Gauge, Diameter, Length, OverLength, CuttingLength, ToolShankPath, TeethNumber)
        {

            dRadiusInf = RadiusInf;
            dRadiusSup = RadiusSup;
            dSideAngle = SideAngle;

        }
        public TopAndSideMill(string Name, string ToolHolderPath, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length, double OverLength, double CuttingLength, string ToolShankPath, int TeethNumber, double RadiusInf, double RadiusSup, double SideAngle, string Type)
            : base("7", Name, ToolHolderPath, ToolNumber, OverHang, Gauge, Diameter, Length, OverLength, CuttingLength, ToolShankPath, TeethNumber, Type)
        {

            dRadiusInf = RadiusInf;
            dRadiusSup = RadiusSup;
            dSideAngle = SideAngle;

        }
        #endregion
    }
}