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
//===== PowerMILLVericutInterface application's Tap ToolClass ============================================================================================================================================================================
//=======================================================================================================================================================================================================================
//=======================================================================================================================================================================================================================


// ERROR: Not supported in C#: OptionDeclaration
namespace PowerMILLExporter.Tools
{

    public class TapTool : Tool
    {

        #region "Properties"
        /// <summary>
        /// Step
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double ToolStep
        {
            get
            {
                return dToolStep;
            }
        }

        private double dToolStep = 0;
        /// <summary>
        /// Chamfer Diameter
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double ChamferDiameter
        {
            get
            {
                return dChamferDiameter;
            }
        }

        private double dChamferDiameter = 0;
        /// <summary>
        /// No cut Diameter
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double NoCutDiameter
        {
            get
            {
                return dNoCutDiameter;
            }
        }

        private double dNoCutDiameter = 0;
        /// <summary>
        /// Length of the Chamfer
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double ChamferLength
        {
            get
            {
                return dChamferLength;
            }
        }
        #endregion
        private double dChamferLength = 0;

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
        /// <param name="ToolStep">Step</param>
        /// <param name="ChamferDiameter">Chamfer Diameter</param>
        /// <param name="NoCutDiameter">No cut Diameter</param>
        /// <param name="ChamferLength">Length of the Chamfer</param>
        /// <remarks></remarks>
        public TapTool(string Name, string ToolHolderPath, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length, double OverLength, double CuttingLength, string ToolShankPath, int TeethNumber, double ToolStep, double ChamferDiameter, double NoCutDiameter, double ChamferLength)
            : base("3", Name, ToolHolderPath, ToolNumber, OverHang, Gauge, Diameter, Length, OverLength, CuttingLength, ToolShankPath, TeethNumber)
        {

            dToolStep = ToolStep;
            dChamferDiameter = ChamferDiameter;
            dNoCutDiameter = NoCutDiameter;
            dChamferLength = ChamferLength;
        }
        public TapTool(string Name, string ToolHolderPath, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length, double OverLength, double CuttingLength, string ToolShankPath, int TeethNumber, double ToolStep, double ChamferDiameter, double NoCutDiameter, double ChamferLength, string Type)
            : base("3", Name, ToolHolderPath, ToolNumber, OverHang, Gauge, Diameter, Length, OverLength, CuttingLength, ToolShankPath, TeethNumber, Type)
        {

            dToolStep = ToolStep;
            dChamferDiameter = ChamferDiameter;
            dNoCutDiameter = NoCutDiameter;
            dChamferLength = ChamferLength;
        }
        #endregion
    }
}