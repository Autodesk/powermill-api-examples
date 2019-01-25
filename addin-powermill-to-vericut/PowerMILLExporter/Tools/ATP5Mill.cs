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
//===== PowerMILLVericutInterface application's ATP5MILL ToolClass ============================================================================================================================================================================
//=======================================================================================================================================================================================================================
//=======================================================================================================================================================================================================================


// ERROR: Not supported in C#: OptionDeclaration
namespace PowerMILLExporter.Tools
{


    public class ATP5Mill : Tool
    {

        #region "Properties"
        /// <summary>
        /// Tool CornerRadius
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double CornerRadius
        {
            get
            {
                return dCornerRadius;
            }
        }

        private double dCornerRadius = 0;
        /// <summary>
        /// Tool Nose Angle
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double NoseAngle
        {
            get
            {
                return dNoseAngle;
            }
        }

        private double dNoseAngle = 0;
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
        #endregion
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
        /// <param name="CornerRadius">Tool CornerRadius</param>
        /// <param name="NoseAngle">Tool Nose Angle</param>
        /// <param name="SideAngle">Side Angle</param>
        /// <remarks></remarks>
        public ATP5Mill(string Name, string ToolHolderPath, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length, double OverLength, double CuttingLength, string ToolShankPath, int TeethNumber, double CornerRadius, double NoseAngle, double SideAngle)
            : base("4", Name, ToolHolderPath, ToolNumber, OverHang, Gauge, Diameter, Length, OverLength, CuttingLength, ToolShankPath, TeethNumber)
        {

            dCornerRadius = CornerRadius;
            dNoseAngle = NoseAngle;
            dSideAngle = SideAngle;
        }
        public ATP5Mill(string Name, string ToolHolderPath, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length, double OverLength, double CuttingLength, string ToolShankPath, int TeethNumber, double CornerRadius, double NoseAngle, double SideAngle, string Type)
            : base("4", Name, ToolHolderPath, ToolNumber, OverHang, Gauge, Diameter, Length, OverLength, CuttingLength, ToolShankPath, TeethNumber, Type)
        {

            dCornerRadius = CornerRadius;
            dNoseAngle = NoseAngle;
            dSideAngle = SideAngle;
        }

        #endregion
    }
}