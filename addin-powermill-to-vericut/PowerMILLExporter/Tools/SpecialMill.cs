// -----------------------------------------------------------------------
// Copyright 2018 Autodesk, Inc. All rights reserved.
// 
// Use of this software is subject to the terms of the Autodesk license
// agreement provided at the time of installation or download, or which
// otherwise accompanies this software in either electronic or hard copy form.
// -----------------------------------------------------------------------

using System;
using System.Collections;
using System.Collections.Generic;
//=======================================================================================================================================================================================================================
//=======================================================================================================================================================================================================================
//===== PowerMILLVericutInterface application's Milling ToolClass ============================================================================================================================================================================
//=======================================================================================================================================================================================================================
//=======================================================================================================================================================================================================================


// ERROR: Not supported in C#: OptionDeclaration
namespace PowerMILLExporter.Tools
{

    public class SpecialMill : Tool
    {

        #region "Properties"
        /// <summary>
        /// Tool Profile
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public List<string> Profile
        {
            get
            {
                return sProfile;
            }
        }
        private List<string> sProfile = new List<string>();

        private double dTipRad = 0;
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
        

        /// <summary>
        /// Tool Radius X Offset
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double RadXOffset
        {
            get
            {
                return dRadXOffset;
            }
        }
        private double dRadXOffset = 0;

        /// <summary>
        /// Tool Radius Y Offset
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double RadYOffset
        {
            get
            {
                return dRadYOffset;
            }
        }
        private double dRadYOffset = 0;
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
        /// <param name="RadXOffset">Tip radius X offset</param>
        /// <param name="RadYOffset">Tip radius Y offset</param>
        /// <param name="TipRadius">Tip radius</param>
        /// <param name="Profile">Tool Profile</param>
        /// <remarks></remarks>
        public SpecialMill(string Name, string ToolHolderPath, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length, double OverLength, double CuttingLength, double RadXOffset, double RadYOffset, double TipRadius, string ToolShankPath, int TeethNumber)
            : base("5", Name, ToolHolderPath, ToolNumber, OverHang, Gauge, Diameter, Length, OverLength, CuttingLength, ToolShankPath, TeethNumber)
        {
            dRadXOffset = RadXOffset;
            dRadYOffset = RadYOffset;
            dTipRad = TipRadius;
            //sProfile = Profile;
        }
        public SpecialMill(string Name, string ToolHolderPath, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length, double OverLength, double CuttingLength, double RadXOffset, double RadYOffset, double TipRadius, string ToolShankPath, int TeethNumber, string Type)
            : base("5", Name, ToolHolderPath, ToolNumber, OverHang, Gauge, Diameter, Length, OverLength, CuttingLength, ToolShankPath, TeethNumber, Type)
        {
            dRadXOffset = RadXOffset;
            dRadYOffset = RadYOffset;
            dTipRad = TipRadius;
            //sProfile = Profile;
		}
        public SpecialMill(string Name, string ToolHolderPath, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length, double OverLength, double CuttingLength, string ToolShankPath, int TeethNumber, List<string> Profile)
            : base("5", Name, ToolHolderPath, ToolNumber, OverHang, Gauge, Diameter, Length, OverLength, CuttingLength, ToolShankPath, TeethNumber)
        {

            sProfile = Profile;
        }
        public SpecialMill(string Name, string ToolHolderPath, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length, double OverLength, double CuttingLength, string ToolShankPath, int TeethNumber, List<string> Profile, string Type)
            : base("5", Name, ToolHolderPath, ToolNumber, OverHang, Gauge, Diameter, Length, OverLength, CuttingLength, ToolShankPath, TeethNumber, Type)
        {

            sProfile = Profile;
        }
        #endregion
    }
}