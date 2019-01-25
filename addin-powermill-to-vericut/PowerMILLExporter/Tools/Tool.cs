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
//===== PowerMILLVericutInterface application's ToolClass ============================================================================================================================================================================
//=======================================================================================================================================================================================================================
//=======================================================================================================================================================================================================================


namespace PowerMILLExporter.Tools
{

    public abstract class Tool
    {

        #region "Properties"

        /// <summary>
        /// Tool Type
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public string Type
        {
            get
            {
                return sType;
            }
        }

        private string sType = string.Empty;
        /// <summary>
        /// Name of the tool
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public string Name
        {
            get
            {
                return sName;
            }
        }

        private string sName = null;
        // ''' <summary>
        // ''' ToolHolder Profile Point Coordinate (X Y)
        // ''' </summary>
        // ''' <value></value>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        //Friend ReadOnly Property ToolHolder As List(Of Double)
        //    Get
        //        Return dToolHolder
        //    End Get
        //End Property
        //Private dToolHolder As New List(Of Double)

        /// <summary>
        /// ToolHolder file path (iges)
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public string ToolHolderPath
        {
            get
            {
                return sToolHolderPath;
            }
        }

        private string sToolHolderPath = null;

        /// <summary>
        /// ToolHolder profile 
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public string ToolHolderProfile
        {
            get
            {
                return sToolHolderProfile;
            }
        }

        private string sToolHolderProfile = null;

        /// <summary>
        /// ToolShank profile 
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public string ToolShankProfile
        {
            get
            {
                return sToolShankProfile;
            }
        }

        private string sToolShankProfile = null;

        /// <summary>
        /// ToolNumber
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int ToolNumber
        {
            get
            {
                return iToolNumber;
            }
        }

        private int iToolNumber = 0;
        /// <summary>
        /// Tool OverHang
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double OverHang
        {
            get
            {
                return dOverHang;
            }
        }

        private double dOverHang = 0;
        /// <summary>
        /// Tool Gauge
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double Gauge
        {
            get
            {
                return dGauge;
            }
        }

        private double dGauge = 0;
        /// <summary>
        /// Tool Diameter
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double Diameter
        {
            get
            {
                return dDiameter;
            }
        }

        private double dDiameter = 0;
        /// <summary>
        /// Tool Length
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double Length
        {
            get
            {
                return dLength;
            }
        }

        private double dLength = 0;
        /// <summary>
        /// Tool Total length
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double OverLength
        {
            get
            {
                return dOverLength;
            }
        }

        private double dOverLength = 0;
        /// <summary>
        /// Tool Cutting Length
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public double CuttingLength
        {
            get
            {
                return dCuttingLength;
            }
        }

        private double dCuttingLength = 0;
        // ''' <summary>
        // ''' ShankDiameter
        // ''' </summary>
        // ''' <value></value>
        // ''' <returns></returns>
        // ''' <remarks></remarks>
        //Friend ReadOnly Property ShankDiameter As Double
        //    Get
        //        Return dShankDiameter
        //    End Get
        //End Property
        //Private dShankDiameter As Double = Nothing

        /// <summary>
        /// ToolShank file path (iges)
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public string ToolShankPath
        {
            get
            {
                return sToolShankPath;
            }
        }

        private string sToolShankPath = null;
        /// <summary>
        /// Tool Teeth Number
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int TeethNumber
        {
            get
            {
                return iTeethNumber;
            }
        }

        private int iTeethNumber = 0;

        /// <summary>
        /// Fidia tool type
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public string ToolType
        {
            get
            {
                return sToolType;
            }
        }

        private string sToolType;
        
        #endregion

        /// <summary>
        /// Sub New
        /// </summary>
        /// <param name="Type">Tool Type</param>
        /// <param name="Name">Name of the tool</param>
        /// <param name="ToolHolderPath">ToolHolder file path (iges)</param>
        /// <param name="ToolNumber">ToolNumber</param>
        /// <param name="OverHang">Tool OverHang</param>
        /// <param name="Gauge">Tool Gauge</param>
        /// <param name="Diameter">Tool Diameter</param>
        /// <param name="Length">Tool Length</param>
        /// <param name="OverLength">Tool Total length</param>
        /// <param name="CuttingLength">Tool Cutting Length</param>
        /// <param name="ToolShankPath">ToolShank file path (iges)</param>
        /// <param name="TeethNumber">Tool Teeth Number</param>
        /// <remarks></remarks>
        /*internal Tool(string Type, string Name, string ToolHolderPath, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length, double OverLength, double CuttingLength,
        string ToolShankPath, int TeethNumber)
        {
            sType = Type;
            sName = Name;
            sToolHolderPath = ToolHolderPath;
            iToolNumber = ToolNumber;
            dOverHang = OverHang;
            dGauge = Gauge;
            dDiameter = Diameter;
            dLength = Length;
            dOverLength = OverLength;
            dCuttingLength = CuttingLength;
            sToolShankPath = ToolShankPath;
            iTeethNumber = TeethNumber;
        }*/

        internal Tool(string Type, string Name, string ToolHolderProfile, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length,
                      double OverLength, double CuttingLength, string ToolShankProfile, int TeethNumber)
        {
            sType = Type;
            sName = Name;
            sToolHolderProfile = ToolHolderProfile;
            iToolNumber = ToolNumber;
            dOverHang = OverHang;
            dGauge = Gauge;
            dDiameter = Diameter;
            dLength = Length;
            dOverLength = OverLength;
            dCuttingLength = CuttingLength;
            sToolShankProfile = ToolShankProfile;
            iTeethNumber = TeethNumber;
        }

        internal Tool(string Type, string Name, string ToolHolderProfile, int ToolNumber, double OverHang, double Gauge, double Diameter, double Length,
                      double OverLength, double CuttingLength, string ToolShankProfile, int TeethNumber, string ToolType)
        {
            sType = Type;
            sName = Name;
            sToolHolderProfile = ToolHolderProfile;
            iToolNumber = ToolNumber;
            dOverHang = OverHang;
            dGauge = Gauge;
            dDiameter = Diameter;
            dLength = Length;
            dOverLength = OverLength;
            dCuttingLength = CuttingLength;
            sToolShankProfile = ToolShankProfile;
            iTeethNumber = TeethNumber;
            sToolType = ToolType;
        }
        
        #region "Methods"
        // ''' <summary>
        // ''' Sub New
        // ''' </summary>
        // ''' <param name="Type">Tool Type</param>
        // ''' <param name="Name">Name of the tool</param>
        // ''' <param name="ToolHolder">ToolHolder Profile Point Coordinate (X Y)</param>
        // ''' <param name="ToolNumber">ToolNumber</param>
        // ''' <param name="OverHang">Tool OverHang</param>
        // ''' <param name="Gauge">Tool Gauge</param>
        // ''' <param name="Diameter">Tool Diameter</param>
        // ''' <param name="Length">Tool Length</param>
        // ''' <param name="OverLength">Tool Total length</param>
        // ''' <param name="CuttingLength">Tool Cutting Length</param>
        // ''' <param name="ShankDiameter">Shank Diameter</param>
        // ''' <param name="TeethNumber">Tool Teeth Number</param>
        // ''' <remarks></remarks>
        //Friend Sub New(Type As String, Name As String, ToolHolder As List(Of Double), ToolNumber As Integer, OverHang As Double, Gauge As Double, Diameter As Double, Length As Double, OverLength As Double, _
        //                CuttingLength As Double, ShankDiameter As Double, TeethNumber As Integer)
        //    sType = Type
        //    sName = Name
        //    dToolHolder = ToolHolder
        //    iToolNumber = ToolNumber
        //    dOverHang = OverHang
        //    dGauge = Gauge
        //    dDiameter = Diameter
        //    dLength = Length
        //    dOverLength = OverLength
        //    dCuttingLength = CuttingLength
        //    dShankDiameter = ShankDiameter
        //    iTeethNumber = TeethNumber
        //End Sub
        #endregion
    }
}