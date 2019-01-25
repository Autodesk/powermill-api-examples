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
//===== PowerMILLVericutInterface application's NCProgramTool ============================================================================================================================================================================
//=======================================================================================================================================================================================================================
//=======================================================================================================================================================================================================================



// ERROR: Not supported in C#: OptionDeclaration
namespace PowerMILLExporter.Tools
{

    /// <summary>
    /// This class is used to get NCprogram Tool number (they can be different from tool number)
    /// </summary>
    /// <remarks></remarks>
    public class NCProgramTool
    {
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
            set
            {
                sName = value;
            }
        }
        private string sName = null;

        /// <summary>
        /// Name of the toolpath
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public string Toolpath
        {
            get
            {
                return sToolpath;
            }
            set
            {
                sToolpath = value;
            }
        }
        private string sToolpath = null;

        /// <summary>
        /// Tool Number
        /// </summary>
        /// <value></value>
        /// <returns></returns>
        /// <remarks></remarks>
        public int Number
        {
            get
            {
                return iNumber;
            }
            set
            {
                iNumber = value;
            }
        }
        private int iNumber = 0;
    }
}