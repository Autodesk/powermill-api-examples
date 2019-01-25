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
using System.Runtime.InteropServices;

namespace PowerMILLVERICUTInterfacePlugin
{
    public class VericutMachineReader
    {
        [DllImport("MR_wrapper.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern string MR_extract_machine_file_name_CPP(string projectFileName);
        [DllImport("MR_wrapper.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int MR_extract_components_CPP(int option, string machineFileName, out IntPtr componentList);
        [DllImport("MR_wrapper.dll", CallingConvention = CallingConvention.Cdecl)]
        private static extern int MR_extract_subsystems_CPP(string machineFile, out IntPtr subsystemsList);

        public static List<string> GetMachineSubsystems(string template_proj_fpath)
        {
            List<string> subsystems = new List<string>();
            String machineFile = MR_extract_machine_file_name_CPP(template_proj_fpath);
            if (String.IsNullOrEmpty(machineFile))
            {
                PowerMILLExporter.Messages.ShowError(Properties.Resources.IDFS_FailMachineFileRead);
                return subsystems;
            }

            IntPtr unManagedSlist = IntPtr.Zero;
            int scount = MR_extract_subsystems_CPP(machineFile, out unManagedSlist);
            if (scount > 0)
            {
                // Convert unmanaged strings received from C++ side to C# format and release memory
                // allocated on C++ side.
                UnmananagedStrings2ManagedStringList(unManagedSlist, scount, out subsystems);
            }
            return subsystems;
        }


        /**********************************************************************/
        /**
         * Converts unmanaged C-style string into C# format.
         */
        static void UnmananagedStrings2ManagedStringList(IntPtr unmanagedStrings,
            int count, out List<string> managedStrings)
        {
            IntPtr[] intPtrArray = new IntPtr[count];
            managedStrings = new List<string>();

            Marshal.Copy(unmanagedStrings, intPtrArray, 0, count);

            for (int i = 0; i < count; i++)
            {
                managedStrings.Add(Marshal.PtrToStringAnsi(intPtrArray[i]));
                // Free memory allocated to unmanaged strings.
                Marshal.FreeCoTaskMem(intPtrArray[i]);
            }
            // Free memory allocated to  array of strings.
            Marshal.FreeCoTaskMem(unmanagedStrings);
        }
    }
}
