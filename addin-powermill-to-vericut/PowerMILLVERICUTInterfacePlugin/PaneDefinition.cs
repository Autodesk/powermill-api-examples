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
using System.Windows;
using System.Windows.Interop;

namespace PowerMILLNCSimulInterfacePlugin
{
    //=============================================================================
    /// <summary>
    /// A class to encapsulate all the required information about a PowerMILL plugin pane.
    /// </summary>
    public class PaneDefinition
    {
        // Private data
        private HwndSource m_hwnd_source;
        private FrameworkElement m_root_element;
        private int m_default_height;
        private int m_minimum_width;
        private string m_title;
        private string m_title_bitmap;

        //=============================================================================
        /// <summary>
        /// PaneDefinition c'tor.
        /// </summary>
        /// <param name="root_element">The root WPF element that defines the content of the pane.</param>
        /// <param name="height">The required pane height, which is fixed.</param>
        /// <param name="min_width">The minimum width for the pane, although a larger width may be used.</param>
        /// <param name="title">The title for the pane.</param>
        /// <param name="bitmap">The path of the pane icon within the resources.</param>
        public PaneDefinition(FrameworkElement root_element, int height, int min_width, string title, string bitmap)
        {
            m_root_element = root_element;
            m_default_height = height;
            m_minimum_width = min_width;
            m_title = title;
            m_title_bitmap = bitmap;
        }

        //=============================================================================
        /// <summary>
        /// Called when the pane is initialised.
        /// </summary>
        /// <param name="ParentWindow">The HWND of the parent window, that this pane will map on to.</param>
        /// <param name="Width">The starting width of the pane.</param>
        /// <param name="Height">The starting height of the pane.</param>
        public void initialise_pane(int ParentWindow, int Width, int Height)
        {
            // Build the parameters required for the HwndSource object
            HwndSourceParameters sourceParams = new HwndSourceParameters("PowerMILLPlugin");
            sourceParams.PositionX = 0;
            sourceParams.PositionY = 0;
            sourceParams.Height = Height;
            sourceParams.Width = Width;
            sourceParams.ParentWindow = new IntPtr(ParentWindow);
            sourceParams.WindowStyle = 0x10000000 | 0x40000000; // WS_VISIBLE | WS_CHILD;   

            // Create the HwndSource object
            m_hwnd_source = new HwndSource(sourceParams);

            // Set the root visual
            m_hwnd_source.RootVisual = m_root_element;

            // Set the size of the Hwnd
            m_hwnd_source.SizeToContent = SizeToContent.WidthAndHeight;

            // Set the size of the control
            m_root_element.Width = Width;
            m_root_element.Height = Height;

            // Add a message hook - this is a workaround so hear about WM_CHAR messages
            m_hwnd_source.AddHook(child_hwnd_source_hook);
        }

        //=============================================================================
        /// <summary>
        /// Called whenever the pane has been resized.
        /// </summary>
        /// <param name="Width">The new width of the pane.</param>
        /// <param name="Height">The new height of the pane.  This should always remain constant.</param>
        public void pane_resized(int Width, int Height)
        {
            // Resize the root element
            m_root_element.Width = Width;
            m_root_element.Height = Height;
        }

        //=============================================================================
        /// <summary>
        /// Returns the current height of the pane.
        /// </summary>
        public int PaneHeight
        {
            get
            {
                return m_default_height;
            }
        }

        //=============================================================================
        /// <summary>
        /// Returns the minimum width of the pane.
        /// </summary>
        public int MinimumPaneWidth
        {
            get
            {
                return m_minimum_width;
            }
        }

        //=============================================================================
        /// <summary>
        /// Returns the title of the pane.
        /// </summary>
        public string PaneTitle
        {
            get
            {
                return m_title;
            }
        }

        //=============================================================================
        /// <summary>
        /// Returns the path of the pane icon in the resources.
        /// </summary>
        public string PaneTitleBitmap
        {
            get
            {
                return m_title_bitmap;
            }
        }

        //=============================================================================
        /// <summary>
        /// Called when the pane is being destroyed.
        /// </summary>
        public void Uninitialise()
        {
            m_root_element.Visibility = Visibility.Hidden;
            m_root_element = null;
            m_hwnd_source = null;
        }

        //=============================================================================
        /// <summary>
        /// Returns the WPF root element of the pane definition.
        /// </summary>
        public FrameworkElement RootElement { get { return m_root_element; } }

        //=============================================================================
        /// <summary>
        /// 
        /// </summary>
        private IntPtr child_hwnd_source_hook(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            // The standard message loop for MFC dialogs includes a call to the IsDialogMessage() function. 
            // This function actually handles certain keys, and thus returns TRUE to PreTranslateMessage, meaning
            // that WM_CHAR messages are never dispatched, and we don't hear about them.  The upshot is that
            // Textboxes etc. in managed apps when working with native code and interop don't recieve characters.
            // The workaround is to listen out for WM_GETDLGCODE, which is posted by IsDialogMessage() - in response,
            // we can shout 'Hey, we want to hear about chars!', so WM_CHAR messages do still get dispatched....
            if (msg == 0x0087) // WM_GETDLGCODE
            {
                handled = true;
                return new IntPtr(0x0004); // DLGC_WANTALLKEYS          0x0080); // DLGC_WANTCHARS
            }

            return new IntPtr(0);
        }
    }
}
