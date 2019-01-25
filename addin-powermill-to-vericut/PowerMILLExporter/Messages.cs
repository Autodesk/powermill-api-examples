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
using System.IO;
using System.Text.RegularExpressions;

namespace PowerMILLExporter
{
    public static class Messages
    {
        public static string PluginName
        {
            get { return plugin_name; }
            set { plugin_name = value; }
        }
        private static string plugin_name;

        public static MessageBoxResult ShowError(string msg)
        {
            return MessageBox.Show(msg, 
                plugin_name, MessageBoxButton.OK, MessageBoxImage.Error);
        }

        public static MessageBoxResult ShowError(string msg, object[] args)
        {
            return MessageBox.Show(String.Format(msg, args),
                plugin_name, MessageBoxButton.OK, MessageBoxImage.Error);
        }

        public static MessageBoxResult ShowError(string msg, string arg1)
        {
            string translation = msg;
            var regex = new Regex(Regex.Escape("%s"));
            var newText = regex.Replace(translation, arg1, 1);
            return MessageBox.Show(newText.Replace("\\n", Environment.NewLine),
                plugin_name, MessageBoxButton.OK, MessageBoxImage.Error);
        }

        public static MessageBoxResult ShowError(string msg, string arg1, string arg2)
        {
            string translation = msg;
            var regex = new Regex(Regex.Escape("%s"));
            var newText = regex.Replace(translation, arg1, 1);
            newText = regex.Replace(newText, arg2, 1);
            return MessageBox.Show(newText.Replace("\\n", Environment.NewLine),
                plugin_name, MessageBoxButton.OK, MessageBoxImage.Error);
        }

        public static MessageBoxResult ShowWarning(string msg)
        {
            return MessageBox.Show(msg,
                plugin_name, MessageBoxButton.YesNo, MessageBoxImage.Warning);
        }

        public static MessageBoxResult ShowWarning(string msg, object[] args)
        {
            return MessageBox.Show(String.Format(msg, args),
                plugin_name, MessageBoxButton.YesNo, MessageBoxImage.Warning);
        }

        public static MessageBoxResult ShowWarning(string msg, string arg1)
        {
            string translation = msg;
            var regex = new Regex(Regex.Escape("%s"));
            var newText = regex.Replace(translation, arg1, 1);
            return MessageBox.Show(newText.Replace("\\n", Environment.NewLine),
                plugin_name, MessageBoxButton.YesNo, MessageBoxImage.Warning);
        }

        public static MessageBoxResult ShowWarning(string msg, string arg1, string arg2)
        {
            string translation = msg;
            var regex = new Regex(Regex.Escape("%s"));
            var newText = regex.Replace(translation, arg1, 1);
            newText = regex.Replace(newText, arg2, 1);
            return MessageBox.Show(newText.Replace("\\n", Environment.NewLine),
                plugin_name, MessageBoxButton.YesNo, MessageBoxImage.Warning);
        }

        public static MessageBoxResult ShowQuestion(string msg)
        {
            return MessageBox.Show(msg,
                plugin_name, MessageBoxButton.YesNo, MessageBoxImage.Warning);
        }

        public static void ShowMessage(string msg)
        {
            MessageBox.Show(msg,
                plugin_name, MessageBoxButton.OK, MessageBoxImage.Information);
        }

        public static string Translate(string msg)
        {
            return msg;
        }
    }
}
