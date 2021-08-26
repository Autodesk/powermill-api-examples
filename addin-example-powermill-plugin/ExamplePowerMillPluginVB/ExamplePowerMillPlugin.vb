' -----------------------------------------------------------------------
' Copyright 2021 Autodesk, Inc. All rights reserved.
'
' Use of this software is subject to the terms of the Autodesk license
' agreement provided at the time of installation or download, or which
' otherwise accompanies this software in either electronic or hard copy form.
' -----------------------------------------------------------------------

Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Threading.Tasks
Imports Delcam.Plugins.Framework

<Guid("0D229F75-EA25-4B51-8EA5-991232FFCCE6")>
<ClassInterface(ClassInterfaceType.None)>
<ComVisible(true)>
Public Class ExamplePowerMillPlugin
    Inherits PluginFrameworkWithPanesAndTabs

    Public Overrides ReadOnly Property PluginAssemblyName As String
        Get
            Return "ExamplePowerMillPlugin"
        End Get
    End Property

    Public Overrides ReadOnly Property PluginGuid As Guid
        Get
            Return New Guid("0D229F75-EA25-4B51-8EA5-991232FFCCE6")
        End Get
    End Property

    Public Overrides ReadOnly Property PluginName As String
        Get
            Return "ExamplePowerMillPlugin"
        End Get
    End Property

    Public Overrides ReadOnly Property PluginAuthor As String
        Get
            Return "Your name here"
        End Get
    End Property

    Public Overrides ReadOnly Property PluginDescription As String
        Get
            Return "Example plugin"
        End Get
    End Property

    Public Overrides ReadOnly Property PluginIconPath As String
        Get
            Return Nothing
        End Get
    End Property

    Public Overrides ReadOnly Property PluginVersion As Version
        Get
            Return New Version(1, 0)
        End Get
    End Property

    Public Overrides ReadOnly Property PowerMILLVersion As Version
        Get
            Return New Version(2021, 0)
        End Get
    End Property

    Public Overrides ReadOnly Property PluginHasOptions As Boolean
        Get
            Return False
        End Get
    End Property

    Protected Overrides Sub register_panes()
        Dim pane = New VerticalPaneExample()
        register_pane(New PaneDefinition(pane, 900, 375, "Plugin for PowerMill", Nothing))
    End Sub

    Protected Overrides Sub register_tabs()
        Dim tab = new HorizontalTabExample()
        register_tab(new TabDefinition(tab, 200, "Plugin for PowerMill", Nothing))
    End Sub

End Class
