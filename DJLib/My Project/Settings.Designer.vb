﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Namespace My
    
    <Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "10.0.0.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Partial Friend NotInheritable Class MySettings
        Inherits Global.System.Configuration.ApplicationSettingsBase
        
        Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings()),MySettings)
        
#Region "My.Settings Auto-Save Functionality"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(ByVal sender As Global.System.Object, ByVal e As Global.System.EventArgs)
        If My.Application.SaveMySettingsOnExit Then
            My.Settings.Save()
        End If
    End Sub
#End If
#End Region
        
        Public Shared ReadOnly Property [Default]() As MySettings
            Get
                
#If _MyType = "WindowsForms" Then
               If Not addedHandler Then
                    SyncLock addedHandlerLockObject
                        If Not addedHandler Then
                            AddHandler My.Application.Shutdown, AddressOf AutoSaveSettings
                            addedHandler = True
                        End If
                    End SyncLock
                End If
#End If
                Return defaultInstance
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Host=hon03nt;port=5432;database=LogisticDb")>  _
        Public Property con1() As String
            Get
                Return CType(Me("con1"),String)
            End Get
            Set
                Me("con1") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Host=hon13-0002L;port=5432;database=hr")>  _
        Public Property con2() As String
            Get
                Return CType(Me("con2"),String)
            End Get
            Set
                Me("con2") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("C50B3C89CB21F4F1422FF158A5B42D0E8DB8CB5CDA1742572A487D9401E3400267682B202B7465118"& _ 
            "91C1BAF47F8D25C07F6C39A104696DB51F17C529AD3CABE")>  _
        Public Property validationkey() As String
            Get
                Return CType(Me("validationkey"),String)
            End Get
            Set
                Me("validationkey") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("8A9BE8FD67AF6979E7D20198CFEA50DD3D3799C77AF2B72F")>  _
        Public Property decriptionkey() As String
            Get
                Return CType(Me("decriptionkey"),String)
            End Get
            Set
                Me("decriptionkey") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("SHA1")>  _
        Public Property validation() As String
            Get
                Return CType(Me("validation"),String)
            End Get
            Set
                Me("validation") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Host=hon10-0046D;port=5432;database=hr;CommandTimeOut=1000;TimeOut=1000;")>  _
        Public Property connectionstring() As String
            Get
                Return CType(Me("connectionstring"),String)
            End Get
            Set
                Me("connectionstring") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("ODBC;DSN=PostgreSQLHR;")>  _
        Public Property oExCon() As String
            Get
                Return CType(Me("oExCon"),String)
            End Get
            Set
                Me("oExCon") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Black")>  _
        Public Property MenuStripGradientBegin() As Global.System.Drawing.Color
            Get
                Return CType(Me("MenuStripGradientBegin"),Global.System.Drawing.Color)
            End Get
            Set
                Me("MenuStripGradientBegin") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("DimGray")>  _
        Public Property MenuStripGradientEnd() As Global.System.Drawing.Color
            Get
                Return CType(Me("MenuStripGradientEnd"),Global.System.Drawing.Color)
            End Get
            Set
                Me("MenuStripGradientEnd") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("WhiteSmoke")>  _
        Public Property MenuStripForeColor() As Global.System.Drawing.Color
            Get
                Return CType(Me("MenuStripForeColor"),Global.System.Drawing.Color)
            End Get
            Set
                Me("MenuStripForeColor") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Maroon")>  _
        Public Property MenuItemPressedGradientBegin() As Global.System.Drawing.Color
            Get
                Return CType(Me("MenuItemPressedGradientBegin"),Global.System.Drawing.Color)
            End Get
            Set
                Me("MenuItemPressedGradientBegin") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Maroon")>  _
        Public Property MenuItemPressedGradientMiddle() As Global.System.Drawing.Color
            Get
                Return CType(Me("MenuItemPressedGradientMiddle"),Global.System.Drawing.Color)
            End Get
            Set
                Me("MenuItemPressedGradientMiddle") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Silver")>  _
        Public Property MenuItemPressedGradientEnd() As Global.System.Drawing.Color
            Get
                Return CType(Me("MenuItemPressedGradientEnd"),Global.System.Drawing.Color)
            End Get
            Set
                Me("MenuItemPressedGradientEnd") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Black")>  _
        Public Property MenuItemSelectedGradientBegin() As Global.System.Drawing.Color
            Get
                Return CType(Me("MenuItemSelectedGradientBegin"),Global.System.Drawing.Color)
            End Get
            Set
                Me("MenuItemSelectedGradientBegin") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Black")>  _
        Public Property MenuItemSelectedGradientEnd() As Global.System.Drawing.Color
            Get
                Return CType(Me("MenuItemSelectedGradientEnd"),Global.System.Drawing.Color)
            End Get
            Set
                Me("MenuItemSelectedGradientEnd") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Black")>  _
        Public Property MenuItemSelected() As Global.System.Drawing.Color
            Get
                Return CType(Me("MenuItemSelected"),Global.System.Drawing.Color)
            End Get
            Set
                Me("MenuItemSelected") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Maroon")>  _
        Public Property ImageMarginGradientBegin() As Global.System.Drawing.Color
            Get
                Return CType(Me("ImageMarginGradientBegin"),Global.System.Drawing.Color)
            End Get
            Set
                Me("ImageMarginGradientBegin") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("WhiteSmoke")>  _
        Public Property ImageMarginGradientMiddle() As Global.System.Drawing.Color
            Get
                Return CType(Me("ImageMarginGradientMiddle"),Global.System.Drawing.Color)
            End Get
            Set
                Me("ImageMarginGradientMiddle") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("WhiteSmoke")>  _
        Public Property ImageMarginGradientEnd() As Global.System.Drawing.Color
            Get
                Return CType(Me("ImageMarginGradientEnd"),Global.System.Drawing.Color)
            End Get
            Set
                Me("ImageMarginGradientEnd") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Black")>  _
        Public Property MenuBorder() As Global.System.Drawing.Color
            Get
                Return CType(Me("MenuBorder"),Global.System.Drawing.Color)
            End Get
            Set
                Me("MenuBorder") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Black")>  _
        Public Property MenuItemBorder() As Global.System.Drawing.Color
            Get
                Return CType(Me("MenuItemBorder"),Global.System.Drawing.Color)
            End Get
            Set
                Me("MenuItemBorder") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Black")>  _
        Public Property ToolStripBorder() As Global.System.Drawing.Color
            Get
                Return CType(Me("ToolStripBorder"),Global.System.Drawing.Color)
            End Get
            Set
                Me("ToolStripBorder") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Maroon")>  _
        Public Property ToolStripContentPanelGradientBegin() As Global.System.Drawing.Color
            Get
                Return CType(Me("ToolStripContentPanelGradientBegin"),Global.System.Drawing.Color)
            End Get
            Set
                Me("ToolStripContentPanelGradientBegin") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("HighlightText")>  _
        Public Property ToolStripContentPanelGradientEnd() As Global.System.Drawing.Color
            Get
                Return CType(Me("ToolStripContentPanelGradientEnd"),Global.System.Drawing.Color)
            End Get
            Set
                Me("ToolStripContentPanelGradientEnd") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("LightGray")>  _
        Public Property ToolStripDropDownBackground() As Global.System.Drawing.Color
            Get
                Return CType(Me("ToolStripDropDownBackground"),Global.System.Drawing.Color)
            End Get
            Set
                Me("ToolStripDropDownBackground") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("HighlightText")>  _
        Public Property ToolStripGradientBegin() As Global.System.Drawing.Color
            Get
                Return CType(Me("ToolStripGradientBegin"),Global.System.Drawing.Color)
            End Get
            Set
                Me("ToolStripGradientBegin") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("HighlightText")>  _
        Public Property ToolStripGradientMiddle() As Global.System.Drawing.Color
            Get
                Return CType(Me("ToolStripGradientMiddle"),Global.System.Drawing.Color)
            End Get
            Set
                Me("ToolStripGradientMiddle") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("HighlightText")>  _
        Public Property ToolStripGradientEnd() As Global.System.Drawing.Color
            Get
                Return CType(Me("ToolStripGradientEnd"),Global.System.Drawing.Color)
            End Get
            Set
                Me("ToolStripGradientEnd") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Black")>  _
        Public Property StatusStripGradientBegin() As Global.System.Drawing.Color
            Get
                Return CType(Me("StatusStripGradientBegin"),Global.System.Drawing.Color)
            End Get
            Set
                Me("StatusStripGradientBegin") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("DimGray")>  _
        Public Property StatusStripGradientEnd() As Global.System.Drawing.Color
            Get
                Return CType(Me("StatusStripGradientEnd"),Global.System.Drawing.Color)
            End Get
            Set
                Me("StatusStripGradientEnd") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Black")>  _
        Public Property ToolStripForeColor() As Global.System.Drawing.Color
            Get
                Return CType(Me("ToolStripForeColor"),Global.System.Drawing.Color)
            End Get
            Set
                Me("ToolStripForeColor") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("WhiteSmoke")>  _
        Public Property StatusStripForeColor() As Global.System.Drawing.Color
            Get
                Return CType(Me("StatusStripForeColor"),Global.System.Drawing.Color)
            End Get
            Set
                Me("StatusStripForeColor") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Gold")>  _
        Public Property ToolStripPanelGradientBegin() As Global.System.Drawing.Color
            Get
                Return CType(Me("ToolStripPanelGradientBegin"),Global.System.Drawing.Color)
            End Get
            Set
                Me("ToolStripPanelGradientBegin") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("InactiveCaption")>  _
        Public Property ToolStripPanelGradientEnd() As Global.System.Drawing.Color
            Get
                Return CType(Me("ToolStripPanelGradientEnd"),Global.System.Drawing.Color)
            End Get
            Set
                Me("ToolStripPanelGradientEnd") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Host=hon10-0046D;port=5432;database=hr")>  _
        Public Property con3() As String
            Get
                Return CType(Me("con3"),String)
            End Get
            Set
                Me("con3") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Host=localhost;port=5432;database=hr")>  _
        Public Property con4() As String
            Get
                Return CType(Me("con4"),String)
            End Get
            Set
                Me("con4") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("ODBC;DSN=PostgreSQLHR;")>  _
        Public Property oExCon1() As String
            Get
                Return CType(Me("oExCon1"),String)
            End Get
            Set
                Me("oExCon1") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("ODBC;DSN=PostgreSQLHRTest;")>  _
        Public Property oExCon2() As String
            Get
                Return CType(Me("oExCon2"),String)
            End Get
            Set
                Me("oExCon2") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Host=hon07nt;port=5432;database=hr;CommandTimeOut=1000;TimeOut=1000;")>  _
        Public Property con5() As String
            Get
                Return CType(Me("con5"),String)
            End Get
            Set
                Me("con5") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("ODBC;DSN=PostgreSQLHRLive;")>  _
        Public Property oExCon3() As String
            Get
                Return CType(Me("oExCon3"),String)
            End Get
            Set
                Me("oExCon3") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Host=hon10-0046D;port=5432;database=hr;CommandTimeOut=1000;TimeOut=1000;")>  _
        Public Property con6() As String
            Get
                Return CType(Me("con6"),String)
            End Get
            Set
                Me("con6") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Host=172.22.13.139;port=5432;database=hr")>  _
        Public Property con7() As String
            Get
                Return CType(Me("con7"),String)
            End Get
            Set
                Me("con7") = value
            End Set
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.SpecialSettingAttribute(Global.System.Configuration.SpecialSetting.ConnectionString),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Host=hon25nt;port=5432;database=hr;CommandTimeOut=1000;TimeOut=1000;")>  _
        Public ReadOnly Property connectionstring1() As String
            Get
                Return CType(Me("connectionstring1"),String)
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Host=hon10-0046D;port=5432;database=hr;CommandTimeOut=1000;TimeOut=1000;")>  _
        Public Property con8() As String
            Get
                Return CType(Me("con8"),String)
            End Get
            Set
                Me("con8") = value
            End Set
        End Property
    End Class
End Namespace

Namespace My
    
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Module MySettingsProperty
        
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>  _
        Friend ReadOnly Property Settings() As Global.DJLib.My.MySettings
            Get
                Return Global.DJLib.My.MySettings.Default
            End Get
        End Property
    End Module
End Namespace
