﻿'------------------------------------------------------------------------------
' <auto-generated>
'     Este código fue generado por una herramienta.
'     Versión de runtime:4.0.30319.42000
'
'     Los cambios en este archivo podrían causar un comportamiento incorrecto y se perderán si
'     se vuelve a generar el código.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict Off
Option Explicit On

Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.ReportSource
Imports CrystalDecisions.Shared
Imports System
Imports System.ComponentModel


Public Class SPYCS103x
    Inherits ReportClass
    
    Public Sub New()
        MyBase.New
    End Sub
    
    Public Overrides Property ResourceName() As String
        Get
            Return "SPYCS103x.rpt"
        End Get
        Set
            'Do nothing
        End Set
    End Property
    
    Public Overrides Property NewGenerator() As Boolean
        Get
            Return true
        End Get
        Set
            'Do nothing
        End Set
    End Property
    
    Public Overrides Property FullResourceName() As String
        Get
            Return "SPYC3.SPYCS103x.rpt"
        End Get
        Set
            'Do nothing
        End Set
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Section1() As CrystalDecisions.CrystalReports.Engine.Section
        Get
            Return Me.ReportDefinition.Sections(0)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Section2() As CrystalDecisions.CrystalReports.Engine.Section
        Get
            Return Me.ReportDefinition.Sections(1)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Section3() As CrystalDecisions.CrystalReports.Engine.Section
        Get
            Return Me.ReportDefinition.Sections(2)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Section4() As CrystalDecisions.CrystalReports.Engine.Section
        Get
            Return Me.ReportDefinition.Sections(3)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Section5() As CrystalDecisions.CrystalReports.Engine.Section
        Get
            Return Me.ReportDefinition.Sections(4)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P01() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(0)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P02() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(1)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P03() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(2)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P04() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(3)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P05() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(4)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P06() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(5)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P07() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(6)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P08() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(7)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P09() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(8)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P10() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(9)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P11() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(10)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P12() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(11)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P13() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(12)
        End Get
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public ReadOnly Property Parameter_P14() As CrystalDecisions.[Shared].IParameterField
        Get
            Return Me.DataDefinition.ParameterFields(13)
        End Get
    End Property
End Class

<System.Drawing.ToolboxBitmapAttribute(GetType(CrystalDecisions.[Shared].ExportOptions), "report.bmp")>  _
Public Class CachedSPYCS103x
    Inherits Component
    Implements ICachedReport
    
    Public Sub New()
        MyBase.New
    End Sub
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public Overridable Property IsCacheable() As Boolean Implements CrystalDecisions.ReportSource.ICachedReport.IsCacheable
        Get
            Return true
        End Get
        Set
            '
        End Set
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public Overridable Property ShareDBLogonInfo() As Boolean Implements CrystalDecisions.ReportSource.ICachedReport.ShareDBLogonInfo
        Get
            Return false
        End Get
        Set
            '
        End Set
    End Property
    
    <Browsable(false),  _
     DesignerSerializationVisibilityAttribute(System.ComponentModel.DesignerSerializationVisibility.Hidden)>  _
    Public Overridable Property CacheTimeOut() As System.TimeSpan Implements CrystalDecisions.ReportSource.ICachedReport.CacheTimeOut
        Get
            Return CachedReportConstants.DEFAULT_TIMEOUT
        End Get
        Set
            '
        End Set
    End Property
    
    Public Overridable Function CreateReport() As CrystalDecisions.CrystalReports.Engine.ReportDocument Implements CrystalDecisions.ReportSource.ICachedReport.CreateReport
        Dim rpt As SPYCS103x = New SPYCS103x()
        rpt.Site = Me.Site
        Return rpt
    End Function
    
    Public Overridable Function GetCustomizedCacheKey(ByVal request As RequestContext) As String Implements CrystalDecisions.ReportSource.ICachedReport.GetCustomizedCacheKey
        Dim key As [String] = Nothing
        '// The following is the code used to generate the default
        '// cache key for caching report jobs in the ASP.NET Cache.
        '// Feel free to modify this code to suit your needs.
        '// Returning key == null causes the default cache key to
        '// be generated.
        '
        'key = RequestContext.BuildCompleteCacheKey(
        '    request,
        '    null,       // sReportFilename
        '    this.GetType(),
        '    this.ShareDBLogonInfo );
        Return key
    End Function
End Class
