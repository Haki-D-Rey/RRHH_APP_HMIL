Imports CrystalDecisions.CrystalReports.Engine
Public Class Frm_Visor_Reportes
    Private Property rptlocation As String
    Dim CR As ReportDocument
    Private Sub Frm_Visor_Reportes_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Call MOSTRAR_REPORTE_CRYSTAL()
    End Sub
    Private Sub Frm_Visor_Reportes_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.Escape Then
            Me.Close()
        End If
    End Sub
    Dim PathRPT As String
    'ULTIMO 155
    Private Sub MOSTRAR_REPORTE_CRYSTAL()
        PathRPT = "\\Hmil-spc-apprd\e$\v3 SPYC\Reportes\"
        Dim CR As New ReportDocument
        '------------------------------------------------------------------------------
        'SELECCION_PARAMETRO = FILTRO_SELECCION_EN_REPORTES
        If SELECCION_PARAMETRO = "" Then
            SELECCION_PARAMETRO = "*"
        End If
        If PARAMETRO = 1 Then
            rptlocation = PathRPT & "SPYC001.rpt"
            CR.Load(rptlocation)
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 2 Then
            rptlocation = PathRPT & "SPYC002.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 3 Then
            rptlocation = PathRPT & "SPYC003.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 4 Then
            rptlocation = PathRPT & "SPYC004.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 5 Then  'CARNET DE IDENTIFICACION DE EMPLEADO (administrativo)
            rptlocation = PathRPT & "SPYC005.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01   'NIVEL PROFESIONAL
            Dim V02 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V02.Value = VALOR02   'NOMBRES Y APELLIDOS
            Dim V03 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V03.Value = VALOR03   'CARGO
            Dim V04 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V04.Value = VALOR04   'CODIGO
            Dim V05 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V05.Value = VALOR05   'CEDULA
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V02) : CR.DataDefinition.ParameterFields("P02").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V03) : CR.DataDefinition.ParameterFields("P03").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V04) : CR.DataDefinition.ParameterFields("P04").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V05) : CR.DataDefinition.ParameterFields("P05").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 6 Then  'CARNET DE IDENTIFICACION DE EMPLEADO (personal tecnico)
            rptlocation = PathRPT & "SPYC006.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01   'NIVEL PROFESIONAL
            Dim V02 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V02.Value = VALOR02   'NOMBRES Y APELLIDOS
            Dim V03 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V03.Value = VALOR03   'CARGO
            Dim V04 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V04.Value = VALOR04   'CODIGO
            Dim V05 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V05.Value = VALOR05   'CEDULA
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V02) : CR.DataDefinition.ParameterFields("P02").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V03) : CR.DataDefinition.ParameterFields("P03").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V04) : CR.DataDefinition.ParameterFields("P04").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V05) : CR.DataDefinition.ParameterFields("P05").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 7 Then 'CARNET DE IDENTIFICACION DE EMPLEADO (medicos y militares)
            rptlocation = PathRPT & "SPYC007.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01   'NIVEL PROFESIONAL
            Dim V02 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V02.Value = VALOR02   'NOMBRES Y APELLIDOS
            Dim V03 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V03.Value = VALOR03   'CARGO
            Dim V04 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V04.Value = VALOR04   'CODIGO
            Dim V05 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V05.Value = VALOR05   'CEDULA
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V02) : CR.DataDefinition.ParameterFields("P02").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V03) : CR.DataDefinition.ParameterFields("P03").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V04) : CR.DataDefinition.ParameterFields("P04").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V05) : CR.DataDefinition.ParameterFields("P05").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 8 Then 'CARNET DE IDENTIFICACION DE EMPLEADO (enfermeria)
            rptlocation = PathRPT & "SPYC008.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01   'NIVEL PROFESIONAL
            Dim V02 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V02.Value = VALOR02   'NOMBRES Y APELLIDOS
            Dim V03 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V03.Value = VALOR03   'CARGO
            Dim V04 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V04.Value = VALOR04   'CODIGO
            Dim V05 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V05.Value = VALOR05   'CEDULA
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V02) : CR.DataDefinition.ParameterFields("P02").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V03) : CR.DataDefinition.ParameterFields("P03").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V04) : CR.DataDefinition.ParameterFields("P04").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V05) : CR.DataDefinition.ParameterFields("P05").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 9 Then 'CARNET DE IDENTIFICACION DE EMPLEADO (voluntariado)
            rptlocation = PathRPT & "SPYC009.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01   'NIVEL PROFESIONAL
            Dim V02 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V02.Value = VALOR02   'NOMBRES Y APELLIDOS
            Dim V03 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V03.Value = VALOR03   'CARGO
            Dim V04 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V04.Value = VALOR04   'CODIGO
            Dim V05 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V05.Value = VALOR05   'CEDULA
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V02) : CR.DataDefinition.ParameterFields("P02").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V03) : CR.DataDefinition.ParameterFields("P03").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V04) : CR.DataDefinition.ParameterFields("P04").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V05) : CR.DataDefinition.ParameterFields("P05").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 10 Then 'CARNET DE IDENTIFICACION DE EMPLEADO (policlinico)
            rptlocation = PathRPT & "SPYC010.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01   'NIVEL PROFESIONAL
            Dim V02 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V02.Value = VALOR02   'NOMBRES Y APELLIDOS
            Dim V03 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V03.Value = VALOR03   'CARGO
            Dim V04 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V04.Value = VALOR04   'CODIGO
            Dim V05 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V05.Value = VALOR05   'CEDULA
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V02) : CR.DataDefinition.ParameterFields("P02").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V03) : CR.DataDefinition.ParameterFields("P03").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V04) : CR.DataDefinition.ParameterFields("P04").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V05) : CR.DataDefinition.ParameterFields("P05").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 11 Then
            rptlocation = PathRPT & "SPYC011.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 12 Then
            rptlocation = PathRPT & "SPYC012.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 13 Then
            rptlocation = PathRPT & "SPYC013.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 14 Then
            rptlocation = PathRPT & "SPYC014.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01   'FIRMANTE
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 15 Then
            rptlocation = PathRPT & "SPYC015.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim Ds1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds1.Value = SELECCION_PARAMETRO
            Dim Ds2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : Ds2.Value = USUARIO_IMPRIME
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 16 Then
            rptlocation = PathRPT & "SPYC016.rpt"
            CR.Load(rptlocation)
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 17 Then
            rptlocation = PathRPT & "SPYC017.rpt"
            CR.Load(rptlocation)
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 18 Then
            rptlocation = PathRPT & "SPYC018.rpt"
            CR.Load(rptlocation)
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 19 Then
            rptlocation = PathRPT & "SPYC019.rpt"
            CR.Load(rptlocation)
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 20 Then
            rptlocation = PathRPT & "SPYC020.rpt"
            CR.Load(rptlocation)
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        '*******************************************************************************************
        '*******************************************************************************************
        '*******************************************************************************************
        'ORGANICA
        If PARAMETRO = 100 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 101 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 102 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 103 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 104 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 105 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 106 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 107 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 108 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 109 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 110 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        '*******************************************************************************************
        '*******************************************************************************************
        '*******************************************************************************************
        'MOVIMIENTOS
        If PARAMETRO = 111 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 112 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 113 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 114 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 115 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 116 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 117 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 118 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 119 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 120 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 151 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        '*******************************************************************************************
        '*******************************************************************************************
        '*******************************************************************************************
        'ASPIRANTES
        If PARAMETRO = 121 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        '*******************************************************************************************
        '*******************************************************************************************
        '*******************************************************************************************
        'EXPEDIENTES
        If PARAMETRO = 122 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 123 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 124 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 125 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 126 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 127 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 128 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 129 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 130 Then
            rptlocation = PathRPT & "SPYCE108.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01   'CODIGO
            Dim V02 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V02.Value = VALOR02   'PRONOMBRE
            Dim V03 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V03.Value = VALOR03   'NOMBRES Y APELLIDOS
            Dim V04 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V04.Value = VALOR04   'CEDULA
            Dim V05 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V05.Value = VALOR05   'FECHA INGRESO PAME
            Dim V06 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V06.Value = VALOR06   'CARGO
            Dim V07 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V07.Value = VALOR07   'UBICACION (NIVELES ORGANIZACIONALES)
            Dim V09 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V09.Value = VALOR09   'GRADO JEFE - GRADO 1ER OFICIAL
            Dim V10 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V10.Value = VALOR10   'NOMBRES Y APELLIDOS JEFE - 'NOMBRES Y APELLIDOS 1ER OFICIAL
            Dim V11 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V11.Value = VALOR11   'ASUNTO
            Dim V12 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V12.Value = VALOR12   'FECHA EN LETRAS
            Dim V13 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V13.Value = VALOR13   'FECHA BAJA
            Dim V14 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V14.Value = VALOR14   'SALARIO
            Dim V15 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V15.Value = VALOR15   'TITULO FIRMA
            Dim V16 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V16.Value = VALOR16   'SALARIO LETRAS
            Dim V17 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V17.Value = VALOR17   'SALARIO LETRAS'
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("p01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V02) : CR.DataDefinition.ParameterFields("p02").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V03) : CR.DataDefinition.ParameterFields("p03").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V04) : CR.DataDefinition.ParameterFields("p04").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V05) : CR.DataDefinition.ParameterFields("p05").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V06) : CR.DataDefinition.ParameterFields("p06").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V07) : CR.DataDefinition.ParameterFields("p07").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V09) : CR.DataDefinition.ParameterFields("p09").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V10) : CR.DataDefinition.ParameterFields("p10").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V11) : CR.DataDefinition.ParameterFields("p11").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V12) : CR.DataDefinition.ParameterFields("p12").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V13) : CR.DataDefinition.ParameterFields("p13").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V14) : CR.DataDefinition.ParameterFields("p14").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V15) : CR.DataDefinition.ParameterFields("p15").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V16) : CR.DataDefinition.ParameterFields("p16").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V17) : CR.DataDefinition.ParameterFields("p17").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 131 Then
            rptlocation = PathRPT & "SPYCE109.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01   'CODIGO
            Dim V02 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V02.Value = VALOR02   'PRONOMBRE
            Dim V03 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V03.Value = VALOR03   'NOMBRES Y APELLIDOS
            Dim V04 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V04.Value = VALOR04   'CEDULA
            Dim V05 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V05.Value = VALOR05   'FECHA INGRESO PAME
            Dim V06 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V06.Value = VALOR06   'CARGO
            Dim V07 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V07.Value = VALOR07   'DEPARTAMENTO
            Dim V09 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V09.Value = VALOR09   'GRADO JEFE - GRADO 1ER OFICIAL
            Dim V10 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V10.Value = VALOR10   'NOMBRES Y APELLIDOS JEFE - 'NOMBRES Y APELLIDOS 1ER OFICIAL
            Dim V11 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V11.Value = VALOR11   'ASUNTO
            Dim V12 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V12.Value = VALOR12   'FECHA EN LETRAS
            Dim V13 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V13.Value = VALOR13   'FECHA BAJA
            Dim V14 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V14.Value = VALOR14   'SALARIO
            Dim V15 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V15.Value = VALOR15   'TITULO FIRMA
            Dim V16 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V16.Value = VALOR16   'SALARIO LETRAS
            Dim V17 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V17.Value = VALOR17  'A QUIEN DIRIGE'
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("p01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V02) : CR.DataDefinition.ParameterFields("p02").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V03) : CR.DataDefinition.ParameterFields("p03").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V04) : CR.DataDefinition.ParameterFields("p04").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V05) : CR.DataDefinition.ParameterFields("p05").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V06) : CR.DataDefinition.ParameterFields("p06").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V07) : CR.DataDefinition.ParameterFields("p07").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V09) : CR.DataDefinition.ParameterFields("p09").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V10) : CR.DataDefinition.ParameterFields("p10").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V11) : CR.DataDefinition.ParameterFields("p11").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V12) : CR.DataDefinition.ParameterFields("p12").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V13) : CR.DataDefinition.ParameterFields("p13").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V14) : CR.DataDefinition.ParameterFields("p14").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V15) : CR.DataDefinition.ParameterFields("p15").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V16) : CR.DataDefinition.ParameterFields("p16").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V17) : CR.DataDefinition.ParameterFields("p17").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 132 Then
            rptlocation = PathRPT & "SPYCE110.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01   'CODIGO
            Dim V02 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V02.Value = VALOR02   'PRONOMBRE
            Dim V03 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V03.Value = VALOR03   'NOMBRES Y APELLIDOS
            Dim V04 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V04.Value = VALOR04   'CEDULA
            Dim V05 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V05.Value = VALOR05   'FECHA INGRESO PAME
            Dim V06 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V06.Value = VALOR06   'CARGO
            Dim V07 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V07.Value = VALOR07   'UBICACION (NIVELES ORGANIZACIONALES)
            Dim V09 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V09.Value = VALOR09   'GRADO JEFE - GRADO 1ER OFICIAL
            Dim V10 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V10.Value = VALOR10   'NOMBRES Y APELLIDOS JEFE - 'NOMBRES Y APELLIDOS 1ER OFICIAL
            Dim V11 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V11.Value = VALOR11   'ASUNTO
            Dim V12 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V12.Value = VALOR12   'FECHA EN LETRAS
            Dim V13 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V13.Value = VALOR13   'FECHA BAJA
            Dim V14 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V14.Value = VALOR14   'SALARIO
            Dim V15 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V15.Value = VALOR15   'TITULO FIRMA
            Dim V16 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V16.Value = VALOR16   'SALARIO LETRAS
            Dim V17 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V17.Value = VALOR17  'A QUIEN DIRIGE'
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("p01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V02) : CR.DataDefinition.ParameterFields("p02").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V03) : CR.DataDefinition.ParameterFields("p03").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V04) : CR.DataDefinition.ParameterFields("p04").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V05) : CR.DataDefinition.ParameterFields("p05").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V06) : CR.DataDefinition.ParameterFields("p06").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V07) : CR.DataDefinition.ParameterFields("p07").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V09) : CR.DataDefinition.ParameterFields("p09").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V10) : CR.DataDefinition.ParameterFields("p10").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V11) : CR.DataDefinition.ParameterFields("p11").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V12) : CR.DataDefinition.ParameterFields("p12").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V13) : CR.DataDefinition.ParameterFields("p13").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V14) : CR.DataDefinition.ParameterFields("p14").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V15) : CR.DataDefinition.ParameterFields("p15").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V16) : CR.DataDefinition.ParameterFields("p16").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V17) : CR.DataDefinition.ParameterFields("p17").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 133 Then
            rptlocation = PathRPT & "SPYCE111.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01   'CODIGO
            Dim V02 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V02.Value = VALOR02   'PRONOMBRE
            Dim V03 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V03.Value = VALOR03   'NOMBRES Y APELLIDOS
            Dim V04 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V04.Value = VALOR04   'CEDULA
            Dim V05 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V05.Value = VALOR05   'FECHA INGRESO PAME
            Dim V06 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V06.Value = VALOR06   'CARGO
            Dim V07 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V07.Value = VALOR07   'UBICACION (NIVELES ORGANIZACIONALES)
            Dim V09 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V09.Value = VALOR09   'GRADO JEFE - GRADO 1ER OFICIAL
            Dim V10 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V10.Value = VALOR10   'NOMBRES Y APELLIDOS JEFE - 'NOMBRES Y APELLIDOS 1ER OFICIAL
            Dim V11 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V11.Value = VALOR11   'ASUNTO
            Dim V12 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V12.Value = VALOR12   'FECHA EN LETRAS
            Dim V13 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V13.Value = VALOR13   'FECHA BAJA
            Dim V14 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V14.Value = VALOR14   'SALARIO
            Dim V15 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V15.Value = VALOR15   'TITULO FIRMA
            Dim V16 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V16.Value = VALOR16   'SALARIO LETRAS
            Dim V17 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V17.Value = VALOR17   'A QUIEN DIRIGE'
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("p01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V02) : CR.DataDefinition.ParameterFields("p02").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V03) : CR.DataDefinition.ParameterFields("p03").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V04) : CR.DataDefinition.ParameterFields("p04").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V05) : CR.DataDefinition.ParameterFields("p05").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V06) : CR.DataDefinition.ParameterFields("p06").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V07) : CR.DataDefinition.ParameterFields("p07").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V09) : CR.DataDefinition.ParameterFields("p09").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V10) : CR.DataDefinition.ParameterFields("p10").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V11) : CR.DataDefinition.ParameterFields("p11").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V12) : CR.DataDefinition.ParameterFields("p12").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V13) : CR.DataDefinition.ParameterFields("p13").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V14) : CR.DataDefinition.ParameterFields("p14").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V15) : CR.DataDefinition.ParameterFields("p15").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V16) : CR.DataDefinition.ParameterFields("p16").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V17) : CR.DataDefinition.ParameterFields("p17").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 150 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 152 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            SELECCION = "(({MAESTRO_DE_CARGOS.ID_SITUACION} = 1))"
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 154 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            'RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 155 Then
            rptlocation = PathRPT & "SPYCE115.rpt"
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        '*******************************************************************************************
        '*******************************************************************************************
        '*******************************************************************************************
        'MEDICOS ACREDITADOS
        If PARAMETRO = 134 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        '*******************************************************************************************
        '*******************************************************************************************
        '*******************************************************************************************
        'BECADOS
        If PARAMETRO = 135 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        '*******************************************************************************************
        '*******************************************************************************************
        '*******************************************************************************************
        'SITUACION DE PERSONAL
        If PARAMETRO = 136 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            Dim DsCODIGO As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsCODIGO.Value = VALOR01
            Dim DsNOMBRESAPELLIDOS As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsNOMBRESAPELLIDOS.Value = VALOR02
            Dim DsCARGO As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsCARGO.Value = VALOR03
            Dim DsF1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF1.Value = VALOR04
            Dim DsF2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF2.Value = VALOR05
            Dim DsUBICACION As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsUBICACION.Value = VALOR06
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsCODIGO) : CR.DataDefinition.ParameterFields("pCODIGO").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsNOMBRESAPELLIDOS) : CR.DataDefinition.ParameterFields("pNOMBRESAPELLIDOS").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsCARGO) : CR.DataDefinition.ParameterFields("pCARGO").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsF1) : CR.DataDefinition.ParameterFields("pF1").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsF2) : CR.DataDefinition.ParameterFields("pF2").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsUBICACION) : CR.DataDefinition.ParameterFields("pUBICACION").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 137 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            Dim DsCODIGO As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsCODIGO.Value = VALOR01
            Dim DsNOMBRESAPELLIDOS As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsNOMBRESAPELLIDOS.Value = VALOR02
            Dim DsCARGO As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsCARGO.Value = VALOR03
            Dim DsF1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF1.Value = VALOR04
            Dim DsF2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF2.Value = VALOR05
            Dim DsUBICACION As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsUBICACION.Value = VALOR06
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsCODIGO) : CR.DataDefinition.ParameterFields("pCODIGO").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsNOMBRESAPELLIDOS) : CR.DataDefinition.ParameterFields("pNOMBRESAPELLIDOS").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsCARGO) : CR.DataDefinition.ParameterFields("pCARGO").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsF1) : CR.DataDefinition.ParameterFields("pF1").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsF2) : CR.DataDefinition.ParameterFields("pF2").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsUBICACION) : CR.DataDefinition.ParameterFields("pUBICACION").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 138 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            VALOR04 = Mid(Frm_Seleccion_Consultas_y_Reportes.DateTimePicker1.Value, 1, 10)
            VALOR05 = Mid(Frm_Seleccion_Consultas_y_Reportes.DateTimePicker2.Value, 1, 10)
            Dim DsF1 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF1.Value = VALOR04
            Dim DsF2 As New CrystalDecisions.Shared.ParameterDiscreteValue() : DsF2.Value = VALOR05
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsF1) : CR.DataDefinition.ParameterFields("pF1").ApplyCurrentValues(RpDatos)
            RpDatos.Add(DsF2) : CR.DataDefinition.ParameterFields("pF2").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 139 Then
            'rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            'CR.Load(rptlocation)
            'If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            'CrystalR.ReportSource = CR
            'GoTo SALTO

            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = VALOR01
            Dim V02 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V02.Value = VALOR02
            Dim V03 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V03.Value = VALOR03
            Dim V04 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V04.Value = VALOR04
            Dim V05 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V05.Value = VALOR05
            Dim V06 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V06.Value = VALOR06
            Dim V07 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V07.Value = VALOR07
            Dim V08 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V08.Value = VALOR08
            Dim V09 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V09.Value = VALOR09
            Dim V10 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V10.Value = VALOR10
            Dim V11 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V11.Value = VALOR11
            Dim V12 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V12.Value = VALOR12
            Dim V13 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V13.Value = VALOR13
            Dim V14 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V14.Value = VALOR14
            Dim V15 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V15.Value = VALOR15
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("p01").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V02) : CR.DataDefinition.ParameterFields("p02").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V03) : CR.DataDefinition.ParameterFields("p03").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V04) : CR.DataDefinition.ParameterFields("p04").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V05) : CR.DataDefinition.ParameterFields("p05").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V06) : CR.DataDefinition.ParameterFields("p06").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V07) : CR.DataDefinition.ParameterFields("p07").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V08) : CR.DataDefinition.ParameterFields("p08").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V09) : CR.DataDefinition.ParameterFields("p09").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V10) : CR.DataDefinition.ParameterFields("p10").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V11) : CR.DataDefinition.ParameterFields("p11").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V12) : CR.DataDefinition.ParameterFields("p12").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V13) : CR.DataDefinition.ParameterFields("p13").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V14) : CR.DataDefinition.ParameterFields("p14").ApplyCurrentValues(RpDatos)
            RpDatos.Add(V15) : CR.DataDefinition.ParameterFields("p15").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 140 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 141 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 142 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 143 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 144 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 145 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 146 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 147 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 148 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 149 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
        If PARAMETRO = 156 Then
            ' Ruta del reporte
            ' Ruta del reporte
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)

            ' Define las claves válidas
            Dim clavesValidas As New HashSet(Of String)({"FechaSP", "SituacionAusenteSP"})
            Dim Filtro_SituacionB As String = ""

            ' Itera sobre el diccionario
            For Each kvp As KeyValuePair(Of String, String) In DiccionarioParametros
                ' Valida si la clave está permitida
                If Not clavesValidas.Contains(kvp.Key) Then
                    ' Asigna el valor de Filtro_SituacionB si es la clave "Filtro_Situacion"
                    If kvp.Key = "Filtro_Situacion" Then
                        Filtro_SituacionB = kvp.Value
                    End If
                    Continue For
                End If

                Dim parametroValores As New CrystalDecisions.Shared.ParameterValues()
                Dim parametroValor As New CrystalDecisions.Shared.ParameterDiscreteValue()

                ' Determina qué parámetro estás configurando en base a la clave
                Select Case kvp.Key
                    Case "FechaSP"
                        parametroValor.Value = CDate(kvp.Value & " 00:00:00")
                    Case "SituacionAusenteSP"
                        parametroValor.Value = kvp.Value
                    Case Else
                        parametroValor.Value = kvp.Value
                End Select

                parametroValores.Add(parametroValor)
                CR.DataDefinition.ParameterFields(kvp.Key).ApplyCurrentValues(parametroValores)
            Next

            ' Obtener todas las secciones del reporte
            Dim secciones As CrystalDecisions.CrystalReports.Engine.Sections = CR.ReportDefinition.Sections

            ' Lógica para recorrer todas las secciones
            For i As Integer = 0 To secciones.Count - 1
                Dim seccion As CrystalDecisions.CrystalReports.Engine.Section = secciones(i)

                ' Verificar si Filtro_SituacionB es "Consolidado Jerárquico"
                If Filtro_SituacionB = "Consolidado Jerárquico" Then
                    ' Ocultar las secciones en las posiciones 2 y 5
                    If i = 2 Or i = 5 Then
                        seccion.SectionFormat.EnableSuppress = True ' Ocultar la sección
                    ElseIf i = 3 Or i = 4 Then
                        seccion.SectionFormat.EnableSuppress = False ' Mostrar la sección
                    End If
                Else
                    ' Ocultar las secciones en las posiciones 3 y 4
                    If i = 2 Or i = 5 Then
                        seccion.SectionFormat.EnableSuppress = False ' Mostrar la sección
                    ElseIf i = 3 Or i = 4 Then
                        seccion.SectionFormat.EnableSuppress = True ' Ocultar la sección
                    End If
                End If
            Next

            ' Credenciales para la base de datos
            CR.SetDatabaseLogon("sa", "P@$$W0RD")

            ' Asignación al visor
            CrystalR.ReportSource = CR
        End If

        '*******************************************************************************************
        '*******************************************************************************************
        '*******************************************************************************************
        'HIGIENE Y SEGURIDAD
        If PARAMETRO = 153 Then
            rptlocation = PathRPT & NOMBRE_DEL_REPORTE_CR
            CR.Load(rptlocation)
            Dim RpDatos As New CrystalDecisions.Shared.ParameterValues()
            Dim V01 As New CrystalDecisions.Shared.ParameterDiscreteValue() : V01.Value = FILTRO_SELECCION_EN_REPORTES
            SELECCION = "(({MAESTRO_DE_CARGOS.ID_SITUACION} = 1))"
            If SELECCION = "" Then : CR.RecordSelectionFormula = "" : Else : CR.RecordSelectionFormula = SELECCION : End If
            RpDatos.Add(V01) : CR.DataDefinition.ParameterFields("P01").ApplyCurrentValues(RpDatos)
            CR.SetDatabaseLogon("sa", "P@$$W0RD")
            CrystalR.ReportSource = CR
            GoTo SALTO
        End If
SALTO:
        CrystalR.ReportSource = CR
        Exit Sub
SALTO2:
        Me.Close()
    End Sub

    Friend Shared Function CrystalReportViewer1() As Object
        Throw New NotImplementedException()
    End Function
End Class