Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Medicos_Acreditados_Buscar
    Private Sub Frm_Medicos_Acreditados_Buscar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox1.Text = "0"
            Me.TextBox2.Text = "Buscar Medico Acreditado..."
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Medicos_Acreditados_Buscar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = "Buscar Medico Acreditado..."
        Call CARGAR_DATAGRIDVIEW()
        Call ARMAR_DATAGRIDVIEW()
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
    End Sub

    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = "Buscar Medico Acreditado..."
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
        Me.Close()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        If Len(Me.TextBox2.Text) = 0 Then
            MsgBox("Debe digitar correctamente el Texto que desea Buscar", vbInformation, "Mensaje del sistema")
            Me.TextBox2.Focus()
        Else
            Call CARGAR_DATAGRIDVIEW()
            Call ARMAR_DATAGRIDVIEW()
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
        End If
    End Sub
    Dim SELECCION As String
    Private Sub CARGAR_DATAGRIDVIEW()
        SELECCION = "ID_MA, APELLIDOS, NOMBRES, ID_E_MED, D_E_MED, N_CEDULA, N_MINSA, TELEFONO_1, TELEFONO_2, CORREO_1, CORREO_2, ID_NACIONALIDAD_D, D_NACIONALIDAD_D, DIRECCION_D, ID_T_SEXO, D_T_SEXO, FEC_NAC, FEC_DESDE, FEC_HASTA, FACEBOOK, TWITTER, WHATSAPP, OBSERVACIONES, AN"
        If Me.CheckBox1.Checked = False Then
            Dim buscar As String = ""
            Dim tamañoArreglo As Integer = 0
            buscar = Me.TextBox2.Text.Trim
            Dim ArrCadena As String() = buscar.Split(" ")
            tamañoArreglo = ArrCadena.GetUpperBound(0) + 1
            If tamañoArreglo = 0 Then
                CADENAsql = "SELECT " & SELECCION & " FROM [dbo].[VISTA MAESTRO DE MEDICOS ACREDITADOS] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%')"
            ElseIf tamañoArreglo = 1 Then
                CADENAsql = "SELECT " & SELECCION & " FROM [dbo].[VISTA MAESTRO DE MEDICOS ACREDITADOS] WHERE AN LIKE '%" & buscar & "%'"
            ElseIf tamañoArreglo > 1 Then
                CADENAsql = "SELECT " & SELECCION & " FROM [dbo].[VISTA MAESTRO DE MEDICOS ACREDITADOS] WHERE AN LIKE '%" & ArrCadena(0) & "%' "
                For index = 0 To ArrCadena.GetUpperBound(0)
                    CADENAsql = CADENAsql & " UNION SELECT " & SELECCION & " FROM [dbo].[VISTA MAESTRO DE MEDICOS ACREDITADOS] WHERE AN LIKE '%" & ArrCadena(index) & "%'"
                Next
            End If
            CADENAsql = CADENAsql & " ORDER BY AN"
        End If
        If Me.CheckBox1.Checked = True Then
            CADENAsql = "Select " & SELECCION & " from [VISTA MAESTRO DE MEDICOS ACREDITADOS] WHERE (AN LIKE '%" & Me.TextBox2.Text & "%') ORDER BY AN"
        End If
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiAdaptador As New SqlClient.SqlDataAdapter(CADENAsql, Conexion.CxRRHH)
        Dim MiDataSet As New DataSet()
        MiAdaptador.Fill(MiDataSet, "[VISTA MAESTRO DE MEDICOS ACREDITADOS]")
        DGV.DataSource = MiDataSet.Tables("[VISTA MAESTRO DE MEDICOS ACREDITADOS]")
        Me.Label2.Text = "Total de Registros: " & DGV.RowCount
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
    End Sub
    Private Sub ARMAR_DATAGRIDVIEW()
        Me.DGV.Columns(0).HeaderText = "Id"
        Me.DGV.Columns(0).Width = 30
        Me.DGV.Columns(0).Visible = True
        Me.DGV.Columns(0).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(0).DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight
        Me.DGV.Columns(0).DefaultCellStyle.BackColor = Color.Black

        Me.DGV.Columns(1).HeaderText = "Apellidos"
        Me.DGV.Columns(1).Width = 150
        Me.DGV.Columns(1).Visible = True
        Me.DGV.Columns(1).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(2).HeaderText = "Nombres"
        Me.DGV.Columns(2).Width = 150
        Me.DGV.Columns(2).Visible = TRUE
        Me.DGV.Columns(2).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(3).HeaderText = "ID_E_MED"
        Me.DGV.Columns(3).Width = 0
        Me.DGV.Columns(3).Visible = False

        Me.DGV.Columns(4).HeaderText = "Especialidad Medica"
        Me.DGV.Columns(4).Width = 0
        Me.DGV.Columns(4).Visible = False
        Me.DGV.Columns(4).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(5).HeaderText = "Cedula"
        Me.DGV.Columns(5).Width = 120
        Me.DGV.Columns(5).Visible = True
        Me.DGV.Columns(5).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(5).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(6).HeaderText = "No. MINSA"
        Me.DGV.Columns(6).Width = 80
        Me.DGV.Columns(6).Visible = True
        Me.DGV.Columns(6).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter
        Me.DGV.Columns(6).DefaultCellStyle.BackColor = Color.Khaki

        Me.DGV.Columns(7).HeaderText = "TELEFONO_1"
        Me.DGV.Columns(7).Width = 0
        Me.DGV.Columns(7).Visible = False

        Me.DGV.Columns(8).HeaderText = "TELEFONO_2"
        Me.DGV.Columns(8).Width = 0
        Me.DGV.Columns(8).Visible = False

        Me.DGV.Columns(9).HeaderText = "CORREO_1"
        Me.DGV.Columns(9).Width = 0
        Me.DGV.Columns(9).Visible = False

        Me.DGV.Columns(10).HeaderText = "CORREO_2"
        Me.DGV.Columns(10).Width = 0
        Me.DGV.Columns(10).Visible = False

        Me.DGV.Columns(11).HeaderText = "ID_NACIONALIDAD_D"
        Me.DGV.Columns(11).Width = 0
        Me.DGV.Columns(11).Visible = False

        Me.DGV.Columns(12).HeaderText = "Nacionalidad"
        Me.DGV.Columns(12).Width = 0
        Me.DGV.Columns(12).Visible = False
        Me.DGV.Columns(12).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(13).HeaderText = "DIRECCION_D"
        Me.DGV.Columns(13).Width = 0
        Me.DGV.Columns(13).Visible = False

        Me.DGV.Columns(14).HeaderText = "ID_T_SEXO"
        Me.DGV.Columns(14).Width = 0
        Me.DGV.Columns(14).Visible = False

        Me.DGV.Columns(15).HeaderText = "Sexo"
        Me.DGV.Columns(15).Width = 0
        Me.DGV.Columns(15).Visible = False
        Me.DGV.Columns(15).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(16).HeaderText = "Fecha Nacimiento"
        Me.DGV.Columns(16).Width = 0
        Me.DGV.Columns(16).Visible = False
        Me.DGV.Columns(16).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(17).HeaderText = "Fecha Desde"
        Me.DGV.Columns(17).Width = 70
        Me.DGV.Columns(17).Visible = True
        Me.DGV.Columns(17).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(18).HeaderText = "Fecha Hasta"
        Me.DGV.Columns(18).Width = 70
        Me.DGV.Columns(18).Visible = True
        Me.DGV.Columns(18).HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter

        Me.DGV.Columns(19).HeaderText = "FACEBOOK"
        Me.DGV.Columns(19).Width = 0
        Me.DGV.Columns(19).Visible = False

        Me.DGV.Columns(20).HeaderText = "TWITTER"
        Me.DGV.Columns(20).Width = 0
        Me.DGV.Columns(20).Visible = False

        Me.DGV.Columns(21).HeaderText = "WHATSAPP"
        Me.DGV.Columns(21).Width = 0
        Me.DGV.Columns(21).Visible = False

        Me.DGV.Columns(22).HeaderText = "OBSERVACIONES"
        Me.DGV.Columns(22).Width = 0
        Me.DGV.Columns(22).Visible = False

        Me.DGV.Columns(23).HeaderText = "AN"
        Me.DGV.Columns(23).Width = 0
        Me.DGV.Columns(23).Visible = False

        'Me.DGV.Columns(24).HeaderText = "FOTO"
        'Me.DGV.Columns(24).Width = 0
        'Me.DGV.Columns(24).Visible = False

    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If (e.KeyCode = Keys.Enter) Then
            If Len(Me.TextBox2.Text) = 0 Then
                MsgBox("Debe digitar correctamente el Texto que desea Buscar", vbInformation, "Mensaje del sistema")
                Me.TextBox2.Focus()
            Else
                Call CARGAR_DATAGRIDVIEW()
                Call ARMAR_DATAGRIDVIEW()
                Me.TextBox2.SelectAll()
                Me.TextBox2.Focus()
            End If
        End If
    End Sub
    Private Sub DGV_Click(sender As Object, e As EventArgs) Handles DGV.Click
        If Me.DGV.RowCount > 0 Then
            For I = 0 To Me.DGV.RowCount - 1
                If DGV.Rows(I).Selected = True Then
                    vMAc_ID_MA = Me.DGV.Item(0, I).Value
                    Exit Sub
                End If
            Next
        End If
    End Sub
    Private Sub DGV_DoubleClick(sender As Object, e As EventArgs) Handles DGV.DoubleClick
        If vMAc_ID_MA <> 0 Then
            Me.TextBox1.Text = "0"
            Me.TextBox2.Text = "Buscar Medico Acreditado..."
            Me.TextBox2.SelectAll()
            Me.TextBox2.Focus()
            CERRAR = False
            Me.Close()
        End If
    End Sub
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Me.TextBox2.SelectAll()
        Me.TextBox2.Focus()
    End Sub
End Class