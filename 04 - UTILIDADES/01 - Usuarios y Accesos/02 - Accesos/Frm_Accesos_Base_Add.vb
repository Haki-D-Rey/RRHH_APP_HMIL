Imports System
Imports System.Data
Imports System.Data.SqlClient
Public Class Frm_Accesos_Base_Add
    Private Sub Frm_Accesos_Base_Add_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox3.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Accesos_Base_Add_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Agregar Registro a --->" & NODO_TEXT
        Me.TextBox1.Text = "0"
        Me.TextBox3.Text = ""
        Me.CheckBox1.Checked = False
        Me.TextBox3.Focus()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Me.TextBox3.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub CheckBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles CheckBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Dim ID_CODIGO_U As Integer
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ''Call VERIFICAR_ACCESOS("037") : If HAY_ACCESO = False Then : Exit Sub : End If
        If Len(TextBox3.Text) = 0 Then
            MsgBox("Debe digitar la Descricpion del Nodo que desea Agregar", vbInformation, "Mensaje del Sistema")
            TextBox3.Focus()
            Exit Sub
        End If
        Call GENERAR_ID_MAESTRO_DE_NODOS_BASE()
        Call AGREGAR_REGISTRO_EN_MAESTRO_DE_NODOS_BASE()
        '------------------------------------------
        Call GENERAR_ID_MAESTRO_DE_NODOS_USUARIOS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_USUARIO FROM [MAESTRO DE USUARIOS] ORDER BY ID_USUARIO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    ID_CODIGO_U = MiDataTable.Rows(I).Item(0).ToString
                    Call AGREGAR_REGISTRO_EN_MAESTRO_DE_NODOS_USUARIOS()
                Next
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        Me.TextBox3.Focus()
        CERRAR = False
        Me.Close()
    End Sub
    '---------------------------------------------------------------------------------------------------------
    Private Sub AGREGAR_REGISTRO_EN_MAESTRO_DE_NODOS_USUARIOS()
        Dim NA As Integer
        If Me.CheckBox1.Checked = True Then
            NA = 1
        Else
            NA = 0
        End If
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()

        Dim SQL As String
        Try
            SQL = "INSERT INTO [MAESTRO DE NODOS USUARIOS] (ID_NODO, DESCRIPCION, ID_NODO_PADRE, ACCESO, CLAVE, ID_USUARIO)" &
                  "values (" & CInt(ID_NODO1) & ",  '" & Trim(Me.TextBox3.Text) & "', " & NODO_TAG & ", " & NA & ", '" & Format(Val(ID_NODO1), "000") & "', " & ID_CODIGO_U & ")"

            Dim cmd As New SqlCommand(SQL, CxRRHH)
            cmd.ExecuteNonQuery()
            cmd = Nothing
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Dim ID_NODO1 As Integer
    Private Sub GENERAR_ID_MAESTRO_DE_NODOS_USUARIOS()
        Dim CODIGO As Integer
        CODIGO = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_NODO FROM [MAESTRO DE NODOS USUARIOS] ORDER BY ID_NODO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    CODIGO = MiDataTable.Rows(I).Item(0).ToString
                Next
                CODIGO += 1
            Else
                CODIGO = 1
            End If
            ID_NODO1 = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    '---------------------------------------------------------------------------------------------------------
    Private Sub AGREGAR_REGISTRO_EN_MAESTRO_DE_NODOS_BASE()
        Dim NA As Integer
        If Me.CheckBox1.Checked = True Then
            NA = 1
        Else
            NA = 0
        End If
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()

        Dim SQL As String
        Try
            SQL = "INSERT INTO [MAESTRO DE NODOS BASE] (ID_NODO, DESCRIPCION, ID_NODO_PADRE, ACCESO, CLAVE)" &
                  "values (" & CInt(Me.TextBox1.Text) & ",  '" & Trim(Me.TextBox3.Text) & "', " & NODO_TAG & ", " & NA & ", '" & Format(Val(Me.TextBox1.Text), "000") & "')"

            Dim cmd As New SqlCommand(SQL, CxRRHH)
            cmd.ExecuteNonQuery()
            cmd = Nothing
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub GENERAR_ID_MAESTRO_DE_NODOS_BASE()
        Dim CODIGO As Integer
        CODIGO = 0
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT ID_NODO FROM [MAESTRO DE NODOS BASE] ORDER BY ID_NODO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                For I = 0 To MiDataTable.Rows.Count - 1
                    CODIGO = MiDataTable.Rows(I).Item(0).ToString
                Next
                CODIGO += 1
            Else
                CODIGO = 1
            End If
            TextBox1.Text = CODIGO
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
End Class