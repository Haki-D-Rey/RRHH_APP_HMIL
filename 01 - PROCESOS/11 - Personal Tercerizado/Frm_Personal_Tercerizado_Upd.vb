Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Personal_Tercerizado_Upd
    Private Sub Frm_Personal_Tercerizado_Upd_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.ComboBox1.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Personal_Tercerizado_Upd_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Me.TextBox1.Text = ptID_P_T
        Call CARGAR_COMBO_CAT_DE_EMPRESAS_TERCERIZADAS()
        Me.ComboBox1.Text = ptD_E_T
        Me.TextBox2.Text = ptCODIGO
        Me.TextBox11.Text = ptAPELLIDOS
        Me.TextBox3.Text = ptNOMBRES
        Me.TextBox4.Text = ptN_CEDULA
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        Me.ComboBox2.Text = ptD_T_SEXO
        Me.TextBox6.Text = ptTELEFONO_1
        Me.TextBox5.Text = ptTELEFONO_2
        Me.TextBox7.Text = ptD_CARGO_ES
        Call CARGAR_COMBO_CAT_DE_ESTADO_DE_PERSONAS()
        Me.ComboBox4.Text = ptD_ESTADO_P
        Me.TextBox8.Text = ptNSS
        Me.TextBox9.Text = ptDIRECCION
        Me.TextBox14.Text = ptOBSERVACIONES


        If ptF_INGRESO <> "" Then
            Me.DateTimePicker1.Value = ptF_INGRESO
            Me.DateTimePicker1.Checked = True
        Else
            Me.DateTimePicker1.Checked = False
        End If

        If ptF_BAJA <> "" Then
            Me.DateTimePicker2.Value = ptF_BAJA
            Me.DateTimePicker2.Checked = True
        Else
            Me.DateTimePicker2.Checked = False
        End If


        Call CARGAR_FOTO()
        Me.ComboBox1.Focus()
    End Sub
    Private Sub CARGAR_FOTO()
        Dim DT As New DataTable
        Dim DA As New SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            CADENAsql = "SELECT * FROM [MAESTRO DE PERSONAL TERCERIZADO] WHERE ID_P_T = " & ptID_P_T & " ORDER BY ID_P_T"
            DA = New SqlDataAdapter(CADENAsql, CxRRHH)
            DA.Fill(DT)
            If DT.Rows.Count > 0 Then
                Dim ms As New System.IO.MemoryStream()
                Dim imageBuffer() As Byte = CType(DT.Rows(0).Item("FOTO"), Byte())
                ms = New System.IO.MemoryStream(imageBuffer)
                Me.PictureBox4.Image = Nothing
                Me.PictureBox4.Image = (Image.FromStream(ms))
                Me.PictureBox4.BackgroundImageLayout = ImageLayout.Stretch
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
        DT.Clear()
        DT.Reset()
    End Sub
    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.ComboBox1.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESTADO_DE_PERSONAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESTADO DE PERSONAS] ORDER BY DESCRIPCION", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox4
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_ESTADO_P"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_EMPRESAS_TERCERIZADAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE EMPRESAS TERCERIZADAS] ORDER BY ID_E_T", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_E_T"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE SEXO] ORDER BY ID_T_SEXO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_T_SEXO"
            End With

        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox11_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox11.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox6_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox6.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub ComboBox4_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox4.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox14_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox14.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox7_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox7.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox8_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox8.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox9_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox9.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Empresa Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If TextBox11.Text = "" Then
            MsgBox("Debe Digitar los Apellidos Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox11.Focus()
            Exit Sub
        End If
        If TextBox3.Text = "" Then
            MsgBox("Debe Digitar los Nombres Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If TextBox4.Text = "" Then
            MsgBox("Debe Digitar la Cedula Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Sexo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox4.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Estado de la Persona Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox4.Focus()
            Exit Sub
        End If
        If TextBox14.Text = "" Then
            Me.TextBox14.Text = "."
        End If
        Call Me.EDITAR_REGISTRO()
        CERRAR = False
        ptID_P_T = Me.TextBox1.Text
        Me.Close()
    End Sub
    Private Sub EDITAR_REGISTRO()
        Dim fFEC_INGRESO As String
        If Me.DateTimePicker1.Checked = True Then
            fFEC_INGRESO = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFEC_INGRESO = "NULL"
        End If
        Dim fechaingreso As Object = If(fFEC_INGRESO <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_INGRESO + "', 105), 23)", "NULL")

        Dim fFEC_BAJA As String
        If Me.DateTimePicker2.Checked = True Then
            fFEC_BAJA = Mid(Me.DateTimePicker2.Value, 1, 10)
        Else
            fFEC_BAJA = "NULL"
        End If
        Dim fechabaja As Object = If(fFEC_BAJA <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_BAJA + "', 105), 23)", "NULL")


        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Dim SQL As String
        Dim ms As New System.IO.MemoryStream() 'creando el memoryStream
        Try
            Me.PictureBox4.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            SQL = "UPDATE [MAESTRO DE PERSONAL TERCERIZADO] SET ID_E_T = " & Me.ComboBox1.SelectedValue & ", APELLIDOS = '" & Me.TextBox11.Text & "', NOMBRES = '" & Me.TextBox3.Text & "', N_CEDULA = '" & Me.TextBox4.Text & "', ID_T_SEXO = " & Me.ComboBox2.SelectedValue & ", TELEFONO_1 = '" & Me.TextBox6.Text & "', TELEFONO_2 = '" & Me.TextBox5.Text & "', D_CARGO_ES = '" & Me.TextBox7.Text & "', ID_ESTADO_P = " & Me.ComboBox4.SelectedValue & ", DIRECCION = '" & Me.TextBox9.Text & "', NSS = '" & Me.TextBox8.Text & "', OBSERVACIONES = '" & Me.TextBox14.Text & "', FOTO = @data, F_INGRESO = " & fechaingreso & ", F_BAJA = " & fechabaja & ", CAUSA_RETIRO = '" & Me.TextBox10.Text & "' WHERE ID_P_T = " & CInt(Me.TextBox1.Text) & ""
            Dim cmd As New SqlCommand(SQL, CxRRHH)
            cmd.Parameters.Add(New SqlParameter("@data", ms.GetBuffer()))
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
    Dim CONTROL_FOTO As Boolean
    Dim data() As Byte 'arreglo de bytes el cual contedra la imagen
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim ios As FileStream 'Manejo de archivos
        Try
            Me.OpenFileDialog1.Filter = "Imagenes (JPG)|*.jpg" 'filtro de archivos del OpenFileDialog 
            If Me.OpenFileDialog1.ShowDialog() = Windows.Forms.DialogResult.Cancel Then ' en caso de que se aplaste el boton cancelar salga y no haga nada
                CONTROL_FOTO = False
                Exit Sub
            Else
                ios = New FileStream(Me.OpenFileDialog1.FileName, FileMode.Open, FileAccess.Read) 'instanciamos en Stream indicandole la ruta del arvhivo,el modo de acceso y si es de lectura o escritura
                ReDim data(ios.Length) 'llenamos el arreglo
                ios.Read(data, 0, CInt(ios.Length)) 'llenamos el arreglo
                Me.PictureBox4.SizeMode = PictureBoxSizeMode.StretchImage 'establecemos como se visualiza la imagen
                Me.PictureBox4.Load(Me.OpenFileDialog1.FileName) 'visualizamos abriendo el archivo seleccionado
                CONTROL_FOTO = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification) 'en caso de error muestre un mensaje
        End Try
    End Sub
    Private Sub DateTimePicker1_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker1.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub DateTimePicker2_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker2.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox10_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox10.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
End Class