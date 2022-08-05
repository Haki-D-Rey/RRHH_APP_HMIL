Imports System
Imports System.ComponentModel
Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Public Class Frm_Medicos_Acreditados_Editar
    Private Sub Frm_Medicos_Acreditados_Editar_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If (e.KeyCode = Keys.Escape) Then
            Me.TextBox2.Focus()
            CERRAR = True
            Me.Close()
        End If
    End Sub
    Private Sub Frm_Medicos_Acreditados_Editar_Load(sender As Object, e As System.EventArgs) Handles MyBase.Load
        Call INICIAR_OBJETOS()
        Me.TextBox6.Focus()
    End Sub

    Private Sub Button2_Click(sender As Object, e As System.EventArgs) Handles Button2.Click
        Me.TextBox2.Focus()
        CERRAR = True
        Me.Close()
    End Sub
    Private Sub INICIAR_OBJETOS()
        Me.TextBox1.Text = vMAc_ID_MA
        Me.TextBox2.Text = vMAc_APELLIDOS
        Me.TextBox3.Text = vMAc_NOMBRES
        Call CARGAR_COMBO_CAT_DE_ESPECIALIDADES_MEDICAS()
        Me.ComboBox1.Text = vMAc_D_E_MED
        Me.TextBox4.Text = vMAc_N_CEDULA
        Me.TextBox5.Text = vMAc_N_MINSA
        Me.TextBox6.Text = vMAc_TELEFONO_1
        Me.TextBox7.Text = vMAc_TELEFONO_2
        Me.TextBox8.Text = vMAc_CORREO_1
        Me.TextBox9.Text = vMAc_CORREO_2
        Call CARGAR_COMBO_CAT_DE_NACIONALIDAD_D()
        Me.ComboBox2.Text = vMAc_D_NACIONALIDAD_D
        Me.TextBox10.Text = vMAc_DIRECCION_D
        Call CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        Me.ComboBox3.Text = vMAc_D_T_SEXO
        If vMAc_FEC_NAC <> "" Then : Me.DateTimePicker1.Value = vMAc_FEC_NAC : Me.DateTimePicker1.Checked = True : Else : Me.DateTimePicker1.Checked = False : End If
        If vMAc_FEC_DESDE <> "" Then : Me.DateTimePicker2.Value = vMAc_FEC_DESDE : Me.DateTimePicker2.Checked = True : Else : Me.DateTimePicker2.Checked = False : End If
        If vMAc_FEC_HASTA <> "" Then : Me.DateTimePicker3.Value = vMAc_FEC_HASTA : Me.DateTimePicker3.Checked = True : Else : Me.DateTimePicker3.Checked = False : End If
        Me.TextBox11.Text = vMAc_FACEBOOK
        Me.TextBox12.Text = vMAc_TWITTER
        Me.TextBox13.Text = vMAc_WHATSAPP
        Me.TextBox14.Text = vMAc_OBSERVACIONES
        Call CARGAR_FOTO()
        Me.TextBox2.Focus()
    End Sub
    Private Sub CARGAR_FOTO()
        Dim DT As New DataTable
        Dim DA As New SqlDataAdapter
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()
        Try
            CADENAsql = "SELECT * FROM [VISTA MAESTRO DE MEDICOS ACREDITADOS] WHERE ID_MA = " & vMAc_ID_MA & " ORDER BY ID_MA"
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
    Private Sub CARGAR_COMBO_CAT_DE_TIPO_DE_SEXO()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE TIPO DE SEXO] ORDER BY ID_T_SEXO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox3
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
    Private Sub CARGAR_COMBO_CAT_DE_NACIONALIDAD_D()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE NACIONALIDAD_D] ORDER BY ID_NACIONALIDAD_D", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox2
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_NACIONALIDAD_D"
            End With
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub CARGAR_COMBO_CAT_DE_ESPECIALIDADES_MEDICAS()
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiSqlDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiSqlDataAdapter = New SqlDataAdapter("SELECT * FROM [CAT DE ESPECIALIDADES MEDICAS] ORDER BY ID_E_MED", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiSqlDataAdapter.Fill(MiDataTable)
            With Me.ComboBox1
                .DataSource = MiDataTable
                .DisplayMember = "DESCRIPCION"
                .ValueMember = "ID_E_MED"
            End With
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Private Sub TextBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox2.KeyDown
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
    Private Sub ComboBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox1.KeyDown
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
    Private Sub TextBox5_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox5.KeyDown
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
    Private Sub ComboBox2_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox2.KeyDown
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
    Private Sub ComboBox3_KeyDown(sender As Object, e As KeyEventArgs) Handles ComboBox3.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
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
    Private Sub DateTimePicker3_KeyDown(sender As Object, e As KeyEventArgs) Handles DateTimePicker3.KeyDown
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
    Private Sub TextBox12_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox12.KeyDown
        If e.KeyCode = Keys.Enter Then
            e.Handled = True
            SendKeys.Send("{TAB}")
        End If
    End Sub
    Private Sub TextBox13_KeyDown(sender As Object, e As KeyEventArgs) Handles TextBox13.KeyDown
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
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If TextBox1.Text = "0" Or TextBox1.Text = "" Then
            MsgBox("No se ha generado el Codigo del registro", vbInformation, "Mensaje del Sistema")
            Me.TextBox1.Text = "0"
            Me.Button2.Focus()
            Exit Sub
        End If
        If TextBox2.Text = "" Then
            MsgBox("Debe Digitar los Apellidos Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox2.Focus()
            Exit Sub
        End If
        If TextBox3.Text = "" Then
            MsgBox("Debe Digitar los Nombres Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox3.Focus()
            Exit Sub
        End If
        If Len(Me.ComboBox1.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Especialidad Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox1.Focus()
            Exit Sub
        End If
        If TextBox4.Text = "" Then
            MsgBox("Debe Digitar el Número de Cédula Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox4.Focus()
            Exit Sub
        End If
        If TextBox5.Text = "" Then
            MsgBox("Debe Digitar el Número MINSA Correctamente", vbInformation, "Mensaje del Sistema")
            Me.TextBox5.Focus()
            Exit Sub
        End If
        If TextBox6.Text = "" Then
            TextBox6.Text = ""
        End If
        If TextBox7.Text = "" Then
            TextBox7.Text = ""
        End If
        If Len(Me.ComboBox2.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar la Nacionalidad Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox2.Focus()
            Exit Sub
        End If
        If TextBox8.Text = "" Then
            TextBox8.Text = ""
        End If
        If TextBox9.Text = "" Then
            TextBox9.Text = ""
        End If
        If Len(Me.ComboBox3.SelectedValue) = 0 Then
            MsgBox("Debe seleccionar el Tipo de Sexo Correctamente", vbInformation, "Mensaje del Sistema")
            Me.ComboBox3.Focus()
            Exit Sub
        End If
        If TextBox10.Text = "" Then
            TextBox10.Text = ""
        End If
        If TextBox11.Text = "" Then
            TextBox11.Text = ""
        End If
        If TextBox12.Text = "" Then
            TextBox12.Text = ""
        End If
        If TextBox13.Text = "" Then
            TextBox13.Text = ""
        End If
        If TextBox14.Text = "" Then
            TextBox14.Text = ""
        End If
        Call Me.EDITAR_REGISTRO()
        CERRAR = False
        vMAc_ID_MA = Me.TextBox1.Text
        Me.TextBox1.Text = "0"
        Me.TextBox2.Text = ""
        Me.TextBox3.Text = ""
        Me.TextBox4.Text = ""
        Me.TextBox5.Text = ""
        Me.TextBox6.Text = ""
        Me.TextBox7.Text = ""
        Me.TextBox8.Text = ""
        Me.TextBox9.Text = ""
        Me.TextBox10.Text = ""
        Me.TextBox11.Text = ""
        Me.TextBox12.Text = ""
        Me.TextBox13.Text = ""
        Me.TextBox14.Text = ""
        Me.TextBox2.Focus()
        Me.Close()
    End Sub
    Dim fFEC_1 As String
    Dim fFEC_2 As String
    Dim fFEC_3 As String
    Private Sub EDITAR_REGISTRO()
        If Me.DateTimePicker1.Checked = True Then
            fFEC_1 = Mid(Me.DateTimePicker1.Value, 1, 10)
        Else
            fFEC_1 = "NULL"
        End If
        Dim F1 As Object = If(fFEC_1 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_1 + "', 105), 23)", "NULL")

        If Me.DateTimePicker2.Checked = True Then
            fFEC_2 = Mid(Me.DateTimePicker2.Value, 1, 10)
        Else
            fFEC_2 = "NULL"
        End If
        Dim F2 As Object = If(fFEC_2 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_2 + "', 105), 23)", "NULL")

        If Me.DateTimePicker3.Checked = True Then
            fFEC_3 = Mid(Me.DateTimePicker3.Value, 1, 10)
        Else
            fFEC_3 = "NULL"
        End If
        Dim F3 As Object = If(fFEC_3 <> "NULL", "Convert(VARCHAR(10), Convert(Date, '" + fFEC_3 + "', 105), 23)", "NULL")


        Dim gAPELLIDOS As Object = If(Me.TextBox2.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox2.Text) + "', 23)", "NULL")
        Dim gNOMBRES As Object = If(Me.TextBox3.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox3.Text) + "', 23)", "NULL")
        Dim gN_CEDULA As Object = If(Me.TextBox4.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox4.Text) + "', 23)", "NULL")
        Dim gN_MINSA As Object = If(Me.TextBox5.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox5.Text) + "', 23)", "NULL")
        Dim gTELEFONO_1 As Object = If(Me.TextBox6.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox6.Text) + "', 23)", "NULL")
        Dim gTELEFONO_2 As Object = If(Me.TextBox7.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox7.Text) + "', 23)", "NULL")
        Dim gCORREO_1 As Object = If(Me.TextBox8.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox8.Text) + "', 23)", "NULL")
        Dim gCORREO_2 As Object = If(Me.TextBox9.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox9.Text) + "', 23)", "NULL")
        Dim gDIRECCION_D As Object = If(Me.TextBox10.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox10.Text) + "', 23)", "NULL")
        Dim gFACEBOOK As Object = If(Me.TextBox11.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox11.Text) + "', 23)", "NULL")
        Dim gTWITTER As Object = If(Me.TextBox12.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox12.Text) + "', 23)", "NULL")
        Dim gWHATSAPP As Object = If(Me.TextBox13.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox13.Text) + "', 23)", "NULL")
        Dim gOBSERVACIONES As Object = If(Me.TextBox14.Text <> "", "Convert(NVARCHAR(200), '" + Trim(Me.TextBox14.Text) + "', 23)", "NULL")


        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Conexion.ABD_RRHH()

        Dim SQL As String
        Dim ms As New System.IO.MemoryStream() 'creando el memoryStream
        Try
            If CONTROL_FOTO = True Then
                Me.PictureBox4.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            Else
                Me.PictureBox7.Image.Save(ms, System.Drawing.Imaging.ImageFormat.Jpeg)
            End If
            SQL = "UPDATE [MAESTRO DE MEDICOS ACREDITADOS] SET NOMBRES = " & gNOMBRES & ", APELLIDOS = " & gAPELLIDOS & ", ID_E_MED  = " & Me.ComboBox1.SelectedValue & ", N_CEDULA = " & gN_CEDULA & ", N_MINSA = " & gN_MINSA & ", TELEFONO_1 = " & gTELEFONO_1 & ", TELEFONO_2 = " & gTELEFONO_2 & ", CORREO_1 = " & gCORREO_1 & ", CORREO_2 = " & gCORREO_2 & ", ID_NACIONALIDAD_D = " & Me.ComboBox2.SelectedValue & ", DIRECCION_D = " & gDIRECCION_D & ", ID_T_SEXO = " & Me.ComboBox3.SelectedValue & ", FEC_NAC = " & F1 & ", FEC_DESDE = " & F2 & ", FEC_HASTA = " & F3 & ", FACEBOOK = " & gFACEBOOK & ", TWITTER = " & gTWITTER & ", WHATSAPP = " & gWHATSAPP & ", OBSERVACIONES = " & gOBSERVACIONES & ", FOTO = @data WHERE ID_MA = " & vMAc_ID_MA & ""
            Dim cmd As New SqlCommand(SQL, CxRRHH)
            cmd.Parameters.Add(New SqlParameter("@data", ms.GetBuffer()))
            cmd.ExecuteNonQuery()
            cmd = Nothing
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
                ReDim Data(ios.Length) 'llenamos el arreglo
                ios.Read(Data, 0, CInt(ios.Length)) 'llenamos el arreglo
                Me.PictureBox4.SizeMode = PictureBoxSizeMode.StretchImage 'establecemos como se visualiza la imagen
                Me.PictureBox4.Load(Me.OpenFileDialog1.FileName) 'visualizamos abriendo el archivo seleccionado
                CONTROL_FOTO = True
            End If
        Catch ex As Exception
            MessageBox.Show(ex.Message, "Mensaje del Sistema", MessageBoxButtons.OK, MessageBoxIcon.Exclamation, MessageBoxDefaultButton.Button1, MessageBoxOptions.ServiceNotification) 'en caso de error muestre un mensaje
        End Try
    End Sub
End Class