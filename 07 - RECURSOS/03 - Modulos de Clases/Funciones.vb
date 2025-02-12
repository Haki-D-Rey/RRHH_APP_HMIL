Option Explicit On
Imports System
Imports System.Data
Imports System.Data.SqlClient
Imports System.Windows.Forms
Imports Microsoft.Office.Interop
Module Funciones
    Public Sub VERIFICAR_ACCESOS(ByVal CLAVE_ACCESO As String)
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS USUARIOS] WHERE ID_USUARIO = " & CInt(isuID_USUARIO) & " AND CLAVE = '" & CLAVE_ACCESO & "' AND ACCESO = 1 ORDER BY ID_USUARIO, ID_NODO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                HAY_ACCESO = True
            Else
                HAY_ACCESO = False
                'MsgBox("El Usuario Actual no tiene Acceso a esta Opci�n [" & MiDataTable.Rows(0).Item(4).ToString & "] " & MiDataTable.Rows(0).Item(1).ToString, vbInformation, "Mensaje del Sistema")
                MsgBox("El Usuario Actual no tiene Acceso a esta Opci�n", vbInformation, "Mensaje del Sistema")
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    Public Sub VERIFICAR_ACCESOS_SIN_MENSAJE(ByVal CLAVE_ACCESO As String)
        If Conexion.CxRRHH.State = ConnectionState.Open Then
            Conexion.CBD_RRHH()
        End If
        Dim MiDataTable As New DataTable
        Dim MiDataAdapter As SqlDataAdapter
        Try
            Conexion.ABD_RRHH()
            MiDataAdapter = New SqlDataAdapter("SELECT * FROM [MAESTRO DE NODOS USUARIOS] WHERE ID_USUARIO = " & CInt(isuID_USUARIO) & " AND CLAVE = '" & CLAVE_ACCESO & "' AND ACCESO = 1 ORDER BY ID_USUARIO, ID_NODO", Conexion.CxRRHH)
            MiDataTable.Clear()
            MiDataAdapter.Fill(MiDataTable)
            If MiDataTable.Rows.Count > 0 Then
                HAY_ACCESO = True
            Else
                HAY_ACCESO = False
            End If
        Catch ex As Exception
            MsgBox("Error: " & ex.Message, vbInformation, "Mensaje del Sistema")
        Finally
            If Conexion.CxRRHH.State = ConnectionState.Open Then
                Conexion.CBD_RRHH()
            End If
        End Try
    End Sub
    '******************************************************************************************************
    '******************************************************************************************************
    Public Function LETRAS_CONSTANCIAS(ByVal NCIFRA As Object) As String
        Dim cifra, bloque, decimales, cadena As String
        Dim posision, unidadmil As Byte
        'Dim longituid As Byte 
        cifra = Format(CType(NCIFRA, Decimal), "###############0.#0")
        decimales = Mid(cifra, Len(cifra) - 1, 2)
        cifra = Left(cifra, Len(cifra) - 3)
        'Verifico que el valor no sea cero
        If cifra = "0" Then
            Return IIf(decimales = "00", "cero", "cero con " & decimales & "/100")
        End If
        'Evaluo su longitud (como m�nimo una cadena debe tener 3 d�gitos)
        If Len(cifra) < 3 Then
            cifra = RELLENAR(cifra, 3)
        End If
        'Invierto la cadena
        cifra = INVERTIR(cifra)
        'Inicializo variables
        posision = 1
        unidadmil = 0
        cadena = ""
        'Selecciono bloques de a tres cifras empezando desde el final (de la cadena invertida)
        Do While posision <= Len(cifra)
            ' Selecciono una porci�n del numero
            bloque = Mid(cifra, posision, 3)
            ' Transformo el n�mero a cadena
            cadena = CONVERTIR_CONSTANCIAS(bloque, unidadmil) & " " & cadena.Trim
            ' Incremento la cantidad desde donde seleccionar la subcadena
            posision = posision + 3
            ' Incremento la posisi�n de la unidad de mil
            unidadmil = unidadmil + 1
        Loop
        ' Cargo la funci�n
        Return IIf(decimales = "00", cadena.Trim.ToLower & " cordobas netos", cadena.Trim.ToLower & " cordobas con " & decimales & "/100")
    End Function
    ' Esta funci�n es complemento de la funci�n de conversi�n.
    ' En los arrays se agrega una posisi�n inicial vac�a ya que VB.NET empieza de la posisi�n cero
    Private Function CONVERTIR_CONSTANCIAS(ByVal cadena As String, ByVal unidadmil As Byte) As String
        'Defino variables
        Dim centena, decena, unidad As Byte
        ' Invierto la subcadena (la original habia sido invertida en el procedimiento NumeroATexto)
        cadena = INVERTIR(cadena)
        ' Determino la longitud de la cadena
        If Len(cadena) < 3 Then
            cadena = RELLENAR(cadena, 3)
        End If
        ' Verifico que la cadena no est� vac�a (000)
        If cadena = "000" Then
            Return ""
        End If
        ' Desarmo el numero (empiezo del d�gito cero por el manejo de cadenas de VB.NET)
        centena = CType(cadena.Substring(0, 1), Byte)
        decena = CType(cadena.Substring(1, 1), Byte)
        unidad = CType(cadena.Substring(2, 1), Byte)
        cadena = ""
        ' Calculo las centenas
        If centena <> 0 Then
            Dim centenas() As String = {"", IIf(decena = 0 And unidad = 0, "cien", "ciento"), "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"}
            cadena = centenas(centena)
        End If
        ' Calculo las decenas
        If decena <> 0 Then
            Dim decenas() As String = {"", IIf(unidad = 0, "diez", IIf(unidad >= 6, "dieci", IIf(unidad = 1, "once", IIf(unidad = 2, "doce", IIf(unidad = 3, "trece", IIf(unidad = 4, "catorce", "quince")))))), IIf(unidad = 0, "veinte", "veinti"), "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"}
            cadena = cadena & " " & decenas(decena)
        End If
        ' Calculo las unidades (no pregunten por que este IF es necesario ... simplemente funciona)
        If decena = 1 And unidad < 6 Then
        Else
            Dim unidades() As String = {"", IIf(decena <> 1, IIf(unidadmil = 1, "un", "uno"), ""), "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"}
            If decena >= 3 And unidad <> 0 Then
                cadena = cadena.Trim & " y "
            End If
            If decena = 0 Then
                cadena = cadena.Trim & " "
            End If
            cadena = cadena & unidades(unidad)
        End If
        ' Evaluo la posision de miles, millones, etc
        If unidadmil <> 0 Then
            Dim agregado() As String = {"", "mil", IIf((centena = 0) And (decena = 0) And (unidad = 1), "mill�n", "millones"), "mil millones", "billones", "mil billones"}
            If (centena = 0) And (decena = 0) And (unidad = 1) And unidadmil = 2 Then
                cadena = "un"
            End If
            cadena = cadena & " " & agregado(unidadmil)
        End If
        ' Cargo la funci�n
        Return cadena.Trim
    End Function
    ' Esta funci�n recibe una cadena de caracteres y la devuelve "invertida".
    'Dim retornar As String
    Public Function INVERTIR(ByVal cadena As String) As String
        ' Defino variables
        Dim retornar As String
        ' Inviero la cadena
        For posision As Short = cadena.Length To 1 Step -1
            retornar = retornar & cadena.Substring(posision - 1, 1)
        Next
        ' Retorno la cadena invertida
        Return retornar
    End Function
    ' Esta funci�n rellena con ceros a la izquierda un n�mero pasado como par�metro. Con el par�metro "cifras" se especifica la cantidad de d�gitos a la izquierda.
    Public Function RELLENAR(ByVal valor As Object, ByVal cifras As Byte) As String
        ' Defino variables
        Dim cadena As String
        ' Verifico el valor pasado
        If Not IsNumeric(valor) Then
            valor = 0
        Else
            valor = CType(valor, Integer)
        End If
        ' Cargo la cadena
        cadena = valor.ToString.Trim
        ' Relleno con los ceros que sean necesarios para llenar los d�gitos pedidos
        For puntero As Byte = (Len(cadena) + 1) To cifras
            cadena = "0" & cadena
        Next puntero
        ' Cargo la funci�n
        Return cadena
    End Function
    Public TITULO_EXCEL As String
    Public Sub EXPORTAR_DATOS_A_EXCEL(ByVal DataGridView1 As DataGridView, ByVal titulo As String)
        Dim m_Excel As New Excel.Application
        m_Excel.Cursor = Excel.XlMousePointer.xlWait
        m_Excel.Visible = True
        Dim objLibroExcel As Excel.Workbook = m_Excel.Workbooks.Add
        Dim objHojaExcel As Excel.Worksheet = objLibroExcel.Worksheets(1)
        With objHojaExcel
            .Visible = Excel.XlSheetVisibility.xlSheetVisible
            .Activate()
            'Encabezado  
            .Range("A1:L1").Merge()
            .Range("A1:L1").Value = TITULO_EXCEL
            .Range("A1:L1").Font.Bold = True
            .Range("A1:L1").Font.Size = 15
            'Copete  
            .Range("A2:L2").Merge()
            .Range("A2:L2").Value = titulo
            .Range("A2:L2").Font.Bold = True
            .Range("A2:L2").Font.Size = 12

            Const primeraLetra As Char = "A"
            Const primerNumero As Short = 3
            Dim Letra As Char, UltimaLetra As Char
            Dim Numero As Integer, UltimoNumero As Integer
            Dim cod_letra As Byte = Asc(primeraLetra) - 1
            Dim sepDec As String = Application.CurrentCulture.NumberFormat.NumberDecimalSeparator
            Dim sepMil As String = Application.CurrentCulture.NumberFormat.NumberGroupSeparator
            'Establecer formatos de las columnas de la hija de c�lculo  
            Dim strColumna As String = ""
            Dim LetraIzq As String = ""
            Dim cod_LetraIzq As Byte = Asc(primeraLetra) - 1
            Letra = primeraLetra
            Numero = primerNumero
            Dim objCelda As Excel.Range
            For Each c As DataGridViewColumn In DataGridView1.Columns
                If c.Visible Then
                    If Letra = "Z" Then
                        Letra = primeraLetra
                        cod_letra = Asc(primeraLetra)
                        cod_LetraIzq += 1
                        LetraIzq = Chr(cod_LetraIzq)
                    Else
                        cod_letra += 1
                        Letra = Chr(cod_letra)
                    End If
                    strColumna = LetraIzq + Letra + Numero.ToString
                    objCelda = .Range(strColumna, Type.Missing)
                    objCelda.Value = c.HeaderText
                    objCelda.EntireColumn.Font.Size = 8
                    'objCelda.EntireColumn.NumberFormat = c.DefaultCellStyle.Format  
                    If c.ValueType Is GetType(Decimal) OrElse c.ValueType Is GetType(Double) Then
                        objCelda.EntireColumn.NumberFormat = "#" + sepMil + "0" + sepDec + "00"
                    End If
                End If
            Next

            Dim objRangoEncab As Excel.Range = .Range(primeraLetra + Numero.ToString, LetraIzq + Letra + Numero.ToString)
            objRangoEncab.BorderAround(1, Excel.XlBorderWeight.xlMedium)
            UltimaLetra = Letra
            Dim UltimaLetraIzq As String = LetraIzq

            'CARGA DE DATOS  
            Dim i As Integer = Numero + 1

            For Each reg As DataGridViewRow In DataGridView1.Rows
                LetraIzq = ""
                cod_LetraIzq = Asc(primeraLetra) - 1
                Letra = primeraLetra
                cod_letra = Asc(primeraLetra) - 1
                For Each c As DataGridViewColumn In DataGridView1.Columns
                    If c.Visible Then
                        If Letra = "Z" Then
                            Letra = primeraLetra
                            cod_letra = Asc(primeraLetra)
                            cod_LetraIzq += 1
                            LetraIzq = Chr(cod_LetraIzq)
                        Else
                            cod_letra += 1
                            Letra = Chr(cod_letra)
                        End If
                        strColumna = LetraIzq + Letra
                        ' ac� deber�a realizarse la carga  
                        .Cells(i, strColumna) = IIf(IsDBNull(reg.ToString), "", reg.Cells(c.Index).Value)
                        '.Cells(i, strColumna) = IIf(IsDBNull(reg.(c.DataPropertyName)), c.DefaultCellStyle.NullValue, reg(c.DataPropertyName))  
                        '.Range(strColumna + i, strColumna + i).In()  

                    End If
                Next
                Dim objRangoReg As Excel.Range = .Range(primeraLetra + i.ToString, strColumna + i.ToString)
                objRangoReg.Rows.BorderAround()
                objRangoReg.Select()
                i += 1
            Next
            UltimoNumero = i

            'Dibujar las l�neas de las columnas  
            LetraIzq = ""
            cod_LetraIzq = Asc("A")
            cod_letra = Asc(primeraLetra)
            Letra = primeraLetra
            For Each c As DataGridViewColumn In DataGridView1.Columns
                If c.Visible Then
                    objCelda = .Range(LetraIzq + Letra + primerNumero.ToString, LetraIzq + Letra + (UltimoNumero - 1).ToString)
                    objCelda.BorderAround()
                    If Letra = "Z" Then
                        Letra = primeraLetra
                        cod_letra = Asc(primeraLetra)
                        LetraIzq = Chr(cod_LetraIzq)
                        cod_LetraIzq += 1
                    Else
                        cod_letra += 1
                        Letra = Chr(cod_letra)
                    End If
                End If
            Next

            'Dibujar el border exterior grueso  
            Dim objRango As Excel.Range = .Range(primeraLetra + primerNumero.ToString, UltimaLetraIzq + UltimaLetra + (UltimoNumero - 1).ToString)
            objRango.Select()
            objRango.Columns.AutoFit()
            objRango.Columns.BorderAround(1, Excel.XlBorderWeight.xlMedium)
        End With

        m_Excel.Cursor = Excel.XlMousePointer.xlDefault
    End Sub
    '******************************************************************************************************
    '******************************************************************************************************
    '******************************************************************************************************
    '******************************************************************************************************
    Public Function LETRAS(ByVal NCIFRA As Object) As String
        Dim cifra, bloque, decimales, cadena As String
        Dim posision, unidadmil As Byte
        'Dim longituid As Byte 
        cifra = Format(CType(NCIFRA, Decimal), "###############0.#0")
        decimales = Mid(cifra, Len(cifra) - 1, 2)
        cifra = Left(cifra, Len(cifra) - 3)
        'Verifico que el valor no sea cero
        If cifra = "0" Then
            Return IIf(decimales = "00", "cero", "cero con " & decimales & "/100")
        End If
        'Evaluo su longitud (como m�nimo una cadena debe tener 3 d�gitos)
        If Len(cifra) < 3 Then
            cifra = RELLENAR(cifra, 3)
        End If
        'Invierto la cadena
        cifra = INVERTIR(cifra)
        'Inicializo variables
        posision = 1
        unidadmil = 0
        cadena = ""
        'Selecciono bloques de a tres cifras empezando desde el final (de la cadena invertida)
        Do While posision <= Len(cifra)
            ' Selecciono una porci�n del numero
            bloque = Mid(cifra, posision, 3)
            ' Transformo el n�mero a cadena
            cadena = CONVERTIR(bloque, unidadmil) & " " & cadena.Trim
            ' Incremento la cantidad desde donde seleccionar la subcadena
            posision = posision + 3
            ' Incremento la posisi�n de la unidad de mil
            unidadmil = unidadmil + 1
        Loop
        ' Cargo la funci�n
        Return IIf(decimales = "00", cadena.Trim.ToLower, cadena.Trim.ToLower & " con " & decimales & "/100")
    End Function
    ' Esta funci�n es complemento de la funci�n de conversi�n.
    ' En los arrays se agrega una posisi�n inicial vac�a ya que VB.NET empieza de la posisi�n cero
    Private Function CONVERTIR(ByVal cadena As String, ByVal unidadmil As Byte) As String
        On Error GoTo MENSAJE
        'Defino variables
        Dim centena, decena, unidad As Byte
        ' Invierto la subcadena (la original habia sido invertida en el procedimiento NumeroATexto)
        cadena = INVERTIR(cadena)
        ' Determino la longitud de la cadena
        If Len(cadena) < 3 Then
            cadena = RELLENAR(cadena, 3)
        End If
        ' Verifico que la cadena no est� vac�a (000)
        If cadena = "000" Then
            Return ""
        End If
        ' Desarmo el numero (empiezo del d�gito cero por el manejo de cadenas de VB.NET)
        centena = CType(cadena.Substring(0, 1), Byte)
        decena = CType(cadena.Substring(1, 1), Byte)
        unidad = CType(cadena.Substring(2, 1), Byte)
        cadena = ""
        ' Calculo las centenas
        If centena <> 0 Then
            Dim centenas() As String = {"", IIf(decena = 0 And unidad = 0, "cien", "ciento"), "doscientos", "trescientos", "cuatrocientos", "quinientos", "seiscientos", "setecientos", "ochocientos", "novecientos"}
            cadena = centenas(centena)
        End If
        ' Calculo las decenas
        If decena <> 0 Then
            Dim decenas() As String = {"", IIf(unidad = 0, "diez", IIf(unidad >= 6, "dieci", IIf(unidad = 1, "once", IIf(unidad = 2, "doce", IIf(unidad = 3, "trece", IIf(unidad = 4, "catorce", "quince")))))), IIf(unidad = 0, "veinte", "veinti"), "treinta", "cuarenta", "cincuenta", "sesenta", "setenta", "ochenta", "noventa"}
            cadena = cadena & " " & decenas(decena)
        End If
        ' Calculo las unidades (no pregunten por que este IF es necesario ... simplemente funciona)
        If decena = 1 And unidad < 6 Then
        Else
            Dim unidades() As String = {"", IIf(decena <> 1, IIf(unidadmil = 1, "un", "uno"), ""), "dos", "tres", "cuatro", "cinco", "seis", "siete", "ocho", "nueve"}
            If decena >= 3 And unidad <> 0 Then
                cadena = cadena.Trim & " y "
            End If
            If decena = 0 Then
                cadena = cadena.Trim & " "
            End If
            cadena = cadena & unidades(unidad)
        End If
        ' Evaluo la posision de miles, millones, etc
        If unidadmil <> 0 Then
            Dim agregado() As String = {"", "mil", IIf((centena = 0) And (decena = 0) And (unidad = 1), "mill�n", "millones"), "mil millones", "billones", "mil billones"}
            If (centena = 0) And (decena = 0) And (unidad = 1) And unidadmil = 2 Then
                cadena = "un"
            End If
            cadena = cadena & " " & agregado(unidadmil)
        End If
        ' Cargo la funci�n
        Return cadena.Trim
        Exit Function
MENSAJE:
        MsgBox("Error: Posible Valor Negativo", vbInformation, "Mensaje del Sistema")
    End Function
    Public Sub EXPORTAR_DATOS_TICKETS(ByVal DataGridView1 As DataGridView)
        Dim m_Excel As New Excel.Application
        m_Excel.Cursor = Excel.XlMousePointer.xlWait
        m_Excel.Visible = True
        Dim objLibroExcel As Excel.Workbook = m_Excel.Workbooks.Add
        Dim objHojaExcel As Excel.Worksheet = objLibroExcel.Worksheets(1)
        With objHojaExcel
            .Visible = Excel.XlSheetVisibility.xlSheetVisible
            .Activate()
            ''Encabezado  
            '.Range("A1:L1").Merge()
            '.Range("A1:L1").Value = TITULO_EXCEL
            '.Range("A1:L1").Font.Bold = True
            '.Range("A1:L1").Font.Size = 15
            ''Copete  
            '.Range("A2:L2").Merge()
            '.Range("A2:L2").Value = titulo
            '.Range("A2:L2").Font.Bold = True
            '.Range("A2:L2").Font.Size = 12

            Const primeraLetra As Char = "A"
            Const primerNumero As Short = 1
            Dim Letra As Char, UltimaLetra As Char
            Dim Numero As Integer, UltimoNumero As Integer
            Dim cod_letra As Byte = Asc(primeraLetra) - 1
            Dim sepDec As String = Application.CurrentCulture.NumberFormat.NumberDecimalSeparator
            Dim sepMil As String = Application.CurrentCulture.NumberFormat.NumberGroupSeparator
            'Establecer formatos de las columnas de la hija de c�lculo  
            Dim strColumna As String = ""
            Dim LetraIzq As String = ""
            Dim cod_LetraIzq As Byte = Asc(primeraLetra) - 1
            Letra = primeraLetra
            Numero = primerNumero
            Dim objCelda As Excel.Range
            For Each c As DataGridViewColumn In DataGridView1.Columns
                If c.Visible Then
                    If Letra = "Z" Then
                        Letra = primeraLetra
                        cod_letra = Asc(primeraLetra)
                        cod_LetraIzq += 1
                        LetraIzq = Chr(cod_LetraIzq)
                    Else
                        cod_letra += 1
                        Letra = Chr(cod_letra)
                    End If
                    strColumna = LetraIzq + Letra + Numero.ToString
                    objCelda = .Range(strColumna, Type.Missing)
                    objCelda.Value = c.HeaderText
                    objCelda.EntireColumn.Font.Size = 8
                    'objCelda.EntireColumn.NumberFormat = c.DefaultCellStyle.Format  
                    If c.ValueType Is GetType(Decimal) OrElse c.ValueType Is GetType(Double) Then
                        objCelda.EntireColumn.NumberFormat = "#" + sepMil + "0" + sepDec + "00"
                    End If
                End If
            Next

            Dim objRangoEncab As Excel.Range = .Range(primeraLetra + Numero.ToString, LetraIzq + Letra + Numero.ToString)
            objRangoEncab.BorderAround(1, Excel.XlBorderWeight.xlMedium)
            UltimaLetra = Letra
            Dim UltimaLetraIzq As String = LetraIzq

            'CARGA DE DATOS  
            Dim i As Integer = Numero + 1

            For Each reg As DataGridViewRow In DataGridView1.Rows
                LetraIzq = ""
                cod_LetraIzq = Asc(primeraLetra) - 1
                Letra = primeraLetra
                cod_letra = Asc(primeraLetra) - 1
                For Each c As DataGridViewColumn In DataGridView1.Columns
                    If c.Visible Then
                        If Letra = "Z" Then
                            Letra = primeraLetra
                            cod_letra = Asc(primeraLetra)
                            cod_LetraIzq += 1
                            LetraIzq = Chr(cod_LetraIzq)
                        Else
                            cod_letra += 1
                            Letra = Chr(cod_letra)
                        End If
                        strColumna = LetraIzq + Letra
                        ' ac� deber�a realizarse la carga  
                        .Cells(i, strColumna) = IIf(IsDBNull(reg.ToString), "", reg.Cells(c.Index).Value)
                        '.Cells(i, strColumna) = IIf(IsDBNull(reg.(c.DataPropertyName)), c.DefaultCellStyle.NullValue, reg(c.DataPropertyName))  
                        '.Range(strColumna + i, strColumna + i).In()  

                    End If
                Next
                Dim objRangoReg As Excel.Range = .Range(primeraLetra + i.ToString, strColumna + i.ToString)
                objRangoReg.Rows.BorderAround()
                objRangoReg.Select()
                i += 1
            Next
            UltimoNumero = i

            'Dibujar las l�neas de las columnas  
            LetraIzq = ""
            cod_LetraIzq = Asc("A")
            cod_letra = Asc(primeraLetra)
            Letra = primeraLetra
            For Each c As DataGridViewColumn In DataGridView1.Columns
                If c.Visible Then
                    objCelda = .Range(LetraIzq + Letra + primerNumero.ToString, LetraIzq + Letra + (UltimoNumero - 1).ToString)
                    objCelda.BorderAround()
                    If Letra = "Z" Then
                        Letra = primeraLetra
                        cod_letra = Asc(primeraLetra)
                        LetraIzq = Chr(cod_LetraIzq)
                        cod_LetraIzq += 1
                    Else
                        cod_letra += 1
                        Letra = Chr(cod_letra)
                    End If
                End If
            Next

            'Dibujar el border exterior grueso  
            Dim objRango As Excel.Range = .Range(primeraLetra + primerNumero.ToString, UltimaLetraIzq + UltimaLetra + (UltimoNumero - 1).ToString)
            objRango.Select()
            objRango.Columns.AutoFit()
            objRango.Columns.BorderAround(1, Excel.XlBorderWeight.xlMedium)
        End With

        m_Excel.Cursor = Excel.XlMousePointer.xlDefault
    End Sub
End Module