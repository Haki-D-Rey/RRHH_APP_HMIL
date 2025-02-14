Imports Microsoft.Office.Interop.Excel

Public Class ExportadorExcelInterop

    Public Shared Function ExportarAExcelOptimizado(dataGridView As DataGridView, rutaArchivo As String) As Boolean
        Try
            ' Verificar que el DataGridView tenga datos
            If dataGridView.Rows.Count = 0 Then
                MessageBox.Show("No hay datos para exportar.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return False
            End If

            ' Iniciar Excel
            Dim excelApp As New Application
            Dim workbook As Workbook = excelApp.Workbooks.Add()
            Dim worksheet As Worksheet = workbook.Sheets(1)

            ' Desactivar actualizaciones para mejorar rendimiento
            excelApp.ScreenUpdating = False
            excelApp.Calculation = XlCalculation.xlCalculationManual
            excelApp.DisplayAlerts = False

            ' Obtener el número de filas y columnas
            Dim numFilas As Integer = dataGridView.Rows.Count
            Dim numColumnas As Integer = dataGridView.Columns.Count

            ' Crear array para almacenar los datos y escribirlos en bloque
            Dim datos(numFilas, numColumnas - 1) As Object

            ' Escribir encabezados en Excel
            Dim headers(numColumnas - 1) As String
            For col As Integer = 0 To numColumnas - 1
                headers(col) = dataGridView.Columns(col).HeaderText
            Next
            Dim headerRange As Range = worksheet.Range(worksheet.Cells(1, 1), worksheet.Cells(1, numColumnas))
            headerRange.Value = headers
            headerRange.Font.Bold = True
            headerRange.Interior.Color = RGB(200, 200, 200)
            headerRange.HorizontalAlignment = XlHAlign.xlHAlignCenter

            ' Llenar array con los datos del DataGridView
            For fila As Integer = 0 To numFilas - 1
                For col As Integer = 0 To numColumnas - 1
                    datos(fila, col) = dataGridView.Rows(fila).Cells(col).Value
                Next
            Next

            ' Escribir datos en Excel de una sola vez
            Dim dataRange As Range = worksheet.Range(worksheet.Cells(2, 1), worksheet.Cells(numFilas + 1, numColumnas))
            dataRange.Value = datos

            ' Ajustar columnas automáticamente
            worksheet.Columns.AutoFit()

            ' Guardar archivo
            workbook.SaveAs(rutaArchivo)
            workbook.Close()
            excelApp.Quit()

            ' Liberar memoria
            System.Runtime.InteropServices.Marshal.ReleaseComObject(worksheet)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)

            MessageBox.Show("Datos exportados a Excel exitosamente en: " & rutaArchivo, "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return True
        Catch ex As Exception
            MessageBox.Show("Error al exportar a Excel: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function


End Class
