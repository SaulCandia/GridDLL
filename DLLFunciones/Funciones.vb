'Componentes de edición de grillas en DLL.
'© 2022 Saúl Candia. - Holding Leonera.
'Programado para VisualBasic.NET

Imports System.Globalization
Imports System.IO
Imports System.Windows.Forms
Imports ExcelDataReader
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Funciones

    Class FORMATFECHA

        Shared Function FORMATEOFECHA(ByVal stringfecha As String)
            'CASO CERO: Si se importa un excel y se debe editar la fecha, la misma aparecerá con hora, por lo que
            'se limpia para usar solo la fecha
            If stringfecha.Contains(":") Then
                Dim arrayFecha() As String
                arrayFecha = stringfecha.Split("-")

                'Se agregan ceros a los días y meses para que del 1 al 10 no queden vacíos
                Dim dia As String = arrayFecha(0).PadLeft(2, "0"c)
                Dim mes As String = arrayFecha(1).PadLeft(2, "0"c)

                'En la importación del Excel al DataGridView, el mes viene en palabras por lo que pasa a convertirse en número
                Dim mesInt As String = (DateTime.ParseExact(mes, "MMM", CultureInfo.CurrentCulture).Month).ToString.PadLeft(2, "0"c)

                Dim anio As String = Date.Today.Year.ToString

                'Fecha va en formato aaaaMMdd para que sea fácilmente editable e insertable en la base de datos
                stringfecha = anio + mesInt + dia
                Return stringfecha

            End If
            'Caso 1: Si la fecha digitada contiene guión, se procede a formatear y separar.
            If stringfecha.Contains("-") Then

                Dim arrayFecha() As String
                arrayFecha = stringfecha.Split("-")

                'Se agregan ceros a los días y meses para que del 1 al 10 no queden vacíos
                Dim dia As String = arrayFecha(0).PadLeft(2, "0"c)
                Dim mes As String = arrayFecha(1).PadLeft(2, "0"c)
                Dim anio As String = Date.Today.Year.ToString

                'Fecha va en formato aaaaMMdd para que sea fácilmente editable e insertable en la base de datos
                stringfecha = anio + mes + dia

                Return stringfecha
            Else
                'Caso 2: Si la fecha digitada NO contiene guión, se procede a formatear como cero,
                'separar los dígitos y verificando si los días y meses contienen ceros que la precedan.
                Try

                    'Dim dateTime As String = stringfecha
                    'Dim dt As DateTime = Convert.ToDateTime(dateTime)
                    'Dim format As String = "yyyy-MM-dd"
                    'stringfecha = dt.ToString(format)
                    'Console.WriteLine(stringfecha)

                    Dim dia As String
                    Dim mes As String
                    Dim anio As String = Date.Today.Year.ToString

                    If (stringfecha.Length <= 3) Then
                        Return "ERROR: El formato de fecha no es válido." & vbCrLf & vbCrLf & "El formato puede ser" & vbCrLf & vbCrLf & "ddMM / ddMMaa / ddMMaaaa / dd-MM-aa / dd-MM-aaaa" & vbCrLf & vbCrLf & "Si el día o el mes son menores a 10, debe agregar un cero adelante."
                    End If

                    If (stringfecha.Length > 6) Then
                        dia = Left(stringfecha, 2)
                        mes = Right(Left(stringfecha, 4), 2)
                    ElseIf (stringfecha.Length <= 6) Then
                        dia = Left(stringfecha, 2)
                        mes = Right(Left(stringfecha, 4), 2)
                    Else
                        dia = Left(stringfecha, 2)
                        mes = Left(Right(stringfecha, 4), 2)
                    End If



                    'Fecha va en formato aaaaMMdd para que sea fácilmente editable e insertable en la base de datos
                    stringfecha = anio + mes + dia
                    Return stringfecha

                Catch ex As Exception
                    Return "ERROR: " + ex.Message.ToString()
                End Try
            End If


        End Function

    End Class

    Class EXPORTARAEXCEL
        Shared Function GUARDARAEXCEL(ByVal grilla As Object, ByVal ubicacionArchivo As String)
            Try
                Dim xlApp As Excel.Application
                Dim xlLibro As Excel.Workbook
                Dim xlHoja As Excel.Worksheet
                Dim misValue As Object = System.Reflection.Missing.Value
                Dim i As Integer
                Dim j As Integer

                xlApp = New Excel.Application
                xlLibro = xlApp.Workbooks.Add(misValue)
                xlHoja = CType(xlLibro.Sheets(1), Excel.Worksheet)

                For i = 0 To grilla.RowCount - 1
                    For j = 0 To grilla.ColumnCount - 1
                        xlHoja.Cells(i + 1, j + 1) = grilla(j, i).Value.ToString
                    Next
                Next

                xlHoja.SaveAs(ubicacionArchivo)
                xlLibro.Close()
                xlApp.Quit()

                Try
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp)
                    xlApp = Nothing
                Catch ex As Exception
                    xlApp = Nothing
                Finally
                    GC.Collect()
                End Try

                Try
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlLibro)
                    xlLibro = Nothing
                Catch ex As Exception
                    xlLibro = Nothing
                Finally
                    GC.Collect()
                End Try

                Try
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(xlHoja)
                    xlHoja = Nothing
                Catch ex As Exception
                    xlHoja = Nothing
                Finally
                    GC.Collect()
                End Try

                Return "EXITO"

            Catch ex As Exception
                Return "ERROR" + ex.Message
            End Try
        End Function
    End Class

    Class IMPORTAREXCEL
        Shared Function IMPORTARDESDEEXCEL(ByVal archivo As String)
            Try
                'Invoca al ExcelDataReader
                Dim reader As IExcelDataReader
                Dim stream = File.Open(archivo, FileMode.Open, FileAccess.Read)
                reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream)
                Dim conf = New ExcelDataSetConfiguration With {
                .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration With {
                    .UseHeaderRow = False
                }
            }
                Dim dataSet = reader.AsDataSet(conf)
                Dim dataTable = dataSet.Tables(0)

                Return dataTable

            Catch ex As Exception
                Return "ERROR: " + ex.Message.ToString()
            End Try
        End Function
    End Class


    Class CORTAR
        Shared Function CORTAR(ByVal grilla As Object)
            Try
                Dim resultado As String

                'Copia elementos al portapapeles
                Dim dataObj As DataObject = grilla.GetClipboardContent()
                If dataObj IsNot Nothing Then Clipboard.SetDataObject(dataObj)

                'Limpia la fila cortada
                For Each dgvCell As DataGridViewCell In grilla.SelectedCells
                    dgvCell.Value = String.Empty
                Next

                resultado = "Exito"

                Return resultado

            Catch ex As Exception
                Return "ERROR: " + ex.Message.ToString()
            End Try
        End Function
    End Class

    Class COPIAR
        Shared Function COPIAR(ByVal grilla As Object)
            Try
                Dim resultado As String

                'Copia elementos al portapapeles
                Dim dataObj As DataObject = grilla.GetClipboardContent()
                If dataObj IsNot Nothing Then Clipboard.SetDataObject(dataObj)

                resultado = "Exito"

                Return resultado

            Catch ex As Exception
                Return "ERROR: " + ex.Message.ToString()
            End Try
        End Function
    End Class


    Class PEGAR
        Shared Function PEGAR(ByVal grilla As DataGridView, datoPortapapeles As String)

            Try
                'Si el tamaño de la grilla es cero, no devuelve valores.
                If grilla.SelectedCells.Count = 0 Then Return Nothing

                'Paso 1: Chequea tamaño de DataGridView existente
                Dim rowIndex As Integer = grilla.Rows.Count - 1
                Dim colIndex As Integer = grilla.Columns.Count - 1

                For Each dgvCell As DataGridViewCell In grilla.SelectedCells
                    If dgvCell.RowIndex < rowIndex Then rowIndex = dgvCell.RowIndex
                    If dgvCell.ColumnIndex < colIndex Then colIndex = dgvCell.ColumnIndex
                Next

                'Determina la coordenadas de inicio
                Dim inicioCell As DataGridViewCell = grilla(colIndex, rowIndex)


                'Paso 2: Recurre al diccionario del portapapeles para conseguir sus valores
                Dim datosCopia As Dictionary(Of Integer, Dictionary(Of Integer, String)) = New Dictionary(Of Integer, Dictionary(Of Integer, String))()
                Dim lines As String() = datoPortapapeles.Split(vbLf)

                For i As Integer = 0 To lines.Length - 1
                    datosCopia(i) = New Dictionary(Of Integer, String)()
                    Dim lineContent As String() = lines(i).Split(vbTab)

                    'Si se copia una celda vacía, queda el valor de la celda en vacío
                    'Sino, ajusta el valor al diccionario del portapapeles

                    If lineContent.Length = 0 Then
                        datosCopia(i)(0) = String.Empty
                    Else

                        For j As Integer = 0 To lineContent.Length - 1
                            datosCopia(i)(j) = lineContent(j)
                        Next
                    End If
                Next

                Dim datosCopiados As Dictionary(Of Integer, Dictionary(Of Integer, String)) = datosCopia

                'Paso 3: Realiza el pegado
                Dim iRowIndex As Integer = inicioCell.RowIndex

                For Each rowKey As Integer In datosCopiados.Keys
                    Dim iColIndex As Integer = inicioCell.ColumnIndex

                    For Each cellKey As Integer In datosCopiados(rowKey).Keys

                        If iColIndex <= grilla.Columns.Count - 1 AndAlso iRowIndex <= grilla.Rows.Count - 1 Then
                            Dim cell As DataGridViewCell = grilla(iColIndex, iRowIndex)
                            cell.Value = datosCopiados(rowKey)(cellKey)
                        End If

                        iColIndex += 1
                    Next

                    iRowIndex += 1
                Next

                Return grilla

            Catch ex As Exception
                Return "ERROR: " + ex.Message.ToString()
            End Try
        End Function
    End Class

End Class
