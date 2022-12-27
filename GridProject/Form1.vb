'Componentes de edición de grillas en DLL.
'© 2022 Saúl Candia. - Holding Leonera.
'Programado para VisualBasic.NET
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports DLLFunciones.Funciones.FORMATFECHA
Imports DLLFunciones.Funciones.EXPORTARAEXCEL
Imports DLLFunciones.Funciones.IMPORTAREXCEL
Imports DLLFunciones.Funciones.CORTAR
Imports DLLFunciones.Funciones.COPIAR
Imports DLLFunciones.Funciones.PEGAR


Public Class Form1
    'Ajustar nombre de localhost/servidor y base de datos
    Public nombreHost = ".\SQLEXPRESS"
    Public nombreBD = "ejemplo"
    Dim Adpt As New SqlDataAdapter()
    Dim ds As New DataSet()

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Carga la grilla según la fuente. En este caso es una Base de datos LOCALHOST
        inicializarDataGridView()

    End Sub

    Private Sub inicializarDataGridView()

        'Inicializa la conexión con LOCALHOST al cargar el Form1
        Dim sConnectionString As String = "Data Source=" + nombreHost + ";Initial Catalog=" + nombreBD + ";Integrated Security=True"
        Dim Conn As New SqlConnection(sConnectionString)
        Dim Sql = "Select C1, C2, C3, convert(varchar, [Fecha], 5) as Fecha from Ejemplo"
        Adpt = New SqlDataAdapter(Sql, Conn)
        Adpt.Fill(ds, "Ejemplo")
        dgvPrincipal.DataSource = ds.Tables(0)

        'Que las columna C1, C2 Y C3 se vea pero no sea editable
        dgvPrincipal.Columns("C1").ReadOnly = True
        dgvPrincipal.Columns("C2").ReadOnly = True
        dgvPrincipal.Columns("C3").ReadOnly = True

        Dim dr As DataGridViewRow = dgvPrincipal.SelectedRows(0)
        lblC1.Text = dr.Cells(0).Value.ToString()
        lblC2.Text = dr.Cells(1).Value.ToString()
        lblC3.Text = dr.Cells(2).Value.ToString()
        lblFecha.Text = dr.Cells(3).Value.ToString()

        dgvPrincipal.Show()
    End Sub



    'El evento para guardar al presionar enter después de editar la celda es CellEndEdit
    Private Sub dgvPrincipal_CellEndEdit(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPrincipal.CellEndEdit
        Try
            'Prepara y reformatea la fecha (dia, mes y año)
            Dim dr As DataGridViewRow = dgvPrincipal.SelectedRows(0)
            lblFecha.Text = dr.Cells(3).Value.ToString()
            Dim stringfecha As String = lblFecha.Text
            Dim stringfechaformateada As String

            'Invoca a DLL para formateo de la fecha
            'PARÁMETROS A ENVIAR: STRING CON LA FECHA
            stringfechaformateada = FORMATEOFECHA(stringfecha)

            'Si hubo un error en el formateo, saldrá un mensaje de error
            If stringfechaformateada.Contains("ERROR") Then
                MessageBox.Show(stringfechaformateada, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                dgvPrincipal.Show()
            Else
                'Si no hubo error, prepara Query para actualizar datos
                Dim sConnectionString As String = "Data Source=" + nombreHost + ";Initial Catalog=" + nombreBD + ";Integrated Security=True"
                Dim Conn As New SqlConnection(sConnectionString)

                'Inserta fecha formateada actualizada
                Dim Sql = "Update Ejemplo set Fecha = CONVERT(DATETIME,'" + stringfechaformateada + "') where C1 = '" + lblC1.Text.ToString + "';"
                Dim comando As New SqlCommand(Sql, Conn)

                Conn.Open()
                comando.ExecuteNonQuery()
                Conn.Close()

                'Formateo de fechas SOLO PARA EFECTOS DE VISUALIZACIÓN DE GRILLA
                Dim mes As String = Microsoft.VisualBasic.Right(Microsoft.VisualBasic.Left(stringfechaformateada, 6), 2)
                Dim dia As String = Microsoft.VisualBasic.Right(stringfechaformateada, 2)
                Dim anio As String = Date.Today.Year.ToString

                'Muestra celda con valor actualizado
                dr.Cells(3).Value = dia + "-" + mes + "-" + anio

                'MessageBox.Show("Guardado y actualizado correctamente", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                dgvPrincipal.Update()
            End If
        Catch ex As Exception
            MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try

    End Sub

    'Al hacer clic en una fila, los datos se almacenan de manera TEMPORAL en etiquetas (LABEL)
    Private Sub dgvPrincipal_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPrincipal.CellClick
        Dim dr As DataGridViewRow = dgvPrincipal.SelectedRows(0)
        lblC1.Text = dr.Cells(0).Value.ToString()
        lblC2.Text = dr.Cells(1).Value.ToString()
        lblC3.Text = dr.Cells(2).Value.ToString()
        lblFecha.Text = dr.Cells(3).Value.ToString()
    End Sub

    'Si se presiona ENTER, se moverá hacia la celda de abajo
    Private Sub dgvPrincipal_KeyDown(sender As Object, e As KeyEventArgs) Handles dgvPrincipal.KeyDown
        If e.KeyCode = Keys.Enter Then
            'Prepara y reformatea la fecha (dia, mes y año)
            Dim dr As DataGridViewRow = dgvPrincipal.SelectedRows(0)
            lblFecha.Text = dr.Cells(2).Value.ToString()
            Dim stringfecha As String = lblFecha.Text

            SendKeys.Send("{down}")

        End If
    End Sub

    'Exportar a Excel
    Private Sub btnExportar_Click(sender As Object, e As EventArgs) Handles btnExportar.Click
        'Preparación de variables de resultado y ubicación del archivo a guardar
        Dim resultado As String
        Dim ubicacionArchivo As String

        'Dialogo guardar como
        Dim guardar As New SaveFileDialog 'Declara cuadro de dialogo
        guardar.Filter = "Hojas de cálculo (.xlsx)|*.xlsx"
        If guardar.ShowDialog = Windows.Forms.DialogResult.OK Then
            ubicacionArchivo = guardar.FileName

            'INVOCA A DLL PARA GUARDAR ARCHIVO (grilla, ubicación de archivo)
            'PARÁMETROS A ENVIAR: DATAGRIDVIEW COMPLETO Y STRING CON LA UBICACIÓN DEL ARCHIVO
            resultado = GUARDARAEXCEL(dgvPrincipal, ubicacionArchivo)

            'Mensaje que aparece según el resultado de la operación
            If resultado.Contains("ERROR") Then
                MessageBox.Show(resultado)
            Else
                MessageBox.Show("Archivo guardado exitosamente")
            End If
        End If

    End Sub

    Private Sub btnImportar_Click(sender As Object, e As EventArgs) Handles btnImportar.Click
        Dim resultado As Object
        Dim ubicacionArchivo As String

        'Dialogo abrir
        Dim abrir As New OpenFileDialog 'Declara cuadro de dialogo
        abrir.Filter = "Hojas de cálculo (.xlsx)|*.xlsx"
        If abrir.ShowDialog = Windows.Forms.DialogResult.OK Then
            ubicacionArchivo = abrir.FileName

            'INVOCA A DLL
            'PARÁMETROS A ENVIAR: STRING CON LA UBICACIÓN DEL ARCHIVO
            resultado = IMPORTARDESDEEXCEL(ubicacionArchivo)

            'Mensaje que aparece según el resultado de la operación
            If resultado.ToString.Contains("ERROR") Then
                MessageBox.Show("El archivo no tiene el formato correcto o está abierto. Debe cerrarlo antes de cargarlo en esta aplicación.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Else
                dgvPrincipal.DataSource = resultado
            End If
        End If


    End Sub

    Private Sub CortarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CortarToolStripMenuItem.Click

        Dim resultado As String
        'INVOCA A DLL
        'PARÁMETROS A ENVIAR: DATAGRIDVIEW COMPLETO
        resultado = CORTAR(dgvPrincipal)

        If resultado.ToString.Contains("ERROR") Then
            MessageBox.Show("No se pudo cortar hacia portapapeles", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MessageBox.Show("Cortado", "Bacán", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If

    End Sub

    Private Sub CopiarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopiarToolStripMenuItem.Click
        Dim resultado As String
        'INVOCA A DLL
        'PARÁMETROS A ENVIAR: DATAGRIDVIEW COMPLETO
        resultado = COPIAR(dgvPrincipal)

        If resultado.ToString.Contains("ERROR") Then
            MessageBox.Show("No se pudo copiar hacia portapapeles", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MessageBox.Show("Copiado", "Bacán", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub PegarToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles PegarToolStripMenuItem.Click

        'Si no hay filas seleccionadas, muestra un mensaje
        If dgvPrincipal.SelectedCells.Count = 0 Then
            MessageBox.Show("Por favor seleccione una fila", "Imposible pegar", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim resultado As DataGridView

        'INVOCA A DLL
        'PARÁMETROS A ENVIAR: DATAGRIDVIEW COMPLETO Y STRING CON LOS DATOS ALMACENADOS EN EL PORTAPAPELES
        resultado = PEGAR(dgvPrincipal, Clipboard.GetText)

        'Actualiza la vista de grilla con el resultado del procesamiento
        dgvPrincipal = resultado
        dgvPrincipal.Update()

        If resultado.ToString.Contains("ERROR") Then
            MessageBox.Show("No se pudo pegar desde portapapeles", "Imposible pegar", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        Else
            MessageBox.Show("PEGADO", "Bacán", MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub
End Class

