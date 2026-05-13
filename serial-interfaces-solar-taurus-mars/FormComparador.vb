Imports System.IO
Imports System.Data
Imports OfficeOpenXml
Imports OfficeOpenXml.License
Imports System.ComponentModel

Public Class FormComparador
    Dim rutaArchivo1 As String = ""
    Dim rutaArchivo2 As String = ""
    Dim plcCounts1 As Dictionary(Of String, Integer)
    Dim plcCounts2 As Dictionary(Of String, Integer)

    ' Esto es necesario para usar la librería EPPlus sin errores de licencia
    Private Sub FormComparador_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ExcelPackage.License.SetNonCommercialOrganization("My Noncommercial organization")

        ' Archivos por defecto
        Dim defaultFile1 As String = "H:\Shared drives\Proyectos PG Control\P0000 TGN\P2224 TGN (HMI Planta PUE)\Datos de Partida\Serial Interface PUE\PDE4030_MC97D68_9002_999_Serial_Interface_Report.xlsx"
        Dim defaultFile2 As String = "H:\Shared drives\Proyectos PG Control\P0000 TGN\P2224 TGN (HMI Planta PUE)\Datos de Partida\Serial Interface LMR\PDD9970_MC97N67_0009_999_Serial_Interface_Report LMR_Mars 100.xlsx"

        If File.Exists(defaultFile1) Then
            rutaArchivo1 = defaultFile1
            lblRuta1.Text = "Archivo 1: " & Path.GetFileName(rutaArchivo1)
        End If
        If File.Exists(defaultFile2) Then
            rutaArchivo2 = defaultFile2
            lblRuta2.Text = "Archivo 2: " & Path.GetFileName(rutaArchivo2)
        End If
    End Sub

    ' --- BOTONES DE SELECCIÓN DE ARCHIVOS ---
    Private Sub btnArchivo1_Click(sender As Object, e As EventArgs) Handles btnArchivo1.Click
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Archivos Excel|*.xlsx;*.xls"
        If ofd.ShowDialog() = DialogResult.OK Then
            rutaArchivo1 = ofd.FileName
            lblRuta1.Text = "Archivo 1: " & Path.GetFileName(rutaArchivo1)
        End If
    End Sub

    Private Sub btnArchivo2_Click(sender As Object, e As EventArgs) Handles btnArchivo2.Click
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Archivos Excel|*.xlsx;*.xls"
        If ofd.ShowDialog() = DialogResult.OK Then
            rutaArchivo2 = ofd.FileName
            lblRuta2.Text = "Archivo 2: " & Path.GetFileName(rutaArchivo2)
        End If
    End Sub

    ' --- BOTÓN DE COMPARACIÓN ---
    Private Sub btnComparar_Click(sender As Object, e As EventArgs) Handles btnComparar.Click
        If String.IsNullOrEmpty(rutaArchivo1) OrElse String.IsNullOrEmpty(rutaArchivo2) Then
            MessageBox.Show("Por favor, elegí los dos archivos excel primero.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Try
            ' Leemos ambos archivos
            Dim tagsArchivo1 = LeerExcel(rutaArchivo1)
            Dim tagsArchivo2 = LeerExcel(rutaArchivo2)

            ' Generamos las tablas
            GenerarReportes(tagsArchivo1, tagsArchivo2)
        Catch ex As Exception
            MessageBox.Show("Error al leer los archivos. Asegurate de que no estén abiertos en Excel." & vbCrLf & "Detalle: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' --- CLASE PARA ALMACENAR LOS DATOS ---
    Class TagData
        Public Property TagName As String
        Public Property PlcAddress As String
        Public Property Description As String
    End Class

    ' --- LÓGICA DE LECTURA DE EXCEL ---
    Private Function LeerExcel(ruta As String) As Dictionary(Of String, TagData)

        ' Usamos un diccionario que ignora mayúsculas/minúsculas para no tener falsos positivos
        Dim dic As New Dictionary(Of String, TagData)(StringComparer.OrdinalIgnoreCase)

        Using package As New ExcelPackage(New FileInfo(ruta))
            Dim hoja = package.Workbook.Worksheets(0) ' Lee la primera hoja del Excel
            If hoja Is Nothing Then Return dic

            Dim totalFilas = If(hoja.Dimension IsNot Nothing, hoja.Dimension.Rows, 0)

            ' Empezamos a leer desde la fila 8
            For fila As Integer = 8 To totalFilas
                ' Columna 1 = Tag, Columna 3 = Dirección (ControlLogix), Columna 4 = Descripción
                Dim tagName As String = If(hoja.Cells(fila, 1).Value?.ToString()?.Trim(), "")

                If String.IsNullOrEmpty(tagName) Then Continue For ' Salta filas en blanco
                If tagName.Equals("Solar TT4000 HMI Tag Name", StringComparison.OrdinalIgnoreCase) Then Continue For ' Salta la fila de encabezado

                Dim plcAddress As String = If(hoja.Cells(fila, 3).Value?.ToString()?.Trim(), "")
                Dim descripcion As String = If(hoja.Cells(fila, 4).Value?.ToString()?.Trim(), "")

                ' Si el tag se repite en el mismo excel, guardamos solo el primero para evitar errores
                If Not dic.ContainsKey(tagName) Then
                    dic.Add(tagName, New TagData With {
                        .TagName = tagName,
                        .PlcAddress = plcAddress,
                        .Description = descripcion
                    })
                End If
            Next
        End Using
        Return dic
    End Function

    ' --- LÓGICA DE COMPARACIÓN ---
    Private Sub GenerarReportes(arch1 As Dictionary(Of String, TagData), arch2 As Dictionary(Of String, TagData))
        Dim dt1 As New DataTable()
        dt1.Columns.Add("Tag Name")
        dt1.Columns.Add("Resultado")
        dt1.Columns.Add("BDM")
        dt1.Columns.Add("Dir. PLC (Arch 1)")
        dt1.Columns.Add("Dir. PLC (Arch 2)")
        dt1.Columns.Add("Descripción (Arch 1)")
        dt1.Columns.Add("Descripción (Arch 2)")

        Dim allTags As New HashSet(Of String)(arch1.Keys, StringComparer.OrdinalIgnoreCase)
        allTags.UnionWith(arch2.Keys)

        plcCounts1 = arch1.Values.Where(Function(x) Not String.IsNullOrWhiteSpace(x.PlcAddress)).GroupBy(Function(x) x.PlcAddress.ToUpper()).ToDictionary(Function(g) g.Key, Function(g) g.Count())
        plcCounts2 = arch2.Values.Where(Function(x) Not String.IsNullOrWhiteSpace(x.PlcAddress)).GroupBy(Function(x) x.PlcAddress.ToUpper()).ToDictionary(Function(g) g.Key, Function(g) g.Count())

        For Each tagName As String In allTags
            Dim in1 = arch1.ContainsKey(tagName)
            Dim in2 = arch2.ContainsKey(tagName)

            Dim plc1 = If(in1, arch1(tagName).PlcAddress, "N/A")
            Dim desc1 = If(in1, arch1(tagName).Description, "N/A")
            Dim plc2 = If(in2, arch2(tagName).PlcAddress, "N/A")
            Dim desc2 = If(in2, arch2(tagName).Description, "N/A")

            Dim resultado As String = ""
            If in1 AndAlso in2 Then
                Dim difPlc As Boolean = (plc1 <> plc2)
                Dim difDesc As Boolean = (desc1 <> desc2)

                If Not difPlc AndAlso Not difDesc Then
                    resultado = "Adecuado"
                ElseIf difPlc AndAlso difDesc Then
                    resultado = "Dif. en Dir. PLC y Descripción"
                ElseIf difPlc Then
                    resultado = "Dif. en Dirección PLC"
                Else
                    resultado = "Dif. en Descripción"
                End If
            ElseIf in1 Then
                resultado = "Solo en Archivo 1"
            Else
                resultado = "Solo en Archivo 2"
            End If

            Dim bdm As String = If(resultado = "Adecuado", "Corregido", "Pendiente")

            dt1.Rows.Add(tagName, resultado, bdm, plc1, plc2, desc1, desc2)
        Next

        dgvReporte1.DataSource = dt1

        ' Reemplazar la columna "BDM" por un ComboBox
        Dim colIndex As Integer = dgvReporte1.Columns("BDM").Index
        dgvReporte1.Columns.RemoveAt(colIndex)

        Dim cmbBDM As New DataGridViewComboBoxColumn()
        cmbBDM.Name = "BDM"
        cmbBDM.HeaderText = "BDM"
        cmbBDM.DataPropertyName = "BDM"
        cmbBDM.Items.AddRange("Pendiente", "Corregido")
        cmbBDM.SortMode = DataGridViewColumnSortMode.Automatic
        cmbBDM.FlatStyle = FlatStyle.Flat
        dgvReporte1.Columns.Insert(colIndex, cmbBDM)

        ' Hacer de solo lectura las demás columnas
        For Each col As DataGridViewColumn In dgvReporte1.Columns
            If col.Name <> "BDM" Then
                col.ReadOnly = True
            End If
        Next

        ' Aplicar colores
        For Each row As DataGridViewRow In dgvReporte1.Rows
            ColorearFila(row)
        Next

        dgvReporte1.AutoResizeColumns()
    End Sub

    Private Sub ColorearFila(row As DataGridViewRow)
        If row.IsNewRow Then Return

        Dim valResultado As String = row.Cells("Resultado").Value?.ToString()
        Dim valBDM As String = ""
        If row.DataGridView.Columns.Contains("BDM") Then
            valBDM = row.Cells("BDM").Value?.ToString()
        End If

        Dim in1 = row.Cells("Dir. PLC (Arch 1)").Value.ToString() <> "N/A"
        Dim in2 = row.Cells("Dir. PLC (Arch 2)").Value.ToString() <> "N/A"
        Dim plc1 = row.Cells("Dir. PLC (Arch 1)").Value.ToString()
        Dim plc2 = row.Cells("Dir. PLC (Arch 2)").Value.ToString()

        Dim isDuplicated As Boolean = (in1 AndAlso Not String.IsNullOrWhiteSpace(plc1) AndAlso plcCounts1 IsNot Nothing AndAlso plcCounts1.ContainsKey(plc1.ToUpper()) AndAlso plcCounts1(plc1.ToUpper()) > 1) OrElse
                                      (in2 AndAlso Not String.IsNullOrWhiteSpace(plc2) AndAlso plcCounts2 IsNot Nothing AndAlso plcCounts2.ContainsKey(plc2.ToUpper()) AndAlso plcCounts2(plc2.ToUpper()) > 1)

        Dim backColor As Color = Color.White
        Dim foreColor As Color = Color.Black

        Select Case valResultado
            Case "Adecuado"
                backColor = Color.LightGreen
            Case "Dif. en Dir. PLC y Descripción", "Dif. en Dirección PLC"
                backColor = Color.LightBlue
            Case "Dif. en Descripción"
                backColor = Color.LightYellow
            Case "Solo en Archivo 1", "Solo en Archivo 2"
                backColor = Color.LightCoral
            Case Else
                backColor = Color.White
        End Select

        If isDuplicated Then
            backColor = Color.White
            foreColor = Color.Red
        End If

        If valBDM = "Corregido" AndAlso valResultado <> "Adecuado" Then
            backColor = Color.FromArgb(200, 255, 200) ' Verde mas clarito
            foreColor = Color.Black
        End If

        row.DefaultCellStyle.BackColor = backColor
        row.DefaultCellStyle.ForeColor = foreColor
        row.DefaultCellStyle.SelectionBackColor = backColor
        row.DefaultCellStyle.SelectionForeColor = foreColor

        ' Asegurar que la celda del selector también se pinte
        If row.DataGridView.Columns.Contains("BDM") Then
            row.Cells("BDM").Style.BackColor = backColor
            row.Cells("BDM").Style.ForeColor = foreColor
            row.Cells("BDM").Style.SelectionBackColor = backColor
            row.Cells("BDM").Style.SelectionForeColor = foreColor
        End If
    End Sub

    Private Sub dgvReporte1_CurrentCellDirtyStateChanged(sender As Object, e As EventArgs) Handles dgvReporte1.CurrentCellDirtyStateChanged
        If dgvReporte1.IsCurrentCellDirty Then
            dgvReporte1.CommitEdit(DataGridViewDataErrorContexts.Commit)
        End If
    End Sub

    Private Sub dgvReporte1_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvReporte1.CellValueChanged
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            If dgvReporte1.Columns(e.ColumnIndex).Name = "BDM" Then
                ColorearFila(dgvReporte1.Rows(e.RowIndex))

                ' Actualizar el color del control de edición si sigue activo
                If dgvReporte1.EditingControl IsNot Nothing Then
                    dgvReporte1.EditingControl.BackColor = dgvReporte1.Rows(e.RowIndex).Cells("BDM").Style.BackColor
                    dgvReporte1.EditingControl.ForeColor = dgvReporte1.Rows(e.RowIndex).Cells("BDM").Style.ForeColor
                End If
            End If
        End If
    End Sub

    Private Sub dgvReporte1_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles dgvReporte1.DataBindingComplete
        ' Restaurar colores al ordenar o filtrar la grilla
        If dgvReporte1.Columns.Contains("BDM") Then
            For Each row As DataGridViewRow In dgvReporte1.Rows
                ColorearFila(row)
            Next
        End If
    End Sub

    Private Sub dgvReporte1_EditingControlShowing(sender As Object, e As DataGridViewEditingControlShowingEventArgs) Handles dgvReporte1.EditingControlShowing
        ' Asegurarse de que el combobox en edición tome el color de la celda de inmediato
        If dgvReporte1.CurrentCell IsNot Nothing AndAlso dgvReporte1.Columns(dgvReporte1.CurrentCell.ColumnIndex).Name = "BDM" Then
            Dim combo As ComboBox = TryCast(e.Control, ComboBox)
            If combo IsNot Nothing Then
                combo.BackColor = dgvReporte1.CurrentCell.Style.BackColor
                combo.ForeColor = dgvReporte1.CurrentCell.Style.ForeColor
                combo.DrawMode = DrawMode.Normal ' No usamos pintado personalizado de items para BDM
            End If
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim dt As DataTable = TryCast(dgvReporte1.DataSource, DataTable)
        If dt Is Nothing Then
            MessageBox.Show("No hay datos para guardar. Realice una comparación primero.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Using sfd As New SaveFileDialog()
            sfd.Filter = "Archivos XML (*.xml)|*.xml"
            sfd.Title = "Guardar tabla como XML"
            sfd.FileName = "ComparadorTags.xml"
            If sfd.ShowDialog() = DialogResult.OK Then
                Try
                    ' Asignarle un nombre para que el proyecto principal la identifique como iFixListHeader
                    ' Asumimos que el prefijo es "iFix_" o la palabra clave que busca StartsWith(iFixListHeader)
                    dt.TableName = "iFix_ComparadorTags"

                    ' Guardarlo dentro de un DataSet ya que el programa principal lee desde un DataSet con ReadXml
                    Dim ds As New DataSet("ComparadorDataSet")

                    ' Hacemos una copia para no alterar el DataSource actual
                    ds.Tables.Add(dt.Copy())

                    ' Agregamos una tabla de metadatos para guardar los nombres de archivo usados
                    Dim dtMeta As New DataTable("Metadata")
                    dtMeta.Columns.Add("RutaArchivo1")
                    dtMeta.Columns.Add("RutaArchivo2")
                    dtMeta.Rows.Add(rutaArchivo1, rutaArchivo2)
                    ds.Tables.Add(dtMeta)

                    ' Guardamos con esquema para que ReadXml reconstruya la tabla correctamente
                    ds.WriteXml(sfd.FileName, XmlWriteMode.WriteSchema)
                    MessageBox.Show("Archivo guardado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Catch ex As Exception
                    MessageBox.Show("Error al guardar el archivo XML: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Using ofd As New OpenFileDialog()
            ofd.Filter = "Archivos XML (*.xml)|*.xml"
            ofd.Title = "Abrir tabla XML"
            If ofd.ShowDialog() = DialogResult.OK Then
                Try
                    Dim ds As New DataSet()
                    ds.ReadXml(ofd.FileName)

                    Dim dtMain As DataTable = Nothing
                    For Each tbl As DataTable In ds.Tables
                        If tbl.TableName = "iFix_ComparadorTags" Then
                            dtMain = tbl.Copy()
                        ElseIf tbl.TableName = "Metadata" Then
                            If tbl.Rows.Count > 0 Then
                                rutaArchivo1 = tbl.Rows(0)("RutaArchivo1").ToString()
                                rutaArchivo2 = tbl.Rows(0)("RutaArchivo2").ToString()
                                lblRuta1.Text = "Archivo 1: " & Path.GetFileName(rutaArchivo1)
                                lblRuta2.Text = "Archivo 2: " & Path.GetFileName(rutaArchivo2)
                            End If
                        End If
                    Next

                    If dtMain IsNot Nothing Then
                        ' Migrar XML viejo si no tiene columna BDM
                        If Not dtMain.Columns.Contains("BDM") Then
                            dtMain.Columns.Add("BDM", GetType(String))
                            For Each r As DataRow In dtMain.Rows
                                If r("Resultado").ToString() = "Adecuado" Then
                                    r("BDM") = "Corregido"
                                Else
                                    r("BDM") = "Pendiente"
                                End If
                            Next
                        End If

                        dgvReporte1.DataSource = dtMain

                        ' Reemplazar la columna "BDM" por un ComboBox para que sea editable
                        If dgvReporte1.Columns.Contains("BDM") AndAlso Not TypeOf dgvReporte1.Columns("BDM") Is DataGridViewComboBoxColumn Then
                            Dim colIndex As Integer = dgvReporte1.Columns("BDM").Index
                            dgvReporte1.Columns.RemoveAt(colIndex)

                            Dim cmbBDM As New DataGridViewComboBoxColumn()
                            cmbBDM.Name = "BDM"
                            cmbBDM.HeaderText = "BDM"
                            cmbBDM.DataPropertyName = "BDM"
                            cmbBDM.Items.AddRange("Pendiente", "Corregido")
                            cmbBDM.SortMode = DataGridViewColumnSortMode.Automatic
                            cmbBDM.FlatStyle = FlatStyle.Flat
                            dgvReporte1.Columns.Insert(colIndex, cmbBDM)
                        End If

                        ' Hacer de solo lectura las demás columnas
                        For Each col As DataGridViewColumn In dgvReporte1.Columns
                            If col.Name <> "BDM" Then
                                col.ReadOnly = True
                            End If
                        Next

                        ' Reconstruir conteos para colores de duplicados
                        plcCounts1 = New Dictionary(Of String, Integer)()
                        plcCounts2 = New Dictionary(Of String, Integer)()
                        For Each row As DataRow In dtMain.Rows
                            Dim p1 = row("Dir. PLC (Arch 1)").ToString().ToUpper()
                            Dim p2 = row("Dir. PLC (Arch 2)").ToString().ToUpper()
                            If p1 <> "N/A" AndAlso Not String.IsNullOrWhiteSpace(p1) Then
                                If plcCounts1.ContainsKey(p1) Then plcCounts1(p1) += 1 Else plcCounts1.Add(p1, 1)
                            End If
                            If p2 <> "N/A" AndAlso Not String.IsNullOrWhiteSpace(p2) Then
                                If plcCounts2.ContainsKey(p2) Then plcCounts2(p2) += 1 Else plcCounts2.Add(p2, 1)
                            End If
                        Next

                        dgvReporte1.AutoResizeColumns()
                        MessageBox.Show("Archivo cargado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("El archivo XML no contiene una tabla válida del comparador.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If

                Catch ex As Exception
                    MessageBox.Show("Error al abrir el archivo XML: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub

End Class