Imports System.IO
Imports System.Data
Imports OfficeOpenXml
Imports OfficeOpenXml.License
Imports System.ComponentModel

Public Class FormComparador
    Dim rutaArchivo1 As String = ""
    Dim rutaArchivo2 As String = ""
    Dim rutaBDM As String = ""
    Dim rutaXmlGuardado As String = ""

    Private Sub FormComparador_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        cmbPrefijo.SelectedIndex = 0
        ExcelPackage.License.SetNonCommercialOrganization("My Noncommercial organization")

        Dim defaultFile1 As String = "H:\Shared drives\Proyectos PG Control\P0000 TGN\P2224 TGN (HMI Planta PUE)\Datos de Partida\Serial Interface PUE\PDE4030_MC97D68_9002_999_Serial_Interface_Report.xlsx"
        Dim defaultFile2 As String = "H:\Shared drives\Proyectos PG Control\P0000 TGN\P2224 TGN (HMI Planta PUE)\Datos de Partida\Serial Interface LMR\PDD9970_MC97N67_0009_999_Serial_Interface_Report LMR_Mars 100.xlsx"

        If File.Exists(defaultFile1) Then
            rutaArchivo1 = defaultFile1
            lblRuta1.Text = "1. Serial Interface Anterior: " & Path.GetFileName(rutaArchivo1)
        End If
        If File.Exists(defaultFile2) Then
            rutaArchivo2 = defaultFile2
            lblRuta2.Text = "2. Serial Interface Nuevo: " & Path.GetFileName(rutaArchivo2)
        End If
    End Sub

    ' --- BOTONES DE SELECCIÓN DE ARCHIVOS ---
    Private Sub btnArchivo1_Click(sender As Object, e As EventArgs) Handles btnArchivo1.Click
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Archivos Excel|*.xlsx;*.xls;*.xlsm"
        If ofd.ShowDialog() = DialogResult.OK Then
            rutaArchivo1 = ofd.FileName
            lblRuta1.Text = "1. Serial Interface Anterior: " & Path.GetFileName(rutaArchivo1)
        End If
    End Sub

    Private Sub btnArchivo2_Click(sender As Object, e As EventArgs) Handles btnArchivo2.Click
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Archivos Excel|*.xlsx;*.xls;*.xlsm"
        If ofd.ShowDialog() = DialogResult.OK Then
            rutaArchivo2 = ofd.FileName
            lblRuta2.Text = "2. Serial Interface Nuevo: " & Path.GetFileName(rutaArchivo2)
        End If
    End Sub

    Private Sub btnBDM_Click(sender As Object, e As EventArgs) Handles btnBDM.Click
        Dim ofd As New OpenFileDialog()
        ofd.Filter = "Archivos Excel|*.xlsx;*.xls;*.xlsm"
        If ofd.ShowDialog() = DialogResult.OK Then
            rutaBDM = ofd.FileName
            lblBDM.Text = "3. Base de Datos (BDM): " & Path.GetFileName(rutaBDM)
        End If
    End Sub

    ' --- CLASE PARA ALMACENAR LOS DATOS ---
    Class TagData
        Public Property TagName As String
        Public Property PlcAddress As String
        Public Property Description As String
    End Class

    ' --- LÓGICA DE LECTURA DE EXCEL ---
    Private Function LeerExcel(ruta As String) As Dictionary(Of String, TagData)
        Dim dic As New Dictionary(Of String, TagData)(StringComparer.OrdinalIgnoreCase)

        Using package As New ExcelPackage(New FileInfo(ruta))
            Dim hoja = package.Workbook.Worksheets(0)
            If hoja Is Nothing Then Return dic

            Dim totalFilas = If(hoja.Dimension IsNot Nothing, hoja.Dimension.Rows, 0)

            For fila As Integer = 8 To totalFilas
                Dim tagName As String = If(hoja.Cells(fila, 1).Value?.ToString()?.Trim(), "")
                If String.IsNullOrEmpty(tagName) Then Continue For
                If tagName.Equals("Solar TT4000 HMI Tag Name", StringComparison.OrdinalIgnoreCase) Then Continue For

                Dim plcAddress As String = If(hoja.Cells(fila, 3).Value?.ToString()?.Trim(), "")
                Dim descripcion As String = If(hoja.Cells(fila, 4).Value?.ToString()?.Trim(), "")

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
    Private Sub btnComparar_Click(sender As Object, e As EventArgs) Handles btnComparar.Click
        If String.IsNullOrEmpty(rutaArchivo1) OrElse String.IsNullOrEmpty(rutaArchivo2) OrElse String.IsNullOrEmpty(rutaBDM) Then
            MessageBox.Show("Por favor, elegí los 3 archivos excel primero.", "Atención", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        Dim prefijo As String = cmbPrefijo.SelectedItem.ToString()
        lblLog.ForeColor = Color.Blue
        lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Leyendo archivos..."
        Application.DoEvents()

        Try
            Dim tagsAnterior = LeerExcel(rutaArchivo1)
            Dim tagsNuevo = LeerExcel(rutaArchivo2)

            Dim dt As New DataTable()
            dt.Columns.Add("Tag Serial")
            dt.Columns.Add("Desc. Serial")
            dt.Columns.Add("Tag BDM")
            dt.Columns.Add("Desc. BDM")
            dt.Columns.Add("Hoja BDM")
            dt.Columns.Add("Dir. Anterior (Serial Anterior)")
            dt.Columns.Add("Dir. Nueva (Sugerida por Serial Nuevo)")
            dt.Columns.Add("Resultado Comparación")
            dt.Columns.Add("Dir. Actual (BDM)")
            dt.Columns.Add("Desc. Anterior")
            dt.Columns.Add("Desc. Nueva")
            dt.Columns.Add("Estado BDM")
            dt.Columns.Add("Fila BDM", GetType(Integer))

            Dim bdmData As New Dictionary(Of String, Tuple(Of String, String, String, Integer, String))(StringComparer.OrdinalIgnoreCase)
            Dim allBaseTags As New HashSet(Of String)(tagsAnterior.Keys, StringComparer.OrdinalIgnoreCase)
            allBaseTags.UnionWith(tagsNuevo.Keys)

            Using package As New ExcelPackage(New FileInfo(rutaBDM))
                Dim hojas As String() = {"BDM", "Analogicas"}
                Dim filasAPintar As New List(Of Tuple(Of String, Integer))

                For Each nombreHoja In hojas
                    Dim hoja = package.Workbook.Worksheets(nombreHoja)
                    If hoja IsNot Nothing Then
                        Dim totalFilas = If(hoja.Dimension IsNot Nothing, hoja.Dimension.Rows, 0)
                        For fila As Integer = 8 To totalFilas
                            Dim tagBdm As String = If(hoja.Cells(fila, 1).Value?.ToString()?.Trim(), "")
                            If String.IsNullOrEmpty(tagBdm) Then Continue For

                            If tagBdm.StartsWith(prefijo, StringComparison.OrdinalIgnoreCase) Then
                                Dim dirActual = If(hoja.Cells(fila, 5).Value?.ToString()?.Trim(), "")
                                Dim descBdm = If(hoja.Cells(fila, 3).Value?.ToString()?.Trim(), "")
                                Dim tagBase = tagBdm.Substring(prefijo.Length)

                                If Not bdmData.ContainsKey(tagBase) Then
                                    bdmData.Add(tagBase, New Tuple(Of String, String, String, Integer, String)(tagBdm, nombreHoja, dirActual, fila, descBdm))
                                End If
                                allBaseTags.Add(tagBase)
                            End If
                        Next
                    End If
                Next

                For Each tagBase In allBaseTags
                    Dim inAnterior = tagsAnterior.ContainsKey(tagBase)
                    Dim inNuevo = tagsNuevo.ContainsKey(tagBase)
                    Dim inBdm = bdmData.ContainsKey(tagBase)

                    Dim dirAnterior = If(inAnterior, tagsAnterior(tagBase).PlcAddress, "N/A")
                    Dim descAnterior = If(inAnterior, tagsAnterior(tagBase).Description, "N/A")

                    Dim dirNuevaRaw = If(inNuevo, tagsNuevo(tagBase).PlcAddress, "N/A")
                    Dim descNueva = If(inNuevo, tagsNuevo(tagBase).Description, "N/A")

                    Dim tagBdmStr = If(inBdm, bdmData(tagBase).Item1, "")
                    Dim tagSerialStr = If(inAnterior OrElse inNuevo, tagBase, "")

                    Dim hoja = If(inBdm, bdmData(tagBase).Item2, "No en BDM")
                    Dim dirActual = If(inBdm, bdmData(tagBase).Item3, "N/A")
                    Dim fila = If(inBdm, bdmData(tagBase).Item4, -1)
                    Dim descBdmStr = If(inBdm, bdmData(tagBase).Item5, "")

                    Dim descSerialStr = If(inNuevo AndAlso descNueva <> "N/A", descNueva, If(inAnterior, descAnterior, ""))

                    Dim dirSugerida = "N/A"
                    If inNuevo Then
                        Dim numTc As String = prefijo.Replace("PUE_TC", "").Replace("_", "")
                        Dim tcName As String = "TC" & numTc
                        If hoja.Equals("BDM", StringComparison.OrdinalIgnoreCase) Then
                            dirSugerida = $"RSLinx;{tcName};[{tcName}]{dirNuevaRaw};No Access Path;;/0"
                        ElseIf hoja.Equals("Analogicas", StringComparison.OrdinalIgnoreCase) Then
                            dirSugerida = $"RSLinx;{tcName};[{tcName}]{dirNuevaRaw};No Access Path"
                        Else
                            dirSugerida = $"RSLinx;{tcName};[{tcName}]{dirNuevaRaw};No Access Path"
                        End If
                    End If

                    Dim resultadoComparacion As String = ""
                    If inAnterior AndAlso inNuevo Then
                        Dim difPlc As Boolean = (dirAnterior <> dirNuevaRaw)
                        Dim difDesc As Boolean = (descAnterior <> descNueva)

                        If Not difPlc AndAlso Not difDesc Then
                            resultadoComparacion = "Sin Diferencias"
                        ElseIf difPlc AndAlso difDesc Then
                            resultadoComparacion = "Dif. en Dir. PLC y Descripción"
                        ElseIf difPlc Then
                            resultadoComparacion = "Dif. en Dirección PLC"
                        Else
                            resultadoComparacion = "Dif. en Descripción"
                        End If
                    ElseIf inAnterior Then
                        resultadoComparacion = "Solo en Serial Anterior"
                    ElseIf inNuevo Then
                        resultadoComparacion = "Solo en Serial Nuevo"
                    Else
                        resultadoComparacion = "No en Seriales"
                    End If

                    Dim estadoBDM As String = ""
                    If inBdm Then
                        If dirActual.Equals(dirSugerida, StringComparison.OrdinalIgnoreCase) Then
                            estadoBDM = "Adecuado"
                            If Not String.IsNullOrEmpty(dirSugerida) AndAlso dirSugerida <> "N/A" Then
                                filasAPintar.Add(New Tuple(Of String, Integer)(hoja, fila))
                            End If
                        ElseIf dirSugerida <> "N/A" Then
                            estadoBDM = "Pendiente"
                        Else
                            estadoBDM = "Pendiente"
                        End If
                    Else
                        estadoBDM = "No en BDM"
                    End If

                    dt.Rows.Add(tagSerialStr, descSerialStr, tagBdmStr, descBdmStr, hoja, dirAnterior, dirSugerida, resultadoComparacion, dirActual, descAnterior, descNueva, estadoBDM, fila)
                Next

                If filasAPintar.Count > 0 Then
                    Dim resp = MessageBox.Show($"Se encontraron {filasAPintar.Count} tags que ya tienen la dirección correcta en la BDM. ¿Desea marcarlos de color verde en el archivo de Excel ahora mismo?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                    If resp = DialogResult.Yes Then
                        For Each tupla In filasAPintar
                            Dim hojaPintar = package.Workbook.Worksheets(tupla.Item1)
                            Dim filaPintar = tupla.Item2
                            Dim maxCol As Integer = If(hojaPintar.Dimension IsNot Nothing, hojaPintar.Dimension.Columns, 10)
                            Using range = hojaPintar.Cells(filaPintar, 1, filaPintar, maxCol)
                                range.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80))
                            End Using
                        Next
                        package.Save()
                        lblLog.ForeColor = Color.DarkGreen
                        lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Éxito: Se marcaron {filasAPintar.Count} tags de verde en el Excel."
                    Else
                        lblLog.ForeColor = Color.Blue
                        lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Info: Coloreado automático omitido por el usuario."
                    End If
                Else
                    lblLog.ForeColor = Color.Blue
                    lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Info: Búsqueda completada."
                End If
            End Using

            dgvPrincipal.DataSource = dt
            ConfigurarGrilla()

        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ConfigurarGrilla()
        dgvPrincipal.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill
        If dgvPrincipal.Columns.Contains("Fila BDM") Then
            dgvPrincipal.Columns("Fila BDM").Visible = False
        End If
        If dgvPrincipal.Columns.Contains("Desc. Anterior") Then
            dgvPrincipal.Columns("Desc. Anterior").Visible = False
        End If
        If dgvPrincipal.Columns.Contains("Desc. Nueva") Then
            dgvPrincipal.Columns("Desc. Nueva").Visible = False
        End If

        If Not chkMostrarDescripciones.Checked Then
            If dgvPrincipal.Columns.Contains("Desc. Serial") Then dgvPrincipal.Columns("Desc. Serial").Visible = False
            If dgvPrincipal.Columns.Contains("Desc. BDM") Then dgvPrincipal.Columns("Desc. BDM").Visible = False
        Else
            If dgvPrincipal.Columns.Contains("Desc. Serial") Then dgvPrincipal.Columns("Desc. Serial").Visible = True
            If dgvPrincipal.Columns.Contains("Desc. BDM") Then dgvPrincipal.Columns("Desc. BDM").Visible = True
        End If

        If Not dgvPrincipal.Columns.Contains("Adecuar") Then
            Dim btnCol As New DataGridViewButtonColumn()
            btnCol.Name = "Adecuar"
            btnCol.HeaderText = "Acción"
            btnCol.Text = "Adecuar"
            btnCol.UseColumnTextForButtonValue = True
            dgvPrincipal.Columns.Add(btnCol)
        End If

        If dgvPrincipal.Columns.Contains("Estado BDM") AndAlso TypeOf dgvPrincipal.Columns("Estado BDM") IsNot DataGridViewComboBoxColumn Then
            Dim comboCol As New DataGridViewComboBoxColumn()
            comboCol.Name = "Estado BDM_Combo"
            comboCol.HeaderText = "Estado BDM"
            comboCol.DataPropertyName = "Estado BDM"
            comboCol.Items.AddRange("Adecuado", "Pendiente", "No en BDM", "Sin Sugerencia")
            comboCol.FlatStyle = FlatStyle.Flat
            Dim oldIdx = dgvPrincipal.Columns("Estado BDM").Index
            dgvPrincipal.Columns.Remove("Estado BDM")
            dgvPrincipal.Columns.Insert(oldIdx, comboCol)
            comboCol.Name = "Estado BDM"
        End If

        For Each col As DataGridViewColumn In dgvPrincipal.Columns
            If col.Name = "Estado BDM" Then
                col.ReadOnly = False
            ElseIf col.Name = "Tag BDM" Then
                col.ReadOnly = True
            Else
                col.ReadOnly = True
            End If
        Next

        ColorearGrilla()
    End Sub

    Private Sub chkMostrarDescripciones_CheckedChanged(sender As Object, e As EventArgs) Handles chkMostrarDescripciones.CheckedChanged
        ConfigurarGrilla()
    End Sub

    Private Sub ColorearGrilla()
        For Each row As DataGridViewRow In dgvPrincipal.Rows
            If row.IsNewRow Then Continue For

            Dim resComp As String = row.Cells("Resultado Comparación").Value?.ToString()
            Dim estBdm As String = row.Cells("Estado BDM").Value?.ToString()

            Dim backColor As Color = Color.White

            Select Case resComp
                Case "Sin Diferencias"
                    backColor = Color.LightGreen
                Case "Dif. en Dir. PLC y Descripción", "Dif. en Dirección PLC"
                    backColor = Color.LightBlue
                Case "Dif. en Descripción"
                    backColor = Color.LightYellow
                Case "Solo en Serial Viejo", "Solo en Serial Nuevo"
                    backColor = Color.LightCoral
            End Select

            If estBdm = "Adecuado" AndAlso resComp <> "Sin Diferencias" Then
                backColor = Color.FromArgb(200, 255, 200) ' Verde mas clarito
            ElseIf estBdm = "Adecuado" AndAlso resComp = "Sin Diferencias" Then
                backColor = Color.LightGreen
            End If

            row.DefaultCellStyle.BackColor = backColor
        Next
    End Sub

    Private Sub dgvPrincipal_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles dgvPrincipal.DataBindingComplete
        ColorearGrilla()
    End Sub

    Private Sub dgvPrincipal_SortStringChanged(sender As Object, e As Zuby.ADGV.AdvancedDataGridView.SortEventArgs) Handles dgvPrincipal.SortStringChanged
        Dim dt As DataTable = TryCast(dgvPrincipal.DataSource, DataTable)
        If dt IsNot Nothing Then
            dt.DefaultView.Sort = dgvPrincipal.SortString
        End If
    End Sub

    Private Sub dgvPrincipal_CellValueChanged(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPrincipal.CellValueChanged
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim colName = dgvPrincipal.Columns(e.ColumnIndex).Name
            If colName = "Estado BDM" Then
                ColorearGrilla()
            End If
        End If
    End Sub

    Private Sub dgvPrincipal_RowPostPaint(sender As Object, e As DataGridViewRowPostPaintEventArgs) Handles dgvPrincipal.RowPostPaint
        Dim grid = DirectCast(sender, DataGridView)
        Dim rowIdx = (e.RowIndex + 1).ToString()
        Dim centerFormat = New StringFormat() With {
            .Alignment = StringAlignment.Center,
            .LineAlignment = StringAlignment.Center
        }
        Dim headerBounds = New Rectangle(e.RowBounds.Left, e.RowBounds.Top, grid.RowHeadersWidth, e.RowBounds.Height)
        e.Graphics.DrawString(rowIdx, Me.Font, SystemBrushes.ControlText, headerBounds, centerFormat)
    End Sub

    Private Sub dgvPrincipal_CellDoubleClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPrincipal.CellDoubleClick
        If e.RowIndex >= 0 AndAlso e.ColumnIndex >= 0 Then
            Dim colName = dgvPrincipal.Columns(e.ColumnIndex).Name
            If colName = "Tag BDM" OrElse colName = "Tag Serial" Then
                Dim row = dgvPrincipal.Rows(e.RowIndex)
                Dim resComp = row.Cells("Resultado Comparación").Value?.ToString()
                Dim estadoBdm = row.Cells("Estado BDM").Value?.ToString()

                Dim isMissingBdm = (estadoBdm = "No en BDM")
                Dim isMissingSerial = (resComp = "No en Seriales")

                If Not isMissingBdm AndAlso Not isMissingSerial Then
                    Return ' La fila ya está matcheada o no aplica
                End If

                Dim dt As DataTable = TryCast(dgvPrincipal.DataSource, DataTable)
                If dt Is Nothing Then Return

                Dim huerfanos As New List(Of KeyValuePair(Of String, String))
                For Each dr As DataRow In dt.Rows
                    If dr.RowState = DataRowState.Deleted Then Continue For

                    Dim rc = dr("Resultado Comparación").ToString()
                    Dim est = dr("Estado BDM").ToString()

                    If isMissingBdm Then
                        If rc = "No en Seriales" Then
                            Dim t = dr("Tag BDM").ToString()
                            Dim d = dr("Desc. BDM").ToString()
                            huerfanos.Add(New KeyValuePair(Of String, String)(t, t & " | " & d))
                        End If
                    Else
                        If est = "No en BDM" Then
                            Dim t = dr("Tag Serial").ToString()
                            Dim d = dr("Desc. Serial").ToString()
                            huerfanos.Add(New KeyValuePair(Of String, String)(t, t & " | " & d))
                        End If
                    End If
                Next

                If huerfanos.Count = 0 Then
                    MessageBox.Show("No hay tags disponibles para matchear.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return
                End If

                Dim formSelector As New Form()
                formSelector.Text = If(isMissingBdm, "Seleccionar Tag BDM huérfano", "Seleccionar Tag Serial huérfano")
                formSelector.Size = New Size(400, 500)
                formSelector.StartPosition = FormStartPosition.CenterParent

                Dim txtBusqueda As New TextBox()
                txtBusqueda.Dock = DockStyle.Top
                txtBusqueda.Font = New Font("Segoe UI", 10)

                Dim lbTags As New ListBox()
                lbTags.Dock = DockStyle.Fill
                lbTags.Font = New Font("Segoe UI", 10)
                lbTags.DisplayMember = "Value"
                lbTags.ValueMember = "Key"
                lbTags.DataSource = huerfanos.ToArray()

                AddHandler txtBusqueda.TextChanged, Sub(s, ev)
                                                        Dim filtro = txtBusqueda.Text.ToLower()
                                                        Dim filtrados = huerfanos.Where(Function(h) h.Value.ToLower().Contains(filtro)).ToArray()
                                                        lbTags.DataSource = filtrados
                                                    End Sub

                Dim btnAceptar As New Button()
                btnAceptar.Text = "Vincular"
                btnAceptar.Dock = DockStyle.Bottom

                formSelector.Controls.Add(lbTags)
                formSelector.Controls.Add(txtBusqueda)
                formSelector.Controls.Add(btnAceptar)

                Dim selectedTag As String = ""

                AddHandler btnAceptar.Click, Sub(s, ev)
                                                 If lbTags.SelectedValue IsNot Nothing Then
                                                     selectedTag = lbTags.SelectedValue.ToString()
                                                     formSelector.DialogResult = DialogResult.OK
                                                 End If
                                                 formSelector.Close()
                                             End Sub

                AddHandler lbTags.DoubleClick, Sub(s, ev)
                                                   If lbTags.SelectedValue IsNot Nothing Then
                                                       selectedTag = lbTags.SelectedValue.ToString()
                                                       formSelector.DialogResult = DialogResult.OK
                                                       formSelector.Close()
                                                   End If
                                               End Sub

                If formSelector.ShowDialog() = DialogResult.OK AndAlso Not String.IsNullOrEmpty(selectedTag) Then
                    Dim filaBuscada As DataRow = Nothing
                    Dim drActual As DataRow = DirectCast(row.DataBoundItem, DataRowView).Row

                    For Each dr As DataRow In dt.Rows
                        If dr.RowState = DataRowState.Deleted Then Continue For
                        Dim rc = dr("Resultado Comparación").ToString()
                        Dim est = dr("Estado BDM").ToString()

                        If isMissingBdm Then
                            If dr("Tag BDM").ToString().Equals(selectedTag, StringComparison.OrdinalIgnoreCase) AndAlso rc = "No en Seriales" Then
                                filaBuscada = dr
                                Exit For
                            End If
                        Else
                            If dr("Tag Serial").ToString().Equals(selectedTag, StringComparison.OrdinalIgnoreCase) AndAlso est = "No en BDM" Then
                                filaBuscada = dr
                                Exit For
                            End If
                        End If
                    Next

                    If filaBuscada IsNot Nothing Then
                        Dim drSerial As DataRow
                        Dim drBdm As DataRow

                        If isMissingBdm Then
                            drSerial = drActual
                            drBdm = filaBuscada
                        Else
                            drSerial = filaBuscada
                            drBdm = drActual
                        End If

                        drSerial("Tag BDM") = drBdm("Tag BDM")
                        drSerial("Desc. BDM") = drBdm("Desc. BDM")
                        drSerial("Hoja BDM") = drBdm("Hoja BDM")
                        drSerial("Dir. Actual (BDM)") = drBdm("Dir. Actual (BDM)")
                        drSerial("Fila BDM") = drBdm("Fila BDM")
                        drBdm.Delete()

                        ' Verificar si la dirección ya coincide y actualizar estado en consecuencia
                        Dim dirSugeridaVinculo = drSerial("Dir. Nueva (Sugerida por Serial Nuevo)").ToString().Trim()
                        Dim dirActualVinculo = drSerial("Dir. Actual (BDM)").ToString().Trim()
                        Dim tagBdmVinculo = drSerial("Tag BDM").ToString()
                        Dim tagSerialVinculo = drSerial("Tag Serial").ToString()

                        If Not String.IsNullOrEmpty(dirSugeridaVinculo) AndAlso dirSugeridaVinculo <> "N/A" AndAlso
                           Not String.IsNullOrEmpty(dirActualVinculo) AndAlso dirActualVinculo <> "N/A" Then
                            If dirSugeridaVinculo.Equals(dirActualVinculo, StringComparison.OrdinalIgnoreCase) Then
                                drSerial("Estado BDM") = "Adecuado"
                                lblLog.ForeColor = Color.DarkGreen
                                lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] ✔ Vinculado: {tagSerialVinculo} → {tagBdmVinculo} | La dirección ya coincide ({dirActualVinculo}). Marcado como Adecuado."
                            Else
                                drSerial("Estado BDM") = "Pendiente"
                                lblLog.ForeColor = Color.DarkOrange
                                lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] ⚠ Vinculado: {tagSerialVinculo} → {tagBdmVinculo} | Dir. BDM actual: [{dirActualVinculo}] → Sugerida: [{dirSugeridaVinculo}]. Requiere adecuación."
                            End If
                        Else
                            drSerial("Estado BDM") = "Pendiente"
                            lblLog.ForeColor = Color.Blue
                            lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Info: {tagSerialVinculo} → {tagBdmVinculo} vinculados. Sin dirección para comparar aún."
                        End If

                        ColorearGrilla()
                    End If
                End If
            End If
        End If
    End Sub

    Private Sub dgvPrincipal_FilterStringChanged(sender As Object, e As Zuby.ADGV.AdvancedDataGridView.FilterEventArgs) Handles dgvPrincipal.FilterStringChanged
        Dim dt As DataTable = TryCast(dgvPrincipal.DataSource, DataTable)
        If dt IsNot Nothing Then
            dt.DefaultView.RowFilter = dgvPrincipal.FilterString
            ColorearGrilla()
        End If
    End Sub

    Private Sub dgvPrincipal_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvPrincipal.CellContentClick
        If e.RowIndex >= 0 AndAlso dgvPrincipal.Columns(e.ColumnIndex).Name = "Adecuar" Then
            Dim row As DataGridViewRow = dgvPrincipal.Rows(e.RowIndex)
            Dim estadoBDM As String = row.Cells("Estado BDM").Value.ToString()
            Dim dirSugerida As String = row.Cells("Dir. Nueva (Sugerida por Serial Nuevo)").Value.ToString()

            If estadoBDM = "Adecuado" Then
                lblLog.ForeColor = Color.Blue
                lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Info: El tag {row.Cells("Tag BDM").Value} ya estaba adecuado."
                Return
            End If

            If estadoBDM = "No en BDM" OrElse dirSugerida = "N/A" Then
                lblLog.ForeColor = Color.Red
                lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Error: El tag {row.Cells("Tag BDM").Value} no se puede adecuar automáticamente."
                Return
            End If

            Dim tagBdm As String = row.Cells("Tag BDM").Value.ToString()
            Dim hojaName As String = row.Cells("Hoja BDM").Value.ToString()
            Dim filaBdm As Integer = CInt(row.Cells("Fila BDM").Value)

            Try
                Using package As New ExcelPackage(New FileInfo(rutaBDM))
                    Dim hoja = package.Workbook.Worksheets(hojaName)
                    If hoja IsNot Nothing Then
                        ' Buscar la fila real por su Tag BDM para no perder la referencia
                        Dim filaReal As Integer = -1
                        Dim totalFilas = If(hoja.Dimension IsNot Nothing, hoja.Dimension.Rows, 0)
                        For f As Integer = 8 To totalFilas
                            If hoja.Cells(f, 1).Value?.ToString()?.Trim().Equals(tagBdm, StringComparison.OrdinalIgnoreCase) Then
                                filaReal = f
                                Exit For
                            End If
                        Next

                        If filaReal <> -1 Then
                            hoja.Cells(filaReal, 5).Value = dirSugerida

                            Dim maxCol As Integer = If(hoja.Dimension IsNot Nothing, hoja.Dimension.Columns, 10)
                            Using range = hoja.Cells(filaReal, 1, filaReal, maxCol)
                                range.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80))
                            End Using

                            package.Save()

                            row.Cells("Estado BDM").Value = "Adecuado"
                            row.Cells("Fila BDM").Value = filaReal
                            ColorearGrilla()

                            lblLog.ForeColor = Color.DarkGreen
                            lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Éxito: Tag {tagBdm} adecuado exitosamente."
                        Else
                            lblLog.ForeColor = Color.Red
                            lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Error: No se encontró el tag {tagBdm} en la hoja {hojaName} de la BDM."
                        End If
                    End If
                End Using
            Catch ex As Exception
                lblLog.ForeColor = Color.Red
                lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Error: {ex.Message}"
            End Try
        End If
    End Sub

    Private Sub btnAdecuarTodo_Click(sender As Object, e As EventArgs) Handles btnAdecuarTodo.Click
        If dgvPrincipal.Rows.Count = 0 OrElse String.IsNullOrEmpty(rutaBDM) Then
            lblLog.ForeColor = Color.Red
            lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Error: No hay datos o BDM cargado para adecuar."
            Return
        End If

        Dim pendientes As New List(Of DataGridViewRow)
        For Each row As DataGridViewRow In dgvPrincipal.Rows
            If Not row.IsNewRow AndAlso row.Cells("Estado BDM").Value.ToString() = "Pendiente" Then
                Dim dirSugerida As String = row.Cells("Dir. Nueva (Sugerida por Serial Nuevo)").Value.ToString()
                If dirSugerida <> "N/A" Then
                    pendientes.Add(row)
                End If
            End If
        Next

        If pendientes.Count = 0 Then
            lblLog.ForeColor = Color.Blue
            lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Info: No hay tags pendientes para adecuar."
            Return
        End If

        lblLog.ForeColor = Color.Blue
        lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Procesando {pendientes.Count} tags. Por favor, espere..."
        Application.DoEvents()

        Try
            Using package As New ExcelPackage(New FileInfo(rutaBDM))
                Dim cantAdecuados As Integer = 0
                For Each row In pendientes
                    Dim tagBdm As String = row.Cells("Tag BDM").Value.ToString()
                    Dim hojaName As String = row.Cells("Hoja BDM").Value.ToString()
                    Dim filaBdm As Integer = CInt(row.Cells("Fila BDM").Value)
                    Dim dirSugerida As String = row.Cells("Dir. Nueva (Sugerida por Serial Nuevo)").Value.ToString()

                    Dim hoja = package.Workbook.Worksheets(hojaName)
                    If hoja IsNot Nothing Then
                        ' Buscar la fila real por su Tag BDM
                        Dim filaReal As Integer = -1
                        Dim totalFilas = If(hoja.Dimension IsNot Nothing, hoja.Dimension.Rows, 0)
                        For f As Integer = 8 To totalFilas
                            If hoja.Cells(f, 1).Value?.ToString()?.Trim().Equals(tagBdm, StringComparison.OrdinalIgnoreCase) Then
                                filaReal = f
                                Exit For
                            End If
                        Next

                        If filaReal <> -1 Then
                            hoja.Cells(filaReal, 5).Value = dirSugerida

                            Dim maxCol As Integer = If(hoja.Dimension IsNot Nothing, hoja.Dimension.Columns, 10)
                            Using range = hoja.Cells(filaReal, 1, filaReal, maxCol)
                                range.Style.Fill.PatternType = Style.ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(146, 208, 80))
                            End Using

                            row.Cells("Estado BDM").Value = "Adecuado"
                            row.Cells("Fila BDM").Value = filaReal
                            ColorearGrilla()
                            cantAdecuados += 1
                        End If
                    End If
                Next

                package.Save()

                lblLog.ForeColor = Color.DarkGreen
                lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Éxito: Se adecuaron {cantAdecuados} tags en total."
            End Using
        Catch ex As Exception
            lblLog.ForeColor = Color.Red
            lblLog.Text = $"[{DateTime.Now.ToString("HH:mm:ss")}] Error masivo: {ex.Message}"
        End Try
    End Sub

    ' --- LÓGICA DE GUARDAR Y ABRIR XML ---
    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        If String.IsNullOrEmpty(rutaXmlGuardado) Then
            SaveAsToolStripMenuItem_Click(sender, e)
        Else
            GuardarXML(rutaXmlGuardado)
        End If
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        Dim dt As DataTable = TryCast(dgvPrincipal.DataSource, DataTable)
        If dt Is Nothing Then
            MessageBox.Show("No hay datos para guardar. Realice una comparación primero.", "Información", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Return
        End If

        Using sfd As New SaveFileDialog()
            sfd.Filter = "Archivos XML (*.xml)|*.xml"
            sfd.Title = "Guardar tabla como XML"
            sfd.FileName = "ComparadorTags.xml"
            If sfd.ShowDialog() = DialogResult.OK Then
                GuardarXML(sfd.FileName)
            End If
        End Using
    End Sub

    Private Sub GuardarXML(ruta As String)
        Dim dt As DataTable = TryCast(dgvPrincipal.DataSource, DataTable)
        If dt Is Nothing Then Return

        Try
            Dim ds As New DataSet("ComparadorDataSet")
            Dim dtCopy As DataTable = dt.Copy()
            dtCopy.TableName = "ComparadorPrincipal"
            ds.Tables.Add(dtCopy)

            Dim dtMeta As New DataTable("Metadata")
            dtMeta.Columns.Add("RutaArchivo1")
            dtMeta.Columns.Add("RutaArchivo2")
            dtMeta.Columns.Add("RutaBDM")
            dtMeta.Columns.Add("Prefijo")
            dtMeta.Rows.Add(rutaArchivo1, rutaArchivo2, rutaBDM, cmbPrefijo.SelectedItem?.ToString())
            ds.Tables.Add(dtMeta)

            ds.WriteXml(ruta, XmlWriteMode.WriteSchema)
            rutaXmlGuardado = ruta
            MessageBox.Show($"Archivo guardado exitosamente en:{vbCrLf}{ruta}", "Guardado", MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show("Error al guardar el archivo XML: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        Using ofd As New OpenFileDialog()
            ofd.Filter = "Archivos XML (*.xml)|*.xml"
            ofd.Title = "Abrir tabla XML"
            If ofd.ShowDialog() = DialogResult.OK Then
                Try
                    Dim ds As New DataSet()
                    ds.ReadXml(ofd.FileName)

                    Dim dtPrincipal As DataTable = Nothing
                    Dim esViejoFormato As Boolean = False

                    If ds.Tables.Contains("ComparadorPrincipal") Then
                        dtPrincipal = ds.Tables("ComparadorPrincipal").Copy()

                        If Not dtPrincipal.Columns.Contains("Desc. Serial") Then
                            dtPrincipal.Columns.Add("Desc. Serial").SetOrdinal(1)
                            For Each r As DataRow In dtPrincipal.Rows
                                Dim dN = If(Not IsDBNull(r("Desc. Nueva")), r("Desc. Nueva").ToString(), "N/A")
                                Dim dA = If(Not IsDBNull(r("Desc. Anterior")), r("Desc. Anterior").ToString(), "N/A")
                                r("Desc. Serial") = If(dN <> "N/A", dN, dA)
                            Next
                        End If
                        If Not dtPrincipal.Columns.Contains("Desc. BDM") Then
                            dtPrincipal.Columns.Add("Desc. BDM").SetOrdinal(3)
                        End If
                    Else
                    ' Migrar desde formato viejo
                    esViejoFormato = True
                    dtPrincipal = New DataTable()
                    dtPrincipal.Columns.Add("Tag Serial")
                    dtPrincipal.Columns.Add("Desc. Serial")
                    dtPrincipal.Columns.Add("Tag BDM")
                    dtPrincipal.Columns.Add("Desc. BDM")
                    dtPrincipal.Columns.Add("Hoja BDM")
                    dtPrincipal.Columns.Add("Dir. Anterior (Serial Anterior)")
                    dtPrincipal.Columns.Add("Dir. Nueva (Sugerida por Serial Nuevo)")
                    dtPrincipal.Columns.Add("Resultado Comparación")
                    dtPrincipal.Columns.Add("Dir. Actual (BDM)")
                    dtPrincipal.Columns.Add("Desc. Anterior")
                    dtPrincipal.Columns.Add("Desc. Nueva")
                    dtPrincipal.Columns.Add("Estado BDM")
                    dtPrincipal.Columns.Add("Fila BDM", GetType(Integer))

                    Dim tManual = If(ds.Tables.Contains("iFix_ComparadorTags"), ds.Tables("iFix_ComparadorTags"), Nothing)
                    Dim tAuto = If(ds.Tables.Contains("AutomaticaTags"), ds.Tables("AutomaticaTags"), Nothing)

                    Dim dictViejo As New Dictionary(Of String, DataRow)(StringComparer.OrdinalIgnoreCase)
                    If tManual IsNot Nothing Then
                        For Each row In tManual.Rows
                            Dim tag = row("Tag Name").ToString()
                            If Not dictViejo.ContainsKey(tag) Then dictViejo.Add(tag, row)
                        Next
                    End If

                    If tAuto IsNot Nothing Then
                        For Each row In tAuto.Rows
                            Dim tagBdm = row("Tag BDM").ToString()
                            Dim prefijo = "PUE_TC1_" ' Adivinar si no hay metadata
                            If ds.Tables.Contains("Metadata") AndAlso ds.Tables("Metadata").Columns.Contains("Prefijo") AndAlso Not IsDBNull(ds.Tables("Metadata").Rows(0)("Prefijo")) Then
                                prefijo = ds.Tables("Metadata").Rows(0)("Prefijo").ToString()
                            End If

                            Dim tagBase = tagBdm
                            If tagBdm.StartsWith(prefijo, StringComparison.OrdinalIgnoreCase) Then
                                tagBase = tagBdm.Substring(prefijo.Length)
                            End If

                            Dim rAnterior = If(dictViejo.ContainsKey(tagBase), dictViejo(tagBase), Nothing)
                            Dim dirAnterior = If(rAnterior IsNot Nothing, rAnterior("Dir. PLC (Arch 1)").ToString(), "N/A")
                            Dim descAnterior = If(rAnterior IsNot Nothing, rAnterior("Descripción (Arch 1)").ToString(), "N/A")
                            Dim descNueva = If(rAnterior IsNot Nothing, rAnterior("Descripción (Arch 2)").ToString(), "N/A")

                            Dim resComp = If(rAnterior IsNot Nothing, rAnterior("Resultado").ToString(), "No en Seriales")
                            Dim estBdm = row("Estado").ToString()
                            Dim dSerial = If(descNueva <> "N/A", descNueva, descAnterior)
                            dtPrincipal.Rows.Add(tagBase, dSerial, tagBdm, "", row("Hoja").ToString(), dirAnterior, row("Dir. Sugerida").ToString(), resComp, row("Dir. BDM").ToString(), descAnterior, descNueva, estBdm, Convert.ToInt32(row("Fila BDM")))

                            ' Eliminar para no duplicar si iteramos despues
                            If dictViejo.ContainsKey(tagBase) Then dictViejo.Remove(tagBase)
                        Next
                    End If

                    ' Agregar los que quedaron en manual y no en automatico
                    For Each kvp In dictViejo
                        Dim tagBase = kvp.Key
                        Dim r = kvp.Value
                        Dim dSerial = r("Descripción (Arch 1)").ToString()
                        dtPrincipal.Rows.Add(tagBase, dSerial, "N/A", "", "N/A", r("Dir. PLC (Arch 1)").ToString(), "N/A", r("Resultado").ToString(), "N/A", dSerial, r("Descripción (Arch 2)").ToString(), "No en BDM", -1)
                    Next
                    End If

                    If ds.Tables.Contains("Metadata") AndAlso ds.Tables("Metadata").Rows.Count > 0 Then
                        Dim tbl = ds.Tables("Metadata")
                        If tbl.Columns.Contains("RutaArchivo1") Then rutaArchivo1 = tbl.Rows(0)("RutaArchivo1").ToString()
                        If tbl.Columns.Contains("RutaArchivo2") Then rutaArchivo2 = tbl.Rows(0)("RutaArchivo2").ToString()
                        If tbl.Columns.Contains("RutaBDM") Then
                            rutaBDM = tbl.Rows(0)("RutaBDM").ToString()
                        ElseIf tbl.Columns.Contains("RutaBDMAuto") Then
                            rutaBDM = tbl.Rows(0)("RutaBDMAuto").ToString() ' Retrocompatibilidad
                        End If

                        If Not String.IsNullOrEmpty(rutaArchivo1) Then lblRuta1.Text = "1. Serial Interface Anterior: " & Path.GetFileName(rutaArchivo1)
                        If Not String.IsNullOrEmpty(rutaArchivo2) Then lblRuta2.Text = "2. Serial Interface Nuevo: " & Path.GetFileName(rutaArchivo2)
                        If Not String.IsNullOrEmpty(rutaBDM) Then lblBDM.Text = "3. Base de Datos (BDM): " & Path.GetFileName(rutaBDM)

                        If tbl.Columns.Contains("Prefijo") AndAlso Not IsDBNull(tbl.Rows(0)("Prefijo")) Then
                            Dim p As String = tbl.Rows(0)("Prefijo").ToString()
                            If cmbPrefijo.Items.Contains(p) Then cmbPrefijo.SelectedItem = p
                        End If
                    End If

                    If dtPrincipal IsNot Nothing Then
                        dgvPrincipal.DataSource = dtPrincipal
                        ConfigurarGrilla()

                        rutaXmlGuardado = ofd.FileName
                        MessageBox.Show("Archivo cargado exitosamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Else
                        MessageBox.Show("El archivo XML no contiene información válida.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    End If

                Catch ex As Exception
                    MessageBox.Show("Error al abrir el archivo XML: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
                End Try
            End If
        End Using
    End Sub
End Class
