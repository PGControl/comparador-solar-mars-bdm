<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FormComparador
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        btnArchivo1 = New Button()
        btnArchivo2 = New Button()
        btnBDM = New Button()
        lblRuta1 = New Label()
        lblRuta2 = New Label()
        lblBDM = New Label()
        lblPrefijo = New Label()
        cmbPrefijo = New ComboBox()
        btnComparar = New Button()
        btnAdecuarTodo = New Button()
        dgvPrincipal = New Zuby.ADGV.AdvancedDataGridView()
        lblLog = New Label()
        MenuStrip1 = New MenuStrip()
        FileToolStripMenuItem = New ToolStripMenuItem()
        OpenToolStripMenuItem = New ToolStripMenuItem()
        SaveToolStripMenuItem = New ToolStripMenuItem()
        SaveAsToolStripMenuItem = New ToolStripMenuItem()
        CType(dgvPrincipal, ComponentModel.ISupportInitialize).BeginInit()
        MenuStrip1.SuspendLayout()
        SuspendLayout()
        ' 
        ' btnArchivo1
        ' 
        btnArchivo1.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnArchivo1.Location = New Point(800, 31)
        btnArchivo1.Name = "btnArchivo1"
        btnArchivo1.Size = New Size(150, 25)
        btnArchivo1.TabIndex = 0
        btnArchivo1.Text = "Examinar..."
        btnArchivo1.UseVisualStyleBackColor = True
        ' 
        ' btnArchivo2
        ' 
        btnArchivo2.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnArchivo2.Location = New Point(800, 60)
        btnArchivo2.Name = "btnArchivo2"
        btnArchivo2.Size = New Size(150, 25)
        btnArchivo2.TabIndex = 1
        btnArchivo2.Text = "Examinar..."
        btnArchivo2.UseVisualStyleBackColor = True
        ' 
        ' btnBDM
        ' 
        btnBDM.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnBDM.Location = New Point(800, 89)
        btnBDM.Name = "btnBDM"
        btnBDM.Size = New Size(150, 25)
        btnBDM.TabIndex = 2
        btnBDM.Text = "Examinar..."
        btnBDM.UseVisualStyleBackColor = True
        ' 
        ' lblRuta1
        ' 
        lblRuta1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblRuta1.AutoEllipsis = True
        lblRuta1.Location = New Point(12, 36)
        lblRuta1.Name = "lblRuta1"
        lblRuta1.Size = New Size(782, 15)
        lblRuta1.TabIndex = 3
        lblRuta1.Text = "1. Serial Interface Anterior: No seleccionado"
        ' 
        ' lblRuta2
        ' 
        lblRuta2.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblRuta2.AutoEllipsis = True
        lblRuta2.Location = New Point(12, 65)
        lblRuta2.Name = "lblRuta2"
        lblRuta2.Size = New Size(782, 15)
        lblRuta2.TabIndex = 4
        lblRuta2.Text = "2. Serial Interface Nuevo: No seleccionado"
        ' 
        ' lblBDM
        ' 
        lblBDM.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblBDM.AutoEllipsis = True
        lblBDM.Location = New Point(12, 94)
        lblBDM.Name = "lblBDM"
        lblBDM.Size = New Size(782, 15)
        lblBDM.TabIndex = 5
        lblBDM.Text = "3. Base de Datos (BDM): No seleccionado"
        ' 
        ' lblPrefijo
        ' 
        lblPrefijo.AutoSize = True
        lblPrefijo.Location = New Point(12, 124)
        lblPrefijo.Name = "lblPrefijo"
        lblPrefijo.Size = New Size(43, 15)
        lblPrefijo.TabIndex = 6
        lblPrefijo.Text = "Prefijo:"
        ' 
        ' cmbPrefijo
        ' 
        cmbPrefijo.DropDownStyle = ComboBoxStyle.DropDownList
        cmbPrefijo.FormattingEnabled = True
        cmbPrefijo.Items.AddRange(New Object() {"PUE_TC1_", "PUE_TC2_", "PUE_TC3_"})
        cmbPrefijo.Location = New Point(61, 121)
        cmbPrefijo.Name = "cmbPrefijo"
        cmbPrefijo.Size = New Size(121, 23)
        cmbPrefijo.TabIndex = 7
        ' 
        ' btnComparar
        ' 
        btnComparar.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        btnComparar.Location = New Point(188, 117)
        btnComparar.Name = "btnComparar"
        btnComparar.Size = New Size(556, 30)
        btnComparar.TabIndex = 8
        btnComparar.Text = "Comparar y Buscar para Adecuar"
        btnComparar.UseVisualStyleBackColor = True
        ' 
        ' btnAdecuarTodo
        ' 
        btnAdecuarTodo.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnAdecuarTodo.Location = New Point(750, 117)
        btnAdecuarTodo.Name = "btnAdecuarTodo"
        btnAdecuarTodo.Size = New Size(200, 30)
        btnAdecuarTodo.TabIndex = 9
        btnAdecuarTodo.Text = "Adecuar Todo"
        btnAdecuarTodo.UseVisualStyleBackColor = True
        ' 
        ' dgvPrincipal
        ' 
        dgvPrincipal.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        dgvPrincipal.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvPrincipal.Location = New Point(12, 153)
        dgvPrincipal.Name = "dgvPrincipal"
        dgvPrincipal.Size = New Size(938, 416)
        dgvPrincipal.TabIndex = 10
        dgvPrincipal.FilterAndSortEnabled = True
        ' 
        ' lblLog
        ' 
        lblLog.Anchor = AnchorStyles.Bottom Or AnchorStyles.Left
        lblLog.AutoSize = True
        lblLog.Font = New Font("Segoe UI", 9.0!, FontStyle.Bold, GraphicsUnit.Point, CByte(0))
        lblLog.ForeColor = Color.DarkGreen
        lblLog.Location = New Point(12, 577)
        lblLog.Name = "lblLog"
        lblLog.Size = New Size(116, 15)
        lblLog.TabIndex = 11
        lblLog.Text = "Esperando acción..."
        ' 
        ' MenuStrip1
        ' 
        MenuStrip1.Items.AddRange(New ToolStripItem() {FileToolStripMenuItem})
        MenuStrip1.Location = New Point(0, 0)
        MenuStrip1.Name = "MenuStrip1"
        MenuStrip1.Size = New Size(962, 24)
        MenuStrip1.TabIndex = 12
        MenuStrip1.Text = "MenuStrip1"
        ' 
        ' FileToolStripMenuItem
        ' 
        FileToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {OpenToolStripMenuItem, SaveToolStripMenuItem, SaveAsToolStripMenuItem})
        FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        FileToolStripMenuItem.Size = New Size(37, 20)
        FileToolStripMenuItem.Text = "File"
        ' 
        ' OpenToolStripMenuItem
        ' 
        OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        OpenToolStripMenuItem.Size = New Size(180, 22)
        OpenToolStripMenuItem.Text = "Open..."
        ' 
        ' SaveToolStripMenuItem
        ' 
        SaveToolStripMenuItem.Name = "SaveToolStripMenuItem"
        SaveToolStripMenuItem.Size = New Size(180, 22)
        SaveToolStripMenuItem.Text = "Save"
        ' 
        ' SaveAsToolStripMenuItem
        ' 
        SaveAsToolStripMenuItem.Name = "SaveAsToolStripMenuItem"
        SaveAsToolStripMenuItem.Size = New Size(180, 22)
        SaveAsToolStripMenuItem.Text = "Save as..."
        ' 
        ' FormComparador
        ' 
        AutoScaleDimensions = New SizeF(7.0F, 15.0F)
        AutoScaleMode = AutoScaleMode.Font
        ClientSize = New Size(962, 601)
        MinimumSize = New Size(800, 500)
        Controls.Add(dgvPrincipal)
        Controls.Add(btnAdecuarTodo)
        Controls.Add(btnComparar)
        Controls.Add(cmbPrefijo)
        Controls.Add(lblPrefijo)
        Controls.Add(lblBDM)
        Controls.Add(lblRuta2)
        Controls.Add(lblRuta1)
        Controls.Add(btnBDM)
        Controls.Add(btnArchivo2)
        Controls.Add(btnArchivo1)
        Controls.Add(lblLog)
        Controls.Add(MenuStrip1)
        MainMenuStrip = MenuStrip1
        Name = "FormComparador"
        Text = "Comparador Serial Interface y BDM"
        CType(dgvPrincipal, ComponentModel.ISupportInitialize).EndInit()
        MenuStrip1.ResumeLayout(False)
        MenuStrip1.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents btnArchivo1 As Button
    Friend WithEvents btnArchivo2 As Button
    Friend WithEvents btnBDM As Button
    Friend WithEvents lblRuta1 As Label
    Friend WithEvents lblRuta2 As Label
    Friend WithEvents lblBDM As Label
    Friend WithEvents lblPrefijo As Label
    Friend WithEvents cmbPrefijo As ComboBox
    Friend WithEvents btnComparar As Button
    Friend WithEvents btnAdecuarTodo As Button
    Friend WithEvents dgvPrincipal As Zuby.ADGV.AdvancedDataGridView
    Friend WithEvents lblLog As Label
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveAsToolStripMenuItem As ToolStripMenuItem

End Class
