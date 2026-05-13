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
        btnComparar = New Button()
        lblRuta1 = New Label()
        lblRuta2 = New Label()
        dgvReporte1 = New DataGridView()
        MenuStrip1 = New MenuStrip()
        FileToolStripMenuItem = New ToolStripMenuItem()
        OpenToolStripMenuItem = New ToolStripMenuItem()
        SaveAsToolStripMenuItem = New ToolStripMenuItem()
        CType(dgvReporte1, ComponentModel.ISupportInitialize).BeginInit()
        MenuStrip1.SuspendLayout()
        SuspendLayout()
        ' 
        ' btnArchivo1
        ' 
        btnArchivo1.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnArchivo1.Location = New Point(622, 36)
        btnArchivo1.Name = "btnArchivo1"
        btnArchivo1.Size = New Size(150, 25)
        btnArchivo1.TabIndex = 0
        btnArchivo1.Text = "Elegir Archivo 1"
        btnArchivo1.UseVisualStyleBackColor = True
        ' 
        ' btnArchivo2
        ' 
        btnArchivo2.Anchor = AnchorStyles.Top Or AnchorStyles.Right
        btnArchivo2.Location = New Point(622, 67)
        btnArchivo2.Name = "btnArchivo2"
        btnArchivo2.Size = New Size(150, 25)
        btnArchivo2.TabIndex = 1
        btnArchivo2.Text = "Elegir Archivo 2"
        btnArchivo2.UseVisualStyleBackColor = True
        ' 
        ' btnComparar
        ' 
        btnComparar.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        btnComparar.Location = New Point(12, 102)
        btnComparar.Name = "btnComparar"
        btnComparar.Size = New Size(760, 30)
        btnComparar.TabIndex = 2
        btnComparar.Text = "Comparar Tags (Nueva Tabla)"
        btnComparar.UseVisualStyleBackColor = True
        ' 
        ' lblRuta1
        ' 
        lblRuta1.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblRuta1.AutoEllipsis = True
        lblRuta1.AutoSize = False
        lblRuta1.Location = New Point(12, 41)
        lblRuta1.Name = "lblRuta1"
        lblRuta1.Size = New Size(600, 15)
        lblRuta1.TabIndex = 3
        lblRuta1.Text = "Archivo 1 no seleccionado"
        ' 
        ' lblRuta2
        ' 
        lblRuta2.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right
        lblRuta2.AutoEllipsis = True
        lblRuta2.AutoSize = False
        lblRuta2.Location = New Point(12, 72)
        lblRuta2.Name = "lblRuta2"
        lblRuta2.Size = New Size(600, 15)
        lblRuta2.TabIndex = 4
        lblRuta2.Text = "Archivo 2 no seleccionado"
        ' 
        ' dgvReporte1
        ' 
        dgvReporte1.Anchor = AnchorStyles.Top Or AnchorStyles.Bottom Or AnchorStyles.Left Or AnchorStyles.Right
        dgvReporte1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        dgvReporte1.Location = New Point(12, 144)
        dgvReporte1.Name = "dgvReporte1"
        dgvReporte1.Size = New Size(760, 346)
        dgvReporte1.TabIndex = 5
        ' 
        ' MenuStrip1
        ' 
        MenuStrip1.Items.AddRange(New ToolStripItem() {FileToolStripMenuItem})
        MenuStrip1.Location = New Point(0, 0)
        MenuStrip1.Name = "MenuStrip1"
        MenuStrip1.Size = New Size(784, 24)
        MenuStrip1.TabIndex = 6
        MenuStrip1.Text = "MenuStrip1"
        ' 
        ' FileToolStripMenuItem
        ' 
        FileToolStripMenuItem.DropDownItems.AddRange(New ToolStripItem() {OpenToolStripMenuItem, SaveAsToolStripMenuItem})
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
        ClientSize = New Size(784, 502)
        MinimumSize = New Size(500, 400)
        Controls.Add(dgvReporte1)
        Controls.Add(lblRuta2)
        Controls.Add(lblRuta1)
        Controls.Add(btnComparar)
        Controls.Add(btnArchivo2)
        Controls.Add(btnArchivo1)
        Controls.Add(MenuStrip1)
        MainMenuStrip = MenuStrip1
        Name = "FormComparador"
        Text = "FormComparador"
        CType(dgvReporte1, ComponentModel.ISupportInitialize).EndInit()
        MenuStrip1.ResumeLayout(False)
        MenuStrip1.PerformLayout()
        ResumeLayout(False)
        PerformLayout()
    End Sub

    Friend WithEvents btnArchivo1 As Button
    Friend WithEvents btnArchivo2 As Button
    Friend WithEvents btnComparar As Button
    Friend WithEvents lblRuta1 As Label
    Friend WithEvents lblRuta2 As Label
    Friend WithEvents dgvReporte1 As DataGridView
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents FileToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents OpenToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents SaveAsToolStripMenuItem As ToolStripMenuItem

End Class
