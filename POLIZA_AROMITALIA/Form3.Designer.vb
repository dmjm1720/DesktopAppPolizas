<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Form3))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.IDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.TIPODataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.EMAILDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.CORREOSINSOFTECBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.COI7DataSet = New POLIZA_AROMITALIA.COI7DataSet()
        Me.TxtDelete = New System.Windows.Forms.TextBox()
        Me.BtnDelete = New System.Windows.Forms.Button()
        Me.TxtIdUpdate = New System.Windows.Forms.TextBox()
        Me.TxtMailUpdate = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.CboUpdate = New System.Windows.Forms.ComboBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.TxtMailInsert = New System.Windows.Forms.TextBox()
        Me.CboInsert = New System.Windows.Forms.ComboBox()
        Me.BtnInsert = New System.Windows.Forms.Button()
        Me.CORREOS_INSOFTECTableAdapter = New POLIZA_AROMITALIA.COI7DataSetTableAdapters.CORREOS_INSOFTECTableAdapter()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Panel3 = New System.Windows.Forms.Panel()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.CORREOSINSOFTECBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.COI7DataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.Panel1.SuspendLayout()
        Me.Panel2.SuspendLayout()
        Me.Panel3.SuspendLayout()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.AutoGenerateColumns = False
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IDDataGridViewTextBoxColumn, Me.TIPODataGridViewTextBoxColumn, Me.EMAILDataGridViewTextBoxColumn})
        Me.DataGridView1.DataSource = Me.CORREOSINSOFTECBindingSource
        Me.DataGridView1.Location = New System.Drawing.Point(2, 231)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(850, 218)
        Me.DataGridView1.TabIndex = 0
        '
        'IDDataGridViewTextBoxColumn
        '
        Me.IDDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.IDDataGridViewTextBoxColumn.DataPropertyName = "ID"
        Me.IDDataGridViewTextBoxColumn.HeaderText = "ID"
        Me.IDDataGridViewTextBoxColumn.Name = "IDDataGridViewTextBoxColumn"
        Me.IDDataGridViewTextBoxColumn.ReadOnly = True
        Me.IDDataGridViewTextBoxColumn.Width = 43
        '
        'TIPODataGridViewTextBoxColumn
        '
        Me.TIPODataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.TIPODataGridViewTextBoxColumn.DataPropertyName = "TIPO"
        Me.TIPODataGridViewTextBoxColumn.HeaderText = "TIPO"
        Me.TIPODataGridViewTextBoxColumn.Name = "TIPODataGridViewTextBoxColumn"
        Me.TIPODataGridViewTextBoxColumn.Width = 57
        '
        'EMAILDataGridViewTextBoxColumn
        '
        Me.EMAILDataGridViewTextBoxColumn.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells
        Me.EMAILDataGridViewTextBoxColumn.DataPropertyName = "EMAIL"
        Me.EMAILDataGridViewTextBoxColumn.HeaderText = "EMAIL"
        Me.EMAILDataGridViewTextBoxColumn.Name = "EMAILDataGridViewTextBoxColumn"
        Me.EMAILDataGridViewTextBoxColumn.Width = 64
        '
        'CORREOSINSOFTECBindingSource
        '
        Me.CORREOSINSOFTECBindingSource.DataMember = "CORREOS_INSOFTEC"
        Me.CORREOSINSOFTECBindingSource.DataSource = Me.COI7DataSet
        '
        'COI7DataSet
        '
        Me.COI7DataSet.DataSetName = "COI7DataSet"
        Me.COI7DataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'TxtDelete
        '
        Me.TxtDelete.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtDelete.Location = New System.Drawing.Point(503, 28)
        Me.TxtDelete.Name = "TxtDelete"
        Me.TxtDelete.Size = New System.Drawing.Size(100, 27)
        Me.TxtDelete.TabIndex = 1
        '
        'BtnDelete
        '
        Me.BtnDelete.BackColor = System.Drawing.Color.Khaki
        Me.BtnDelete.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnDelete.Location = New System.Drawing.Point(619, 17)
        Me.BtnDelete.Name = "BtnDelete"
        Me.BtnDelete.Size = New System.Drawing.Size(145, 38)
        Me.BtnDelete.TabIndex = 2
        Me.BtnDelete.Text = "Borrar registro"
        Me.BtnDelete.UseVisualStyleBackColor = False
        '
        'TxtIdUpdate
        '
        Me.TxtIdUpdate.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtIdUpdate.Location = New System.Drawing.Point(111, 27)
        Me.TxtIdUpdate.Name = "TxtIdUpdate"
        Me.TxtIdUpdate.Size = New System.Drawing.Size(100, 27)
        Me.TxtIdUpdate.TabIndex = 3
        '
        'TxtMailUpdate
        '
        Me.TxtMailUpdate.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtMailUpdate.Location = New System.Drawing.Point(392, 27)
        Me.TxtMailUpdate.Name = "TxtMailUpdate"
        Me.TxtMailUpdate.Size = New System.Drawing.Size(211, 27)
        Me.TxtMailUpdate.TabIndex = 5
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(107, 5)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(68, 19)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "Id Correo"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(224, 5)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(37, 19)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "Tipo"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(398, 5)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(50, 19)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "e-mail"
        '
        'CboUpdate
        '
        Me.CboUpdate.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboUpdate.FormattingEnabled = True
        Me.CboUpdate.Items.AddRange(New Object() {"Ventas", "Compras"})
        Me.CboUpdate.Location = New System.Drawing.Point(217, 27)
        Me.CboUpdate.Name = "CboUpdate"
        Me.CboUpdate.Size = New System.Drawing.Size(169, 27)
        Me.CboUpdate.TabIndex = 9
        '
        'Button1
        '
        Me.Button1.BackColor = System.Drawing.Color.DarkKhaki
        Me.Button1.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Button1.Location = New System.Drawing.Point(619, 16)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(145, 38)
        Me.Button1.TabIndex = 10
        Me.Button1.Text = "Actualizar registro"
        Me.Button1.UseVisualStyleBackColor = False
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(499, 6)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(68, 19)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Id Correo"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(398, 1)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(50, 19)
        Me.Label5.TabIndex = 14
        Me.Label5.Text = "e-mail"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(224, 1)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(37, 19)
        Me.Label6.TabIndex = 13
        Me.Label6.Text = "Tipo"
        '
        'TxtMailInsert
        '
        Me.TxtMailInsert.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TxtMailInsert.Location = New System.Drawing.Point(392, 23)
        Me.TxtMailInsert.Name = "TxtMailInsert"
        Me.TxtMailInsert.Size = New System.Drawing.Size(211, 27)
        Me.TxtMailInsert.TabIndex = 17
        '
        'CboInsert
        '
        Me.CboInsert.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CboInsert.FormattingEnabled = True
        Me.CboInsert.Items.AddRange(New Object() {"Ventas", "Compras"})
        Me.CboInsert.Location = New System.Drawing.Point(217, 23)
        Me.CboInsert.Name = "CboInsert"
        Me.CboInsert.Size = New System.Drawing.Size(169, 27)
        Me.CboInsert.TabIndex = 18
        '
        'BtnInsert
        '
        Me.BtnInsert.BackColor = System.Drawing.Color.Khaki
        Me.BtnInsert.Font = New System.Drawing.Font("Calibri", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BtnInsert.Location = New System.Drawing.Point(619, 12)
        Me.BtnInsert.Name = "BtnInsert"
        Me.BtnInsert.Size = New System.Drawing.Size(145, 38)
        Me.BtnInsert.TabIndex = 19
        Me.BtnInsert.Text = "Guardar registro"
        Me.BtnInsert.UseVisualStyleBackColor = False
        '
        'CORREOS_INSOFTECTableAdapter
        '
        Me.CORREOS_INSOFTECTableAdapter.ClearBeforeFill = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.Panel1.Controls.Add(Me.Label6)
        Me.Panel1.Controls.Add(Me.BtnInsert)
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.CboInsert)
        Me.Panel1.Controls.Add(Me.TxtMailInsert)
        Me.Panel1.Location = New System.Drawing.Point(51, 8)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(775, 62)
        Me.Panel1.TabIndex = 20
        Me.Panel1.Tag = ""
        '
        'Panel2
        '
        Me.Panel2.BackColor = System.Drawing.Color.SkyBlue
        Me.Panel2.Controls.Add(Me.TxtIdUpdate)
        Me.Panel2.Controls.Add(Me.TxtMailUpdate)
        Me.Panel2.Controls.Add(Me.Label2)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.Button1)
        Me.Panel2.Controls.Add(Me.Label3)
        Me.Panel2.Controls.Add(Me.CboUpdate)
        Me.Panel2.Location = New System.Drawing.Point(51, 76)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(775, 62)
        Me.Panel2.TabIndex = 21
        '
        'Panel3
        '
        Me.Panel3.BackColor = System.Drawing.SystemColors.InactiveCaption
        Me.Panel3.Controls.Add(Me.Label4)
        Me.Panel3.Controls.Add(Me.TxtDelete)
        Me.Panel3.Controls.Add(Me.BtnDelete)
        Me.Panel3.Location = New System.Drawing.Point(51, 144)
        Me.Panel3.Name = "Panel3"
        Me.Panel3.Size = New System.Drawing.Size(775, 62)
        Me.Panel3.TabIndex = 22
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(854, 447)
        Me.Controls.Add(Me.Panel3)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "Form3"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Administración de correos"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.CORREOSINSOFTECBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.COI7DataSet, System.ComponentModel.ISupportInitialize).EndInit()
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.Panel3.ResumeLayout(False)
        Me.Panel3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents TxtDelete As TextBox
    Friend WithEvents BtnDelete As Button
    Friend WithEvents TxtIdUpdate As TextBox
    Friend WithEvents TxtMailUpdate As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents CboUpdate As ComboBox
    Friend WithEvents Button1 As Button
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents TxtMailInsert As TextBox
    Friend WithEvents CboInsert As ComboBox
    Friend WithEvents BtnInsert As Button
    Friend WithEvents COI7DataSet As COI7DataSet
    Friend WithEvents CORREOSINSOFTECBindingSource As BindingSource
    Friend WithEvents CORREOS_INSOFTECTableAdapter As COI7DataSetTableAdapters.CORREOS_INSOFTECTableAdapter
    Friend WithEvents IDDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents TIPODataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents EMAILDataGridViewTextBoxColumn As DataGridViewTextBoxColumn
    Friend WithEvents Panel1 As Panel
    Friend WithEvents Panel2 As Panel
    Friend WithEvents Panel3 As Panel
End Class
