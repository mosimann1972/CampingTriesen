<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Meldewesen
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
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

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle
        Me.lblSuchTitel = New System.Windows.Forms.Label
        Me.btnLoeschen = New System.Windows.Forms.Button
        Me.btnSuchen = New System.Windows.Forms.Button
        Me.btnMeldescheinDrucken = New System.Windows.Forms.Button
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.PrintDocument3 = New System.Drawing.Printing.PrintDocument
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'lblSuchTitel
        '
        Me.lblSuchTitel.AutoSize = True
        Me.lblSuchTitel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSuchTitel.Location = New System.Drawing.Point(36, 9)
        Me.lblSuchTitel.Name = "lblSuchTitel"
        Me.lblSuchTitel.Size = New System.Drawing.Size(60, 20)
        Me.lblSuchTitel.TabIndex = 140
        Me.lblSuchTitel.Text = "Suche"
        '
        'btnLoeschen
        '
        Me.btnLoeschen.Location = New System.Drawing.Point(713, 20)
        Me.btnLoeschen.Name = "btnLoeschen"
        Me.btnLoeschen.Size = New System.Drawing.Size(75, 23)
        Me.btnLoeschen.TabIndex = 135
        Me.btnLoeschen.Text = "Abbrechen"
        Me.btnLoeschen.UseVisualStyleBackColor = True
        '
        'btnSuchen
        '
        Me.btnSuchen.Location = New System.Drawing.Point(549, 20)
        Me.btnSuchen.Name = "btnSuchen"
        Me.btnSuchen.Size = New System.Drawing.Size(76, 23)
        Me.btnSuchen.TabIndex = 134
        Me.btnSuchen.Text = "Suchen"
        Me.btnSuchen.UseVisualStyleBackColor = True
        '
        'btnMeldescheinDrucken
        '
        Me.btnMeldescheinDrucken.Location = New System.Drawing.Point(631, 20)
        Me.btnMeldescheinDrucken.Name = "btnMeldescheinDrucken"
        Me.btnMeldescheinDrucken.Size = New System.Drawing.Size(76, 23)
        Me.btnMeldescheinDrucken.TabIndex = 141
        Me.btnMeldescheinDrucken.Text = "Drucken"
        Me.btnMeldescheinDrucken.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle1.SelectionBackColor = System.Drawing.Color.White
        Me.DataGridView1.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.BackgroundColor = System.Drawing.Color.White
        Me.DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(40, 63)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(748, 608)
        Me.DataGridView1.TabIndex = 142
        '
        'PrintDocument3
        '
        '
        'Meldewesen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(866, 701)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.btnMeldescheinDrucken)
        Me.Controls.Add(Me.lblSuchTitel)
        Me.Controls.Add(Me.btnLoeschen)
        Me.Controls.Add(Me.btnSuchen)
        Me.Name = "Meldewesen"
        Me.Text = "Meldewesen"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblSuchTitel As System.Windows.Forms.Label
    Friend WithEvents btnLoeschen As System.Windows.Forms.Button
    Friend WithEvents btnSuchen As System.Windows.Forms.Button
    Friend WithEvents btnMeldescheinDrucken As System.Windows.Forms.Button
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents PrintDocument3 As System.Drawing.Printing.PrintDocument
End Class
