<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AnAbreise
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
		Me.lblSuchTitel1 = New System.Windows.Forms.Label()
		Me.DataGridView1 = New System.Windows.Forms.DataGridView()
		Me.btnLoeschen = New System.Windows.Forms.Button()
		Me.btnSuchen = New System.Windows.Forms.Button()
		Me.lblSuchTitel2 = New System.Windows.Forms.Label()
		Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
		Me.PrintDocument2 = New System.Drawing.Printing.PrintDocument()
		Me.btnPrint = New System.Windows.Forms.Button()
		CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
		Me.SuspendLayout()
		'
		'lblSuchTitel1
		'
		Me.lblSuchTitel1.AutoSize = True
		Me.lblSuchTitel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.lblSuchTitel1.Location = New System.Drawing.Point(36, 9)
		Me.lblSuchTitel1.Name = "lblSuchTitel1"
		Me.lblSuchTitel1.Size = New System.Drawing.Size(60, 20)
		Me.lblSuchTitel1.TabIndex = 140
		Me.lblSuchTitel1.Text = "Suche"
		'
		'DataGridView1
		'
		Me.DataGridView1.BackgroundColor = System.Drawing.Color.White
		Me.DataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.None
		Me.DataGridView1.Location = New System.Drawing.Point(40, 63)
		Me.DataGridView1.Name = "DataGridView1"
		Me.DataGridView1.Size = New System.Drawing.Size(748, 608)
		Me.DataGridView1.TabIndex = 137
		'
		'btnLoeschen
		'
		Me.btnLoeschen.Location = New System.Drawing.Point(605, 20)
		Me.btnLoeschen.Name = "btnLoeschen"
		Me.btnLoeschen.Size = New System.Drawing.Size(75, 23)
		Me.btnLoeschen.TabIndex = 135
		Me.btnLoeschen.Text = "Abbrechen"
		Me.btnLoeschen.UseVisualStyleBackColor = True
		'
		'btnSuchen
		'
		Me.btnSuchen.Location = New System.Drawing.Point(524, 20)
		Me.btnSuchen.Name = "btnSuchen"
		Me.btnSuchen.Size = New System.Drawing.Size(75, 23)
		Me.btnSuchen.TabIndex = 134
		Me.btnSuchen.Text = "Suchen"
		Me.btnSuchen.UseVisualStyleBackColor = True
		'
		'lblSuchTitel2
		'
		Me.lblSuchTitel2.AutoSize = True
		Me.lblSuchTitel2.Location = New System.Drawing.Point(192, 27)
		Me.lblSuchTitel2.Name = "lblSuchTitel2"
		Me.lblSuchTitel2.Size = New System.Drawing.Size(102, 13)
		Me.lblSuchTitel2.TabIndex = 136
		Me.lblSuchTitel2.Text = "Bitte Datum wählen:"
		'
		'DateTimePicker1
		'
		Me.DateTimePicker1.Location = New System.Drawing.Point(292, 24)
		Me.DateTimePicker1.Name = "DateTimePicker1"
		Me.DateTimePicker1.Size = New System.Drawing.Size(202, 20)
		Me.DateTimePicker1.TabIndex = 143
		'
		'PrintDocument2
		'
		'
		'btnPrint
		'
		Me.btnPrint.Location = New System.Drawing.Point(713, 20)
		Me.btnPrint.Name = "btnPrint"
		Me.btnPrint.Size = New System.Drawing.Size(75, 23)
		Me.btnPrint.TabIndex = 144
		Me.btnPrint.Text = "Drucken"
		Me.btnPrint.UseVisualStyleBackColor = True
		'
		'AnAbreise
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(866, 701)
		Me.Controls.Add(Me.btnPrint)
		Me.Controls.Add(Me.DateTimePicker1)
		Me.Controls.Add(Me.lblSuchTitel1)
		Me.Controls.Add(Me.DataGridView1)
		Me.Controls.Add(Me.btnLoeschen)
		Me.Controls.Add(Me.btnSuchen)
		Me.Controls.Add(Me.lblSuchTitel2)
		Me.Name = "AnAbreise"
		Me.Text = "AnAbreise"
		CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
		Me.ResumeLayout(False)
		Me.PerformLayout()

End Sub
    Friend WithEvents lblSuchTitel1 As System.Windows.Forms.Label
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents btnLoeschen As System.Windows.Forms.Button
    Friend WithEvents btnSuchen As System.Windows.Forms.Button
    Friend WithEvents lblSuchTitel2 As System.Windows.Forms.Label
    Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
    Friend WithEvents PrintDocument2 As System.Drawing.Printing.PrintDocument
    Friend WithEvents btnPrint As System.Windows.Forms.Button
End Class
