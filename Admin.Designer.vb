<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Admin
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
		Me.Label3 = New System.Windows.Forms.Label()
		Me.Label1 = New System.Windows.Forms.Label()
		Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
		Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
		Me.Button1 = New System.Windows.Forms.Button()
		Me.Button2 = New System.Windows.Forms.Button()
		Me.Button3 = New System.Windows.Forms.Button()
		Me.SuspendLayout()
		'
		'Label3
		'
		Me.Label3.AutoSize = True
		Me.Label3.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
		Me.Label3.Location = New System.Drawing.Point(40, 24)
		Me.Label3.Name = "Label3"
		Me.Label3.Size = New System.Drawing.Size(59, 20)
		Me.Label3.TabIndex = 130
		Me.Label3.Text = "Admin"
		'
		'Label1
		'
		Me.Label1.AutoSize = True
		Me.Label1.Location = New System.Drawing.Point(40, 72)
		Me.Label1.Name = "Label1"
		Me.Label1.Size = New System.Drawing.Size(198, 13)
		Me.Label1.TabIndex = 131
		Me.Label1.Text = "Neues Jahr für Zimmerbuchung erfassen"
		'
		'DateTimePicker1
		'
		Me.DateTimePicker1.Location = New System.Drawing.Point(40, 104)
		Me.DateTimePicker1.Name = "DateTimePicker1"
		Me.DateTimePicker1.Size = New System.Drawing.Size(200, 20)
		Me.DateTimePicker1.TabIndex = 149
		'
		'DateTimePicker2
		'
		Me.DateTimePicker2.Location = New System.Drawing.Point(256, 104)
		Me.DateTimePicker2.Name = "DateTimePicker2"
		Me.DateTimePicker2.Size = New System.Drawing.Size(200, 20)
		Me.DateTimePicker2.TabIndex = 150
		'
		'Button1
		'
		Me.Button1.Location = New System.Drawing.Point(488, 104)
		Me.Button1.Name = "Button1"
		Me.Button1.Size = New System.Drawing.Size(75, 23)
		Me.Button1.TabIndex = 151
		Me.Button1.Text = "Go"
		Me.Button1.UseVisualStyleBackColor = True
		'
		'Button2
		'
		Me.Button2.Location = New System.Drawing.Point(592, 104)
		Me.Button2.Name = "Button2"
		Me.Button2.Size = New System.Drawing.Size(128, 23)
		Me.Button2.TabIndex = 152
		Me.Button2.Text = "Akt. Jahr löschen"
		Me.Button2.UseVisualStyleBackColor = True
		'
		'Button3
		'
		Me.Button3.Location = New System.Drawing.Point(640, 480)
		Me.Button3.Name = "Button3"
		Me.Button3.Size = New System.Drawing.Size(75, 23)
		Me.Button3.TabIndex = 153
		Me.Button3.Text = "schliessen"
		Me.Button3.UseVisualStyleBackColor = True
		'
		'Admin
		'
		Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
		Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
		Me.ClientSize = New System.Drawing.Size(843, 604)
		Me.Controls.Add(Me.Button3)
		Me.Controls.Add(Me.Button2)
		Me.Controls.Add(Me.Button1)
		Me.Controls.Add(Me.DateTimePicker2)
		Me.Controls.Add(Me.DateTimePicker1)
		Me.Controls.Add(Me.Label1)
		Me.Controls.Add(Me.Label3)
		Me.Name = "Admin"
		Me.Text = "Admin"
		Me.ResumeLayout(False)
		Me.PerformLayout()

End Sub
	Friend WithEvents Label3 As System.Windows.Forms.Label
 Friend WithEvents Label1 As System.Windows.Forms.Label
 Friend WithEvents DateTimePicker1 As System.Windows.Forms.DateTimePicker
 Friend WithEvents DateTimePicker2 As System.Windows.Forms.DateTimePicker
 Friend WithEvents Button1 As System.Windows.Forms.Button
 Friend WithEvents Button2 As System.Windows.Forms.Button
 Friend WithEvents Button3 As System.Windows.Forms.Button
End Class
