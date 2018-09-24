<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AdressenErfassen
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
        Me.txtNachName = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.txtVorname = New System.Windows.Forms.TextBox()
        Me.txtAdresse = New System.Windows.Forms.TextBox()
        Me.txtPLZ = New System.Windows.Forms.TextBox()
        Me.txtOrt = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.txtGeburtsdatum = New System.Windows.Forms.TextBox()
        Me.txtEMailAdresse = New System.Windows.Forms.TextBox()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.txtAusweisNummer = New System.Windows.Forms.TextBox()
        Me.txtId = New System.Windows.Forms.RadioButton()
        Me.txtPass = New System.Windows.Forms.RadioButton()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.btnSpeichern = New System.Windows.Forms.Button()
        Me.btnAbbrechen = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.PrintDocument1 = New System.Drawing.Printing.PrintDocument()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.RadioButton3 = New System.Windows.Forms.RadioButton()
        Me.RadioButton4 = New System.Windows.Forms.RadioButton()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.cmbNation = New System.Windows.Forms.ComboBox()
        Me.PrintDocument2 = New System.Drawing.Printing.PrintDocument()
        Me.cmbLand = New System.Windows.Forms.ComboBox()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'txtNachName
        '
        Me.txtNachName.Location = New System.Drawing.Point(97, 33)
        Me.txtNachName.Name = "txtNachName"
        Me.txtNachName.Size = New System.Drawing.Size(367, 20)
        Me.txtNachName.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(34, 36)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(35, 13)
        Me.Label1.TabIndex = 36
        Me.Label1.Text = "Name"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(34, 62)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(49, 13)
        Me.Label2.TabIndex = 37
        Me.Label2.Text = "Vorname"
        '
        'txtVorname
        '
        Me.txtVorname.Location = New System.Drawing.Point(97, 59)
        Me.txtVorname.Name = "txtVorname"
        Me.txtVorname.Size = New System.Drawing.Size(367, 20)
        Me.txtVorname.TabIndex = 1
        '
        'txtAdresse
        '
        Me.txtAdresse.Location = New System.Drawing.Point(97, 85)
        Me.txtAdresse.Name = "txtAdresse"
        Me.txtAdresse.Size = New System.Drawing.Size(367, 20)
        Me.txtAdresse.TabIndex = 2
        '
        'txtPLZ
        '
        Me.txtPLZ.Location = New System.Drawing.Point(97, 112)
        Me.txtPLZ.Name = "txtPLZ"
        Me.txtPLZ.Size = New System.Drawing.Size(367, 20)
        Me.txtPLZ.TabIndex = 3
        '
        'txtOrt
        '
        Me.txtOrt.Location = New System.Drawing.Point(97, 138)
        Me.txtOrt.Name = "txtOrt"
        Me.txtOrt.Size = New System.Drawing.Size(367, 20)
        Me.txtOrt.TabIndex = 4
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(34, 115)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(27, 13)
        Me.Label3.TabIndex = 42
        Me.Label3.Text = "PLZ"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(34, 141)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(21, 13)
        Me.Label4.TabIndex = 43
        Me.Label4.Text = "Ort"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(34, 88)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(45, 13)
        Me.Label5.TabIndex = 44
        Me.Label5.Text = "Adresse"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(34, 195)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(38, 13)
        Me.Label6.TabIndex = 45
        Me.Label6.Text = "Nation"
        '
        'txtGeburtsdatum
        '
        Me.txtGeburtsdatum.Location = New System.Drawing.Point(97, 219)
        Me.txtGeburtsdatum.Name = "txtGeburtsdatum"
        Me.txtGeburtsdatum.Size = New System.Drawing.Size(367, 20)
        Me.txtGeburtsdatum.TabIndex = 7
        '
        'txtEMailAdresse
        '
        Me.txtEMailAdresse.Location = New System.Drawing.Point(97, 247)
        Me.txtEMailAdresse.Name = "txtEMailAdresse"
        Me.txtEMailAdresse.Size = New System.Drawing.Size(367, 20)
        Me.txtEMailAdresse.TabIndex = 8
        '
        'Label8
        '
        Me.Label8.AutoSize = True
        Me.Label8.Location = New System.Drawing.Point(34, 250)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(36, 13)
        Me.Label8.TabIndex = 49
        Me.Label8.Text = "E-Mail"
        '
        'Label9
        '
        Me.Label9.AutoSize = True
        Me.Label9.Location = New System.Drawing.Point(34, 276)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(46, 13)
        Me.Label9.TabIndex = 50
        Me.Label9.Text = "Ausweis"
        '
        'txtAusweisNummer
        '
        Me.txtAusweisNummer.Location = New System.Drawing.Point(97, 273)
        Me.txtAusweisNummer.Name = "txtAusweisNummer"
        Me.txtAusweisNummer.Size = New System.Drawing.Size(273, 20)
        Me.txtAusweisNummer.TabIndex = 9
        '
        'txtId
        '
        Me.txtId.AutoSize = True
        Me.txtId.Location = New System.Drawing.Point(376, 274)
        Me.txtId.Name = "txtId"
        Me.txtId.Size = New System.Drawing.Size(34, 17)
        Me.txtId.TabIndex = 10
        Me.txtId.TabStop = True
        Me.txtId.Text = "Id"
        Me.txtId.UseVisualStyleBackColor = True
        '
        'txtPass
        '
        Me.txtPass.AutoSize = True
        Me.txtPass.Location = New System.Drawing.Point(419, 274)
        Me.txtPass.Name = "txtPass"
        Me.txtPass.Size = New System.Drawing.Size(48, 17)
        Me.txtPass.TabIndex = 11
        Me.txtPass.TabStop = True
        Me.txtPass.Text = "Pass"
        Me.txtPass.UseVisualStyleBackColor = True
        '
        'Label15
        '
        Me.Label15.AutoSize = True
        Me.Label15.Location = New System.Drawing.Point(34, 222)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(61, 13)
        Me.Label15.TabIndex = 54
        Me.Label15.Text = "Geb.Datum"
        '
        'btnSpeichern
        '
        Me.btnSpeichern.Location = New System.Drawing.Point(502, 266)
        Me.btnSpeichern.Name = "btnSpeichern"
        Me.btnSpeichern.Size = New System.Drawing.Size(75, 23)
        Me.btnSpeichern.TabIndex = 12
        Me.btnSpeichern.Text = "Speichern"
        Me.btnSpeichern.UseVisualStyleBackColor = True
        '
        'btnAbbrechen
        '
        Me.btnAbbrechen.Location = New System.Drawing.Point(583, 266)
        Me.btnAbbrechen.Name = "btnAbbrechen"
        Me.btnAbbrechen.Size = New System.Drawing.Size(75, 23)
        Me.btnAbbrechen.TabIndex = 13
        Me.btnAbbrechen.Text = "Abbrechen"
        Me.btnAbbrechen.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(664, 266)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(156, 23)
        Me.Button1.TabIndex = 14
        Me.Button1.Text = "Meldeschein und Anmeldung drucken"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Location = New System.Drawing.Point(18, 19)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(33, 17)
        Me.RadioButton1.TabIndex = 56
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "D"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(57, 19)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(32, 17)
        Me.RadioButton2.TabIndex = 57
        Me.RadioButton2.TabStop = True
        Me.RadioButton2.Text = "E"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'RadioButton3
        '
        Me.RadioButton3.AutoSize = True
        Me.RadioButton3.Location = New System.Drawing.Point(95, 19)
        Me.RadioButton3.Name = "RadioButton3"
        Me.RadioButton3.Size = New System.Drawing.Size(31, 17)
        Me.RadioButton3.TabIndex = 58
        Me.RadioButton3.TabStop = True
        Me.RadioButton3.Text = "F"
        Me.RadioButton3.UseVisualStyleBackColor = True
        '
        'RadioButton4
        '
        Me.RadioButton4.AutoSize = True
        Me.RadioButton4.Location = New System.Drawing.Point(132, 19)
        Me.RadioButton4.Name = "RadioButton4"
        Me.RadioButton4.Size = New System.Drawing.Size(28, 17)
        Me.RadioButton4.TabIndex = 59
        Me.RadioButton4.TabStop = True
        Me.RadioButton4.Text = "I"
        Me.RadioButton4.UseVisualStyleBackColor = True
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.RadioButton1)
        Me.GroupBox1.Controls.Add(Me.RadioButton4)
        Me.GroupBox1.Controls.Add(Me.RadioButton2)
        Me.GroupBox1.Controls.Add(Me.RadioButton3)
        Me.GroupBox1.Location = New System.Drawing.Point(552, 192)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(169, 50)
        Me.GroupBox1.TabIndex = 60
        Me.GroupBox1.TabStop = False
        '
        'Label7
        '
        Me.Label7.AutoSize = True
        Me.Label7.Location = New System.Drawing.Point(499, 213)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(47, 13)
        Me.Label7.TabIndex = 61
        Me.Label7.Text = "Sprache"
        '
        'cmbNation
        '
        Me.cmbNation.FormattingEnabled = True
        Me.cmbNation.Location = New System.Drawing.Point(97, 192)
        Me.cmbNation.MaxDropDownItems = 15
        Me.cmbNation.Name = "cmbNation"
        Me.cmbNation.Size = New System.Drawing.Size(367, 21)
        Me.cmbNation.TabIndex = 6
        '
        'cmbLand
        '
        Me.cmbLand.FormattingEnabled = True
        Me.cmbLand.Location = New System.Drawing.Point(97, 164)
        Me.cmbLand.MaxDropDownItems = 15
        Me.cmbLand.Name = "cmbLand"
        Me.cmbLand.Size = New System.Drawing.Size(367, 21)
        Me.cmbLand.TabIndex = 5
        '
        'Label10
        '
        Me.Label10.AutoSize = True
        Me.Label10.Location = New System.Drawing.Point(34, 167)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(31, 13)
        Me.Label10.TabIndex = 63
        Me.Label10.Text = "Land"
        '
        'AdressenErfassen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(841, 411)
        Me.Controls.Add(Me.cmbLand)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.cmbNation)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.btnAbbrechen)
        Me.Controls.Add(Me.btnSpeichern)
        Me.Controls.Add(Me.txtNachName)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtVorname)
        Me.Controls.Add(Me.txtAdresse)
        Me.Controls.Add(Me.txtPLZ)
        Me.Controls.Add(Me.txtOrt)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.txtGeburtsdatum)
        Me.Controls.Add(Me.txtEMailAdresse)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.txtAusweisNummer)
        Me.Controls.Add(Me.txtId)
        Me.Controls.Add(Me.txtPass)
        Me.Controls.Add(Me.Label15)
        Me.Name = "AdressenErfassen"
        Me.Text = "AdressenErfassen"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtNachName As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents txtVorname As System.Windows.Forms.TextBox
    Friend WithEvents txtAdresse As System.Windows.Forms.TextBox
    Friend WithEvents txtPLZ As System.Windows.Forms.TextBox
    Friend WithEvents txtOrt As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents txtGeburtsdatum As System.Windows.Forms.TextBox
    Friend WithEvents txtEMailAdresse As System.Windows.Forms.TextBox
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents txtAusweisNummer As System.Windows.Forms.TextBox
    Friend WithEvents txtId As System.Windows.Forms.RadioButton
    Friend WithEvents txtPass As System.Windows.Forms.RadioButton
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents btnSpeichern As System.Windows.Forms.Button
    Friend WithEvents btnAbbrechen As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button
    Friend WithEvents PrintDocument1 As System.Drawing.Printing.PrintDocument
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton3 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton4 As System.Windows.Forms.RadioButton
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents cmbNation As System.Windows.Forms.ComboBox
    Friend WithEvents PrintDocument2 As System.Drawing.Printing.PrintDocument
    Friend WithEvents cmbLand As System.Windows.Forms.ComboBox
    Friend WithEvents Label10 As System.Windows.Forms.Label
End Class
