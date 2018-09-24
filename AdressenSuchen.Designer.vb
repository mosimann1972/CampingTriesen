<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class AdressenSuchen
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
        Me.txtSearchString = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btnLoeschen = New System.Windows.Forms.Button()
        Me.btnSuchen = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.chkBegleitpersonen = New System.Windows.Forms.CheckBox()
        Me.chkAktiv = New System.Windows.Forms.CheckBox()
        Me.lblSuchTitel = New System.Windows.Forms.Label()
        Me.txtBookingId = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtPlatznummer = New System.Windows.Forms.TextBox()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'txtSearchString
        '
        Me.txtSearchString.Location = New System.Drawing.Point(112, 36)
        Me.txtSearchString.Name = "txtSearchString"
        Me.txtSearchString.Size = New System.Drawing.Size(304, 20)
        Me.txtSearchString.TabIndex = 0
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(37, 39)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(38, 13)
        Me.Label1.TabIndex = 38
        Me.Label1.Text = "Suche"
        '
        'btnLoeschen
        '
        Me.btnLoeschen.Location = New System.Drawing.Point(699, 39)
        Me.btnLoeschen.Name = "btnLoeschen"
        Me.btnLoeschen.Size = New System.Drawing.Size(75, 23)
        Me.btnLoeschen.TabIndex = 4
        Me.btnLoeschen.Text = "Abbrechen"
        Me.btnLoeschen.UseVisualStyleBackColor = True
        '
        'btnSuchen
        '
        Me.btnSuchen.Location = New System.Drawing.Point(618, 39)
        Me.btnSuchen.Name = "btnSuchen"
        Me.btnSuchen.Size = New System.Drawing.Size(75, 23)
        Me.btnSuchen.TabIndex = 3
        Me.btnSuchen.Text = "Suchen"
        Me.btnSuchen.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(40, 97)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(748, 574)
        Me.DataGridView1.TabIndex = 41
        '
        'chkBegleitpersonen
        '
        Me.chkBegleitpersonen.AutoSize = True
        Me.chkBegleitpersonen.Location = New System.Drawing.Point(427, 39)
        Me.chkBegleitpersonen.Name = "chkBegleitpersonen"
        Me.chkBegleitpersonen.Size = New System.Drawing.Size(102, 17)
        Me.chkBegleitpersonen.TabIndex = 42
        Me.chkBegleitpersonen.Text = "Begleitpersonen"
        Me.chkBegleitpersonen.UseVisualStyleBackColor = True
        '
        'chkAktiv
        '
        Me.chkAktiv.AutoSize = True
        Me.chkAktiv.Location = New System.Drawing.Point(535, 39)
        Me.chkAktiv.Name = "chkAktiv"
        Me.chkAktiv.Size = New System.Drawing.Size(77, 17)
        Me.chkAktiv.TabIndex = 43
        Me.chkAktiv.Text = "verrechnet"
        Me.chkAktiv.UseVisualStyleBackColor = True
        '
        'lblSuchTitel
        '
        Me.lblSuchTitel.AutoSize = True
        Me.lblSuchTitel.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSuchTitel.Location = New System.Drawing.Point(36, 9)
        Me.lblSuchTitel.Name = "lblSuchTitel"
        Me.lblSuchTitel.Size = New System.Drawing.Size(60, 20)
        Me.lblSuchTitel.TabIndex = 127
        Me.lblSuchTitel.Text = "Suche"
        '
        'txtBookingId
        '
        Me.txtBookingId.Location = New System.Drawing.Point(112, 62)
        Me.txtBookingId.Name = "txtBookingId"
        Me.txtBookingId.Size = New System.Drawing.Size(103, 20)
        Me.txtBookingId.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(37, 65)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(69, 13)
        Me.Label2.TabIndex = 129
        Me.Label2.Text = "BuchungsNr."
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(240, 65)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(67, 13)
        Me.Label3.TabIndex = 130
        Me.Label3.Text = "Platznummer"
        '
        'txtPlatznummer
        '
        Me.txtPlatznummer.Location = New System.Drawing.Point(313, 62)
        Me.txtPlatznummer.Name = "txtPlatznummer"
        Me.txtPlatznummer.Size = New System.Drawing.Size(103, 20)
        Me.txtPlatznummer.TabIndex = 2
        '
        'AdressenSuchen
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(866, 706)
        Me.Controls.Add(Me.txtPlatznummer)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.txtBookingId)
        Me.Controls.Add(Me.lblSuchTitel)
        Me.Controls.Add(Me.chkAktiv)
        Me.Controls.Add(Me.chkBegleitpersonen)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.btnLoeschen)
        Me.Controls.Add(Me.btnSuchen)
        Me.Controls.Add(Me.txtSearchString)
        Me.Controls.Add(Me.Label1)
        Me.Name = "AdressenSuchen"
        Me.Text = "AdressenSuchen"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents txtSearchString As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnLoeschen As System.Windows.Forms.Button
    Friend WithEvents btnSuchen As System.Windows.Forms.Button
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents chkBegleitpersonen As System.Windows.Forms.CheckBox
    Friend WithEvents chkAktiv As System.Windows.Forms.CheckBox
    Friend WithEvents lblSuchTitel As System.Windows.Forms.Label
    Friend WithEvents txtBookingId As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents txtPlatznummer As System.Windows.Forms.TextBox
End Class
