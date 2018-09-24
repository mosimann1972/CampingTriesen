

Imports System.Data.OleDb
Imports System.Console
Imports System.Drawing.Printing
Imports System.IO
Imports System.Math
Imports Microsoft.Office.Interop

Public Class Belegung

	Public KlickDatum As Date
	Public First As Boolean = True
	Public AnzahlTage As Integer

	Dim conn As New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=camping.mdb")

Private Sub Belegung_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

	Button3.Enabled = False
	Button4.Enabled = False

	Start(Now())

End Sub

Private Sub Start(Datum As Date)

		AlleGruen()

		Dim d As Date
		Dim SuchDatum As String
		Dim PrintDatum As String
		Dim i As Integer
		d = Datum

		For i = 1 To 7

			SuchDatum = "#" & d.Month & "/" & d.Day & "/" & d.Year & "#"
			PrintDatum = d.Day & "." & d.Month & "." & d.Year
			Select_tbZimmer(SuchDatum, PrintDatum, i)
			d = d.AddDays(1)

		Next

End Sub

Public Sub Select_tbZimmer(SuchDatum As String, PrintDatum As String, Tag As Integer)

		Dim cmd As New OleDbCommand("SELECT * from tbZimmerGebucht WHERE Datum = " & SuchDatum & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet()

		Try
			da.Fill(ds, "tbZimmerGebucht")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try

		Me.Controls("Datum" & Tag.ToString).Text = PrintDatum


		If ds.Tables("tbZimmerGebucht").Rows.Count = 0 Then
			Exit Sub
		End If

		For ii = 0 To ds.Tables("tbZimmerGebucht").Rows.Count - 1

			Dim CountRun As Integer

			Select Case Tag
			Case 1
				CountRun = 0
			Case 2
				CountRun = 23
			Case 3
				CountRun = 46
			Case 4
				CountRun = 69
			Case 5
				CountRun = 92
			Case 6
				CountRun = 115
			Case 7
				CountRun = 138
			End Select


			For i = 3 To 25

				CountRun = CountRun + 1

				If ds.Tables("tbZimmerGebucht").Rows(ii).Item(i) = True Then
					Me.Controls("lbl" & CountRun.ToString).BackColor = Color.Red
					'Me.Controls("lbl" & CountRun.ToString).Text

				End If

			Next

		Next


	End Sub

	Private Sub AlleGruen()

	For i = 1 To 161

		Me.Controls("lbl" & i.ToString).BackColor = Color.Green
		Me.Controls("lbl" & i.ToString).ForeColor = Color.White

	Next

	End Sub


	Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click

		Me.Close()

	End Sub

Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

	Button3.Enabled = True
	Button4.Enabled = True

	If First = True Then
		KlickDatum = Now()
		First = False
		AnzahlTage = 7
	Else
		AnzahlTage = AnzahlTage + 7
	End If

	Start(KlickDatum.AddDays(AnzahlTage))

End Sub

Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click

	If First = False Then
		AnzahlTage = AnzahlTage - 7
	End If

	Start(KlickDatum.AddDays(AnzahlTage))

End Sub

Private Sub Button4_Click(sender As System.Object, e As System.EventArgs) Handles Button4.Click

	AnzahlTage = 0

	Start(Now())

End Sub


End Class