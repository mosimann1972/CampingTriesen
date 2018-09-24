
Imports System.Data.OleDb
Imports System.Console
Imports System.Drawing.Printing
Imports System.IO
Imports System.Math
Imports Microsoft.Office.Interop

Public Class Admin

Dim conn As New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=camping.mdb")

Private Sub Admin_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load



End Sub


Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click


		Me.Cursor = Cursors.WaitCursor
		Button1.Enabled = False


		Dim AnzahlTage As Long = DateDiff(DateInterval.Day, _
										DateTimePicker1.Value.Date, _
										DateTimePicker2.Value.Date, _
										FirstDayOfWeek.Monday, _
										FirstWeekOfYear.Jan1)

		Dim x As String
		Dim d As Date
		Dim Jahr As Integer


		d = DateTimePicker1.Value
		Jahr = d.Year

		For i = 0 To AnzahlTage

			x = "#" & d.Month & "/" & d.Day & "/" & d.Year & "#"

			Dim cmd As New OleDbCommand("Insert Into tbZimmerGebucht (Datum,Jahr, Moddate) Values(" & x & "," & Jahr & ",Now())", conn)

			Dim da As New OleDbDataAdapter(cmd)
			Dim ds As New DataSet()

			Try
				da.Fill(ds, "tbZimmerGebucht")
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try

			d = d.AddDays(1)

		Next

		Me.Cursor = Cursors.Default

		MsgBox("Daten erfasst", MsgBoxStyle.Exclamation)

		Button1.Enabled = True

End Sub


Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles Button2.Click

		Me.Cursor = Cursors.WaitCursor
		Button2.Enabled = False

		Dim d As Date
		Dim Jahr As Integer

		d = Now()
		Jahr = d.Year


		Dim cmd As New OleDbCommand("Delete * From tbZimmerGebucht Where Jahr  = " & Jahr & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet("tbZimmerGebucht")

		If MsgBox("Achtung: Daten werden gelöscht und können nicht wieder hergestellt werden!", MsgBoxStyle.YesNoCancel + MsgBoxStyle.Critical) = MsgBoxResult.Yes Then
			Try
				da.Fill(ds, "tbZimmerGebucht")
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try

			MsgBox("Daten gelöscht", MsgBoxStyle.Exclamation)

			Button2.Enabled = True
			Me.Cursor = Cursors.Default

		Else
			Button2.Enabled = True
			Me.Cursor = Cursors.Default
			Exit Sub
		End If


End Sub

Private Sub Button3_Click(sender As System.Object, e As System.EventArgs) Handles Button3.Click
		Me.Close()
End Sub
End Class