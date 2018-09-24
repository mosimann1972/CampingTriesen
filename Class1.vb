
Imports System.Data.OleDb
Imports System.Console
Imports System.IO

Public Class Class1


    Dim conn As New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=camping.mdb")

	Public Sub StartLoeschenCamping(ByVal i As Integer)


		If MsgBox("Wollen Sie wirklich löschen?", vbOKCancel + MsgBoxStyle.Critical) = vbCancel Then
			Exit Sub
		End If

		Dim cmd As New OleDbCommand("Delete * From tbBooking Where BookingId = " & i & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet("tbBooking")

		Try
			da.Fill(ds, "tbBooking")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try


		Dim cmd1 As New OleDbCommand("Delete * From tbBookingtbAdress Where BookingId = " & i & "", conn)
		Dim da1 As New OleDbDataAdapter(cmd1)
		Dim ds1 As New DataSet("tbBookingtbAdress")

		Try
			da1.Fill(ds, "tbBookingtbAdress")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try

		MsgBox("Buchung gelöscht", MsgBoxStyle.Exclamation)


	End Sub


Public Sub StartLoeschenZimmer(ByVal i As Integer)


		If MsgBox("Wollen Sie wirklich löschen?", vbOKCancel + MsgBoxStyle.Critical) = vbCancel Then
			Exit Sub
		End If

		Dim cmd As New OleDbCommand("Delete * From tbZimmer Where BookingId = " & i & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet("tbZimmer")

		Try
			da.Fill(ds, "tbZimmer")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try


		Dim cmd1 As New OleDbCommand("Delete * From tbZimmertbAdress Where BookingId = " & i & "", conn)
		Dim da1 As New OleDbDataAdapter(cmd1)
		Dim ds1 As New DataSet("tbZimmertbAdress")

		Try
			da1.Fill(ds, "tbZimmertbAdress")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try


		CheckIn02.Delete_ZimmerBuchungen_Final(i)

		MsgBox("Zimmer gelöscht", MsgBoxStyle.Exclamation)

End Sub


End Class
