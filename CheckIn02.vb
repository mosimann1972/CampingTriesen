
Imports System.Data.OleDb
Imports System.Console
Imports System.Drawing.Printing
Imports System.IO
Imports System.Math
Imports Microsoft.Office.Interop

Public Class CheckIn02

	Public Public_AdressId As Integer
	Public Public_BookingId As Integer
	Public StartFlag As String = ""
	Public DokumentTitel As String = ""
	Public BungalowCheck As Boolean
	Public SprachCode As String

	Public LandErfasst1 As String
	Public LandErfasst2 As String
	Public LandErfasst3 As String
	Public LandErfasst4 As String
	Public LandErfasst5 As String
	Public LandErfasst6 As String
	Public LandErfasst7 As String
	Public LandErfasst8 As String
	Public LandErfasst9 As String
	Public LandErfasst10 As String

	Public Preis_Bett_4er As Double
	Public Preis_Bett_6er As Double

	Public Preis_Zimmer1_Komplett As Double
	Public Preis_Zimmer2_Komplett As Double
	Public Preis_Zimmer3_Komplett As Double

	Public Preis_Zimmer1_Einzel As Double
	Public Preis_Zimmer2_Einzel As Double
	Public Preis_Zimmer3_Einzel As Double

	Public Preis_Zimmer1_Total As Double
	Public Preis_Zimmer2_Total As Double
	Public Preis_Zimmer3_Total As Double

	Public Preis_Zimmer1_Kinderbett As Double
	Public Preis_Zimmer2_Kinderbett As Double
	Public Preis_Zimmer3_Kinderbett As Double

	Public CheckZimmerBuchungen As Boolean

	Dim conn As New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=camping.mdb")

Private Sub CheckIn02_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

		LaenderLaden()

		PreisLaden()

End Sub

Private Sub LaenderLaden()

	Dim dt1 As DataTable
		dt1 = Form1.Laenderliste()
		cmbLand1.DataSource = dt1
		cmbLand1.DisplayMember = "Land"
		If LandErfasst1 <> "" Then
			cmbLand1.Text = LandErfasst1.ToString
		End If


		Dim dt2 As DataTable
		dt2 = Form1.Laenderliste()
		cmbLand2.DataSource = dt2
		cmbLand2.DisplayMember = "Land"
		If LandErfasst2 <> "" Then
			cmbLand2.Text = LandErfasst2.ToString
		End If

		Dim dt3 As DataTable
		dt3 = Form1.Laenderliste()
		cmbLand3.DataSource = dt3
		cmbLand3.DisplayMember = "Land"
		If LandErfasst3 <> "" Then
			cmbLand3.Text = LandErfasst3.ToString
		End If

		Dim dt4 As DataTable
		dt4 = Form1.Laenderliste()
		cmbLand4.DataSource = dt4
		cmbLand4.DisplayMember = "Land"
		If LandErfasst4 <> "" Then
			cmbLand4.Text = LandErfasst4.ToString
		End If

		Dim dt5 As DataTable
		dt5 = Form1.Laenderliste()
		cmbLand5.DataSource = dt5
		cmbLand5.DisplayMember = "Land"
		If LandErfasst5 <> "" Then
			cmbLand5.Text = LandErfasst5.ToString
		End If

		Dim dt6 As DataTable
		dt6 = Form1.Laenderliste()
		cmbLand6.DataSource = dt6
		cmbLand6.DisplayMember = "Land"
		If LandErfasst6 <> "" Then
			cmbLand6.Text = LandErfasst6.ToString
		End If

		Dim dt7 As DataTable
		dt7 = Form1.Laenderliste()
		cmbLand7.DataSource = dt7
		cmbLand7.DisplayMember = "Land"
		If LandErfasst7 <> "" Then
			cmbLand7.Text = LandErfasst7.ToString
		End If

		Dim dt8 As DataTable
		dt8 = Form1.Laenderliste()
		cmbLand8.DataSource = dt8
		cmbLand8.DisplayMember = "Land"
		If LandErfasst8 <> "" Then
			cmbLand8.Text = LandErfasst8.ToString
		End If

		Dim dt9 As DataTable
		dt9 = Form1.Laenderliste()
		cmbLand9.DataSource = dt9
		cmbLand9.DisplayMember = "Land"
		If LandErfasst9 <> "" Then
			cmbLand9.Text = LandErfasst9.ToString
		End If

		Dim dt10 As DataTable
		dt10 = Form1.Laenderliste()
		cmbLand10.DataSource = dt10
		cmbLand10.DisplayMember = "Land"
		If LandErfasst10 <> "" Then
			cmbLand10.Text = LandErfasst10.ToString
		End If


		LandErfasst1 = ""
		LandErfasst2 = ""
		LandErfasst3 = ""
		LandErfasst4 = ""
		LandErfasst5 = ""
		LandErfasst6 = ""
		LandErfasst7 = ""
		LandErfasst8 = ""
		LandErfasst9 = ""
		LandErfasst10 = ""



End Sub


Public Sub StartBooking(ByVal i As Integer)

		StartFlag = "StartBooking"

		Public_AdressId = i

		Dim cmd As New OleDbCommand("Select * from tbAdress where AdressId = " & Public_AdressId & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet()

		da.Fill(ds, "tbAdress")

		lblNameKunde.Text = ds.Tables("tbAdress").Rows(0).Item(1) & " " & ds.Tables("tbAdress").Rows(0).Item(2)

		DateTimePicker2_DatumNeu()


		If lblBookingId.Text = "---" Then
			btnBestaetigungMitMeldeschein.Enabled = False
			btnBestaetigungOhneMeldeschein.Enabled = False
			btnQuittung.Enabled = False

		Else
			btnBestaetigungMitMeldeschein.Enabled = True
			btnBestaetigungOhneMeldeschein.Enabled = True
			btnQuittung.Enabled = True

		End If


	End Sub

Private Sub SpeichernBookingDaten()


		Check_ZimmerBuchungen()

		If CheckZimmerBuchungen = True Then
			Exit Sub
		End If

		Delete_ZimmerBuchungen()


		Dim strRabattArt As String = ""

		If txtRabattProzent.Checked = True Then
			strRabattArt = "%"
		End If

		If txtRabattCHF.Checked = True Then
			strRabattArt = "CHF"
		End If

		InitCMB()

		Dim cmd As New OleDbCommand("Insert Into tbZimmer (Anreise, Abreise, AnzahlNaechte,Aktiv,Erwachsene,Kinder" _
			& ",B11,B12,B13,B14,Z1Z,Z1E,Z1K,B21,B22,B23,B24,B25,B26,Z2Z,Z2E,Z2K,B31,B32,B33,B34,Z3Z,Z3E,Z3K,ParkplatzMotorrad,ParkplatzAuto,Handtuecher,Fruehstueck,Kueche,AutoKennZeichen,Bemerkungen,RabattWert,RabattArt,SonstigesText,SonstigesBetrag)" _
				& "Values('" & DateTimePicker1.Value.Date & "'" _
				& ", '" & DateTimePicker2.Value.Date & "'" _
				& ", '" & txtNaechte.Text & "'" _
				& ", " & True & "" _
				& ", '" & cmb01.Text & "'" _
				& ", '" & cmb02.Text & "'" _
				& ", '" & chkBett11.CheckState & "'" _
				& ", '" & chkBett12.CheckState & "'" _
				& ", '" & chkBett13.CheckState & "'" _
				& ", '" & chkBett14.CheckState & "'" _
				& ", '" & chkZimmer1Komplett.CheckState & "'" _
				& ", '" & chkZimmer1Einzel.CheckState & "'" _
				& ", '" & chkZimmer1Kinderbett.CheckState & "'" _
				& ", '" & chkBett21.CheckState & "'" _
				& ", '" & chkBett22.CheckState & "'" _
				& ", '" & chkBett23.CheckState & "'" _
				& ", '" & chkBett24.CheckState & "'" _
				& ", '" & chkBett25.CheckState & "'" _
				& ", '" & chkBett26.CheckState & "'" _
				& ", '" & chkZimmer2Komplett.CheckState & "'" _
				& ", '" & chkZimmer2Einzel.CheckState & "'" _
				& ", '" & chkZimmer2Kinderbett.CheckState & "'" _
				& ", '" & chkBett31.CheckState & "'" _
				& ", '" & chkBett32.CheckState & "'" _
				& ", '" & chkBett33.CheckState & "'" _
				& ", '" & chkBett34.CheckState & "'" _
				& ", '" & chkZimmer3Komplett.CheckState & "'" _
				& ", '" & chkZimmer3Einzel.CheckState & "'" _
				& ", '" & chkZimmer3Kinderbett.CheckState & "'" _
				& ", '" & cmb03.Text & "'" _
				& ", '" & cmb04.Text & "'" _
				& ", '" & cmb05.Text & "'" _
				& ", '" & cmb06.Text & "'" _
				& ", '" & cmb07.Text & "'" _
				& ", '" & txtAutoKennZeichen.Text & "'" _
				& ", '" & txtBemerkungen.Text & "','" & txtRabatt.Text & "','" & strRabattArt.ToString & "','" & txtSonstigesText.Text & "','" & txtSonstigesBetrag.Text & "')", conn)

		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet("tbBooking")

		Try
			da.Fill(ds, "tbBooking")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try


		'---------------------------------

		Dim cmd1 As New OleDbCommand("Select * from tbZimmer Order by BookingId ASC", conn)
		Dim da1 As New OleDbDataAdapter(cmd1)
		Dim ds1 As New DataSet()

		Dim d, e As Integer

		da1.Fill(ds1, "tbZimmer1")

		d = ds1.Tables("tbZimmer1").Rows.Count
		e = ds1.Tables("tbZimmer1").Rows(d - 1).Item(0)

		Public_BookingId = e

		'---------------------------------

		Dim cmd2 As New OleDbCommand("Insert Into tbZimmertbAdress (BookingId, AdressId, Leader, Moddate) Values(" & e & ", " & Public_AdressId & ",true,Now())", conn)
		Dim da2 As New OleDbDataAdapter(cmd2)
		Dim ds2 As New DataSet()

		Try
			da2.Fill(ds2, "tbZimmertbAdress")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try

	End Sub

	Private Sub InitCMB()

		If cmb01.Text = "" Then cmb01.Text = 0
		If cmb02.Text = "" Then cmb02.Text = 0
		If cmb03.Text = "" Then cmb03.Text = 0
		If cmb04.Text = "" Then cmb04.Text = 0
		If cmb05.Text = "" Then cmb05.Text = 0
		If cmb06.Text = "" Then cmb06.Text = 0
		If cmb07.Text = "" Then cmb07.Text = 0


	End Sub

Public Sub Delete_ZimmerBuchungen()


		Dim cmd As New OleDbCommand("Delete * From tbZimmerGebucht Where BookingId = " & Public_BookingId & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet("tbZimmerGebucht")

		Try
			da.Fill(ds, "tbZimmerGebucht")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try


	End Sub


Public Sub Delete_ZimmerBuchungen_Final(BookingId As Integer)


		Dim cmd As New OleDbCommand("Delete * From tbZimmerGebucht Where BookingId = " & BookingId & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet("tbZimmerGebucht")

		Try
			da.Fill(ds, "tbZimmerGebucht")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try


	End Sub


Public Sub Check_ZimmerBuchungen()


		CheckZimmerBuchungen = False

		Dim CheckFound As Boolean

		Dim d As Date
		d = DateTimePicker1.Value

		Dim AnzahlTage As Long = DateDiff(DateInterval.Day, _
									DateTimePicker1.Value.Date, _
									DateTimePicker2.Value.Date, _
									FirstDayOfWeek.Monday, _
									FirstWeekOfYear.Jan1)

		Dim x As String

		For i = 0 To AnzahlTage - 1

			CheckFound = True

			x = "#" & d.Month & "/" & d.Day & "/" & d.Year & "#"

			Dim cmd As New OleDbCommand("Select * from tbZimmerGebucht Where Datum = " & x & " And BookingId <> " & Public_BookingId & "", conn)
			Dim da As New OleDbDataAdapter(cmd)
			Dim ds As New DataSet()

			Try
				da.Fill(ds, "tbZimmerGebucht")
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try

			If ds.Tables("tbZimmerGebucht").Rows.Count = 0 Then
				CheckFound = False
			End If


			'Debug.WriteLine(ds.Tables("tbZimmerGebucht").Rows.Count)


			'Wenn was gefunden, dann check
			If CheckFound = True Then



				For xx = 0 To ds.Tables("tbZimmerGebucht").Rows.Count - 1

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(3) = True And chkBett11.Checked = True Then
						MsgBox("Bett 11 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(4) = True And chkBett12.Checked = True Then
						MsgBox("Bett 12 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(5) = True And chkBett13.Checked = True Then
						MsgBox("Bett 13 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(6) = True And chkBett14.Checked = True Then
						MsgBox("Bett 14 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(7) = True And chkZimmer1Komplett.Checked = True Then
						MsgBox("Zimmer 1 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(8) = True And chkZimmer1Einzel.Checked = True Then
						MsgBox("Zimmer 1 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(9) = True And chkZimmer1Kinderbett.Checked = True Then
						MsgBox("Zimmer 1 Kinderbett zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If




					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(10) = True And chkBett21.Checked = True Then
						MsgBox("Bett 21 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(11) = True And chkBett22.Checked = True Then
						MsgBox("Bett 22 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(12) = True And chkBett23.Checked = True Then
						MsgBox("Bett 23 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(13) = True And chkBett24.Checked = True Then
						MsgBox("Bett 24 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(14) = True And chkBett25.Checked = True Then
						MsgBox("Bett 25 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(15) = True And chkBett26.Checked = True Then
						MsgBox("Bett 26 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(17) = True And chkZimmer2Komplett.Checked = True Then
						MsgBox("Zimmer 2 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(18) = True And chkZimmer2Einzel.Checked = True Then
						MsgBox("Zimmer 2 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If


					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(19) = True And chkZimmer2Kinderbett.Checked = True Then
						MsgBox("Zimmer 2 Kinderbett zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If





					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(20) = True And chkBett31.Checked = True Then
						MsgBox("Bett 31 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(21) = True And chkBett32.Checked = True Then
						MsgBox("Bett 32 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(22) = True And chkBett33.Checked = True Then
						MsgBox("Bett 33 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(23) = True And chkBett34.Checked = True Then
						MsgBox("Bett 34 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(24) = True And chkZimmer3Komplett.Checked = True Then
						MsgBox("Zimmer 3 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(25) = True And chkZimmer3Einzel.Checked = True Then
						MsgBox("Zimmer 3 zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

					If ds.Tables("tbZimmerGebucht").Rows(xx).Item(25) = True And chkZimmer3Kinderbett.Checked = True Then
						MsgBox("Zimmer 3 Kinderbett zum Zeitpunkt schon reserviert", MsgBoxStyle.OkCancel + MsgBoxStyle.Critical)
						CheckZimmerBuchungen = True
						Exit Sub
					End If

				Next

			End If

			d = d.AddDays(1)

		Next

End Sub


	Public Sub Update_ZimmerBuchungen()

		Dim d As Date
		d = DateTimePicker1.Value

		Dim AnzahlTage As Long = DateDiff(DateInterval.Day, _
									DateTimePicker1.Value.Date, _
									DateTimePicker2.Value.Date, _
									FirstDayOfWeek.Monday, _
									FirstWeekOfYear.Jan1)

		Dim x As String

		For i = 0 To AnzahlTage - 1

			x = "#" & d.Month & "/" & d.Day & "/" & d.Year & "#"

			Dim cmd As New OleDbCommand("Insert Into tbZimmerGebucht (Datum,Jahr,BookingId" _
			& ",B11,B12,B13,B14,Z1Z,Z1E,Z1K,B21,B22,B23,B24,B25,B26,Z2Z,Z2E,Z2K,B31,B32,B33,B34,Z3Z,Z3E,Z3K)" _
				& "Values(" & x & "" _
				& ", " & d.Year & "" _
				& ", " & Public_BookingId & "" _
				& ", '" & chkBett11.CheckState & "'" _
				& ", '" & chkBett12.CheckState & "'" _
				& ", '" & chkBett13.CheckState & "'" _
				& ", '" & chkBett14.CheckState & "'" _
				& ", '" & chkZimmer1Komplett.CheckState & "'" _
				& ", '" & chkZimmer1Einzel.CheckState & "'" _
				& ", '" & chkZimmer1Kinderbett.CheckState & "'" _
				& ", '" & chkBett21.CheckState & "'" _
				& ", '" & chkBett22.CheckState & "'" _
				& ", '" & chkBett23.CheckState & "'" _
				& ", '" & chkBett24.CheckState & "'" _
				& ", '" & chkBett25.CheckState & "'" _
				& ", '" & chkBett26.CheckState & "'" _
				& ", '" & chkZimmer2Komplett.CheckState & "'" _
				& ", '" & chkZimmer2Einzel.CheckState & "'" _
				& ", '" & chkZimmer2Kinderbett.CheckState & "'" _
				& ", '" & chkBett31.CheckState & "'" _
				& ", '" & chkBett32.CheckState & "'" _
				& ", '" & chkBett33.CheckState & "'" _
				& ", '" & chkBett34.CheckState & "'" _
				& ", '" & chkZimmer3Komplett.CheckState & "'" _
				& ", '" & chkZimmer3Einzel.CheckState & "'" _
				& ", '" & chkZimmer3Kinderbett.CheckState & "')", conn)

			Dim da As New OleDbDataAdapter(cmd)
			Dim ds As New DataSet("tbZimmerGebucht")

			Try
				da.Fill(ds, "tbZimmerGebucht")
			Catch ex As Exception
				MessageBox.Show(ex.Message)
			End Try

			d = d.AddDays(1)


		Next

	End Sub

	Private Sub PreisLaden()

		Dim cmd As New OleDbCommand("Select * from tbPreis", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet()

		da.Fill(ds, "tbPreis")

		txtTotalZimmer1.Text = String.Format("{0:N}", 0)
		txtTotalZimmer2.Text = String.Format("{0:N}", 0)
		txtTotalZimmer3.Text = String.Format("{0:N}", 0)

		txtPreis03.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(5).Item(5))
		txtPreis04.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(3).Item(5))
		txtPreis05.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(24).Item(5))
		txtPreis06.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(25).Item(5))
		txtPreis07.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(26).Item(5))

		Preis_Bett_4er = String.Format("{0:N}", ds.Tables("tbPreis").Rows(17).Item(5))
		Preis_Bett_6er = String.Format("{0:N}", ds.Tables("tbPreis").Rows(18).Item(5))

		Preis_Zimmer1_Einzel = String.Format("{0:N}", ds.Tables("tbPreis").Rows(19).Item(5))
		Preis_Zimmer2_Einzel = String.Format("{0:N}", ds.Tables("tbPreis").Rows(20).Item(5))
		Preis_Zimmer3_Einzel = String.Format("{0:N}", ds.Tables("tbPreis").Rows(19).Item(5))

		Preis_Zimmer1_Komplett = String.Format("{0:N}", ds.Tables("tbPreis").Rows(21).Item(5))
		Preis_Zimmer2_Komplett = String.Format("{0:N}", ds.Tables("tbPreis").Rows(22).Item(5))
		Preis_Zimmer3_Komplett = String.Format("{0:N}", ds.Tables("tbPreis").Rows(21).Item(5))

		Preis_Zimmer1_Kinderbett = String.Format("{0:N}", ds.Tables("tbPreis").Rows(23).Item(5))
		Preis_Zimmer2_Kinderbett = String.Format("{0:N}", ds.Tables("tbPreis").Rows(23).Item(5))
		Preis_Zimmer3_Kinderbett = String.Format("{0:N}", ds.Tables("tbPreis").Rows(23).Item(5))


		Berechnen_Zimmer_StartMutation()

		Berechnen()
		TotalBerechnen()


	End Sub

	Public Sub StartMutation(ByVal i As Integer)

		StartFlag = "StartMutation"

		Dim BookingId As Integer
		BookingId = i
		Public_BookingId = i

		lblBookingId.Text = BookingId.ToString

		Dim da As New OleDbDataAdapter("SELECT * FROM tbZimmer " _
				& "INNER JOIN (tbAdress INNER JOIN tbZimmertbAdress ON tbAdress.AdressId = tbZimmertbAdress.AdressId) " _
				& "ON tbZimmer.BookingId = tbZimmertbAdress.BookingId where tbZimmertbAdress.Leader = true and tbZimmer.BookingId = " & BookingId & "", conn)

		Dim ds As New DataSet()

		da.Fill(ds, "tbZimmer")

		Public_AdressId = ds.Tables("tbZimmer").Rows(0).Item(41).ToString()

		lblNameKunde.Text = (ds.Tables("tbZimmer").Rows(0).Item(42).ToString()) & " " & (ds.Tables("tbZimmer").Rows(0).Item(43).ToString())

		DateTimePicker1.Value = ds.Tables("tbZimmer").Rows(0).Item(1).ToString()

		DateTimePicker2.Value = ds.Tables("tbZimmer").Rows(0).Item(2).ToString()

		txtNaechte.Text = ds.Tables("tbZimmer").Rows(0).Item(3).ToString()

		txtAutoKennZeichen.Text = ds.Tables("tbZimmer").Rows(0).Item(35).ToString()

		txtBemerkungen.Text = ds.Tables("tbZimmer").Rows(0).Item(36).ToString()

		txtRabatt.Text = ds.Tables("tbZimmer").Rows(0).Item(37).ToString()

		If AdressenErfassen.CheckDBNull(ds.Tables("tbZimmer").Rows(0).Item(38)) <> "" Then
			If ds.Tables("tbZimmer").Rows(0).Item(38) = "%" Then
				txtRabattProzent.Checked = True
			Else
				txtRabattCHF.Checked = True
			End If
		End If


		chkBett11.Checked = ds.Tables("tbZimmer").Rows(0).Item(7)
		chkBett12.Checked = ds.Tables("tbZimmer").Rows(0).Item(8)
		chkBett13.Checked = ds.Tables("tbZimmer").Rows(0).Item(9)
		chkBett14.Checked = ds.Tables("tbZimmer").Rows(0).Item(10)
		chkZimmer1Komplett.Checked = ds.Tables("tbZimmer").Rows(0).Item(11)
		chkZimmer1Einzel.Checked = ds.Tables("tbZimmer").Rows(0).Item(12)
		chkZimmer1Kinderbett.Checked = ds.Tables("tbZimmer").Rows(0).Item(13)

		chkBett21.Checked = ds.Tables("tbZimmer").Rows(0).Item(14)
		chkBett22.Checked = ds.Tables("tbZimmer").Rows(0).Item(15)
		chkBett23.Checked = ds.Tables("tbZimmer").Rows(0).Item(16)
		chkBett24.Checked = ds.Tables("tbZimmer").Rows(0).Item(17)
		chkBett25.Checked = ds.Tables("tbZimmer").Rows(0).Item(18)
		chkBett26.Checked = ds.Tables("tbZimmer").Rows(0).Item(19)
		chkZimmer2Komplett.Checked = ds.Tables("tbZimmer").Rows(0).Item(20)
		chkZimmer2Einzel.Checked = ds.Tables("tbZimmer").Rows(0).Item(21)
		chkZimmer2Kinderbett.Checked = ds.Tables("tbZimmer").Rows(0).Item(22)

		chkBett31.Checked = ds.Tables("tbZimmer").Rows(0).Item(23)
		chkBett32.Checked = ds.Tables("tbZimmer").Rows(0).Item(24)
		chkBett33.Checked = ds.Tables("tbZimmer").Rows(0).Item(25)
		chkBett34.Checked = ds.Tables("tbZimmer").Rows(0).Item(26)
		chkZimmer3Komplett.Checked = ds.Tables("tbZimmer").Rows(0).Item(27)
		chkZimmer3Einzel.Checked = ds.Tables("tbZimmer").Rows(0).Item(28)
		chkZimmer3Kinderbett.Checked = ds.Tables("tbZimmer").Rows(0).Item(29)

		txtSonstigesText.Text = ds.Tables("tbZimmer").Rows(0).Item(38).ToString()
		txtSonstigesBetrag.Text = ds.Tables("tbZimmer").Rows(0).Item(39).ToString()

		cmb01.Text = ds.Tables("tbZimmer").Rows(0).Item(5)
		cmb02.Text = ds.Tables("tbZimmer").Rows(0).Item(6)
		cmb03.Text = ds.Tables("tbZimmer").Rows(0).Item(30)
		cmb04.Text = ds.Tables("tbZimmer").Rows(0).Item(31)
		cmb05.Text = ds.Tables("tbZimmer").Rows(0).Item(32)
		cmb06.Text = ds.Tables("tbZimmer").Rows(0).Item(33)
		cmb07.Text = ds.Tables("tbZimmer").Rows(0).Item(34)


		'-----------------------------------------------------------------------

		Dim da1 As New OleDbDataAdapter("SELECT tbAdress.AdressId, tbAdress.NachName, tbAdress.VorName,tbAdress.Geburtsdatum, tbAdress.Land FROM tbZimmertbAdress " _
					& "INNER JOIN tbAdress ON tbZimmertbAdress.AdressId = tbAdress.AdressId " _
					& "where tbZimmertbAdress.BookingId = " & BookingId & "", conn)

		Dim ds1 As New DataSet()

		da1.Fill(ds1, "tbZimmer")

		Dim rc As Integer = ds1.Tables("tbZimmer").Rows.Count()
		rc = rc - 1

		If rc >= 1 Then
			txtNachName1.Text = ds1.Tables("tbZimmer").Rows(1).Item(1).ToString()
			txtVorname1.Text = ds1.Tables("tbZimmer").Rows(1).Item(2).ToString()
			txtGeburtsdatum1.Text = ds1.Tables("tbZimmer").Rows(1).Item(3)
			LandErfasst1 = ds1.Tables("tbZimmer").Rows(1).Item(4).ToString()
			lblId1.Text = ds1.Tables("tbZimmer").Rows(1).Item(0).ToString()
		End If

		If rc >= 2 Then
			txtNachName2.Text = ds1.Tables("tbZimmer").Rows(2).Item(1).ToString()
			txtVorname2.Text = ds1.Tables("tbZimmer").Rows(2).Item(2).ToString()
			txtGeburtsdatum2.Text = ds1.Tables("tbZimmer").Rows(2).Item(3).ToString()
			LandErfasst2 = ds1.Tables("tbZimmer").Rows(2).Item(4).ToString()
			lblId2.Text = ds1.Tables("tbZimmer").Rows(2).Item(0).ToString()
		End If

		If rc >= 3 Then
			txtNachName3.Text = ds1.Tables("tbZimmer").Rows(3).Item(1).ToString()
			txtVorname3.Text = ds1.Tables("tbZimmer").Rows(3).Item(2).ToString()
			txtGeburtsdatum3.Text = ds1.Tables("tbZimmer").Rows(3).Item(3).ToString()
			LandErfasst3 = ds1.Tables("tbZimmer").Rows(3).Item(4).ToString()
			lblId3.Text = ds1.Tables("tbZimmer").Rows(3).Item(0).ToString()
		End If

		If rc >= 4 Then
			txtNachName4.Text = ds1.Tables("tbZimmer").Rows(4).Item(1).ToString()
			txtVorname4.Text = ds1.Tables("tbZimmer").Rows(4).Item(2).ToString()
			txtGeburtsdatum4.Text = ds1.Tables("tbZimmer").Rows(4).Item(3).ToString()
			LandErfasst4 = ds1.Tables("tbZimmer").Rows(4).Item(4).ToString()
			lblId4.Text = ds1.Tables("tbZimmer").Rows(4).Item(0).ToString()
		End If

		If rc >= 5 Then
			txtNachName5.Text = ds1.Tables("tbZimmer").Rows(5).Item(1).ToString()
			txtVorname5.Text = ds1.Tables("tbZimmer").Rows(5).Item(2).ToString()
			txtGeburtsdatum5.Text = ds1.Tables("tbZimmer").Rows(5).Item(3).ToString()
			LandErfasst5 = ds1.Tables("tbZimmer").Rows(5).Item(4).ToString()
			lblId5.Text = ds1.Tables("tbZimmer").Rows(5).Item(0).ToString()
		End If

		If rc >= 6 Then
			txtNachName6.Text = ds1.Tables("tbZimmer").Rows(6).Item(1).ToString()
			txtVorname6.Text = ds1.Tables("tbZimmer").Rows(6).Item(2).ToString()
			txtGeburtsdatum6.Text = ds1.Tables("tbZimmer").Rows(6).Item(3).ToString()
			LandErfasst6 = ds1.Tables("tbZimmer").Rows(6).Item(4).ToString()
			lblId6.Text = ds1.Tables("tbZimmer").Rows(6).Item(0).ToString()
		End If

		If rc >= 7 Then
			txtNachName7.Text = ds1.Tables("tbZimmer").Rows(7).Item(1).ToString()
			txtVorname7.Text = ds1.Tables("tbZimmer").Rows(7).Item(2).ToString()
			txtGeburtsdatum7.Text = ds1.Tables("tbZimmer").Rows(7).Item(3).ToString()
			LandErfasst7 = ds1.Tables("tbZimmer").Rows(7).Item(4).ToString()
			lblId7.Text = ds1.Tables("tbZimmer").Rows(7).Item(0).ToString()
		End If

		If rc >= 8 Then
			txtNachName8.Text = ds1.Tables("tbZimmer").Rows(8).Item(1).ToString()
			txtVorname8.Text = ds1.Tables("tbZimmer").Rows(8).Item(2).ToString()
			txtGeburtsdatum8.Text = ds1.Tables("tbZimmer").Rows(8).Item(3).ToString()
			LandErfasst8 = ds1.Tables("tbZimmer").Rows(8).Item(4).ToString()
			lblId8.Text = ds1.Tables("tbZimmer").Rows(8).Item(0).ToString()
		End If

		If rc >= 9 Then
			txtNachName9.Text = ds1.Tables("tbZimmer").Rows(9).Item(1).ToString()
			txtVorname9.Text = ds1.Tables("tbZimmer").Rows(9).Item(2).ToString()
			txtGeburtsdatum9.Text = ds1.Tables("tbZimmer").Rows(9).Item(3).ToString()
			LandErfasst9 = ds1.Tables("tbZimmer").Rows(9).Item(4).ToString()
			lblId9.Text = ds1.Tables("tbZimmer").Rows(9).Item(0).ToString()
		End If

		If rc >= 10 Then
			txtNachName10.Text = ds1.Tables("tbZimmer").Rows(10).Item(1).ToString()
			txtVorname10.Text = ds1.Tables("tbZimmer").Rows(10).Item(2).ToString()
			txtGeburtsdatum10.Text = ds1.Tables("tbZimmer").Rows(10).Item(3).ToString()
			LandErfasst10 = ds1.Tables("tbZimmer").Rows(10).Item(4).ToString()
			lblId10.Text = ds1.Tables("tbZimmer").Rows(10).Item(0).ToString()
		End If


		'-----------------------------------------------------------------------


		If lblBookingId.Text = "---" Then
			btnBestaetigungMitMeldeschein.Enabled = False
			btnBestaetigungOhneMeldeschein.Enabled = False
			btnQuittung.Enabled = False

		Else
			btnBestaetigungMitMeldeschein.Enabled = True
			btnBestaetigungOhneMeldeschein.Enabled = True
			btnQuittung.Enabled = True

		End If


		If ds.Tables("tbZimmer").Rows(0).Item(4).ToString() = False Then

			btnSpeichern.Enabled = False
			btnSpeichernUndSchliessen.Enabled = False
			btnBestaetigungMitMeldeschein.Enabled = False
			btnBestaetigungOhneMeldeschein.Enabled = False
			btnQuittung.Enabled = False
		End If



	End Sub

	Private Sub MutierenBookingDaten()


		Check_ZimmerBuchungen()

		If CheckZimmerBuchungen = True Then
			Exit Sub
		End If

		Delete_ZimmerBuchungen()



		Dim strRabattArt As String

		If txtRabattProzent.Checked = True Then
			strRabattArt = "%"
		Else
			If txtRabattCHF.Checked = True Then
				strRabattArt = "CHF"
			Else
				strRabattArt = ""
			End If
		End If

		Dim cmd As New OleDbCommand("UPDATE tbZimmer SET Anreise = '" & DateTimePicker1.Value.Date & "'" _
			& ", Abreise = '" & DateTimePicker2.Value.Date & "'" _
			& ", AnzahlNaechte = '" & txtNaechte.Text & "'" _
			& ", Aktiv = " & True & "" _
			& ", Erwachsene = '" & cmb01.Text & "'" _
			& ", Kinder = '" & cmb02.Text & "'" _
			& ", B11 = '" & chkBett11.CheckState & "'" _
			& ", B12 = '" & chkBett12.CheckState & "'" _
			& ", B13 = '" & chkBett13.CheckState & "'" _
			& ", B14 = '" & chkBett14.CheckState & "'" _
			& ", Z1Z = '" & chkZimmer1Komplett.CheckState & "'" _
			& ", Z1E = '" & chkZimmer1Einzel.CheckState & "'" _
			& ", Z1K = '" & chkZimmer1Kinderbett.CheckState & "'" _
			& ", B21 = '" & chkBett21.CheckState & "'" _
			& ", B22 = '" & chkBett22.CheckState & "'" _
			& ", B23 = '" & chkBett23.CheckState & "'" _
			& ", B24 = '" & chkBett24.CheckState & "'" _
			& ", B25 = '" & chkBett25.CheckState & "'" _
			& ", B26 = '" & chkBett26.CheckState & "'" _
			& ", Z2Z = '" & chkZimmer2Komplett.CheckState & "'" _
			& ", Z2E = '" & chkZimmer2Einzel.CheckState & "'" _
			& ", Z2K = '" & chkZimmer2Kinderbett.CheckState & "'" _
			& ", B31 = '" & chkBett31.CheckState & "'" _
			& ", B32 = '" & chkBett32.CheckState & "'" _
			& ", B33 = '" & chkBett33.CheckState & "'" _
			& ", B34 = '" & chkBett34.CheckState & "'" _
			& ", Z3Z = '" & chkZimmer3Komplett.CheckState & "'" _
			& ", Z3E = '" & chkZimmer3Einzel.CheckState & "'" _
			& ", Z3K = '" & chkZimmer3Kinderbett.CheckState & "'" _
			& ", ParkplatzMotorrad = '" & cmb03.Text & "'" _
			& ", ParkplatzAuto = '" & cmb04.Text & "'" _
			& ", Handtuecher = '" & cmb05.Text & "'" _
			& ", Fruehstueck = '" & cmb06.Text & "'" _
			& ", Kueche = '" & cmb07.Text & "'" _
			& ", Autokennzeichen = '" & txtAutoKennZeichen.Text & "'" _
			& ", Bemerkungen = '" & txtBemerkungen.Text & "'" _
			& ", RabattWert = '" & txtRabatt.Text & "'" _
			& ", RabattArt = '" & strRabattArt.ToString & "'" _
			& ", SonstigesText = '" & txtSonstigesText.Text & "'" _
			& ", SonstigesBetrag = '" & txtSonstigesBetrag.Text & "'" _
		& " WHERE BookingId = " & Public_BookingId & "", conn)


		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet("tbAdress")

		Try
			da.Fill(ds, "tbAdress")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try

		End Sub

	Private Sub txtSonstigesBetrag_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtSonstigesBetrag.TextChanged
		TotalBerechnen()
	End Sub

	Private Sub txtRabatt_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRabatt.TextChanged
		TotalBerechnen()
	End Sub

	Private Sub txtRabattProzent_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRabattProzent.Click
		TotalBerechnen()
	End Sub

	Private Sub txtRabattCHF_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles txtRabattCHF.Click
		TotalBerechnen()
	End Sub

	Private Sub cmb01_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb01.SelectedIndexChanged
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub cmb02_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb02.SelectedIndexChanged
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub cmb03_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb03.SelectedIndexChanged
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub cmb04_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb04.SelectedIndexChanged
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub cmb05_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb05.SelectedIndexChanged
		Berechnen()
		TotalBerechnen()
	End Sub
	Private Sub cmb06_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb06.SelectedIndexChanged
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub cmb07_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb07.SelectedIndexChanged
		Berechnen()
		TotalBerechnen()
	End Sub


Private Sub Berechnen()

		Dim s1, s2, s3, s4, s5, s6, s7 As Double
		Dim w1, w2, w3, w4, w5, w6, w7 As Double
		Dim v1, v2, v3, v4, v5, v6, v7 As Integer

		Dim AnzahlNaechte As Integer = CInt(txtNaechte.Text)

		If txtPreis03.Text <> "" Then w3 = txtPreis03.Text
		If txtPreis04.Text <> "" Then w4 = txtPreis04.Text
		If txtPreis05.Text <> "" Then w5 = txtPreis05.Text
		If txtPreis06.Text <> "" Then w6 = txtPreis06.Text
		If txtPreis07.Text <> "" Then w7 = txtPreis07.Text

		If cmb03.Text <> "" Then v3 = cmb03.Text
		If cmb04.Text <> "" Then v4 = cmb04.Text
		If cmb05.Text <> "" Then v5 = cmb05.Text
		If cmb06.Text <> "" Then v6 = cmb06.Text
		If cmb07.Text <> "" Then v7 = cmb07.Text


		s3 = w3 * v3 * AnzahlNaechte
		s4 = w4 * v4 * AnzahlNaechte
		s5 = w5 * v5
		s6 = w6 * v6
		s7 = w7 * v7

		total03.Text = String.Format("{0:N}", s3)
		total04.Text = String.Format("{0:N}", s4)
		total05.Text = String.Format("{0:N}", s5)
		total06.Text = String.Format("{0:N}", s6)
		total07.Text = String.Format("{0:N}", s7)


		Dim z1 As Double
		Dim z2 As Double
		Dim z3 As Double

		z1 = Preis_Zimmer1_Total * AnzahlNaechte
		z2 = Preis_Zimmer2_Total * AnzahlNaechte
		z3 = Preis_Zimmer3_Total * AnzahlNaechte

		txtTotalZimmer1Total.Text = String.Format("{0:N}", z1)
		txtTotalZimmer2Total.Text = String.Format("{0:N}", z2)
		txtTotalZimmer3Total.Text = String.Format("{0:N}", z3)

	End Sub

Private Sub TotalBerechnen()

		Dim total As Double
		Dim t3, t4, t5, t6, t7 As Double


		'TotalZimmer
		Dim totalZimmer As Double
		totalZimmer = Preis_Zimmer1_Total + Preis_Zimmer2_Total + Preis_Zimmer3_Total

		'TotalNächte
		Dim TotalNaechte As Integer = CInt(txtNaechte.Text)


		If total03.Text = "" Then total03.Text = 0
		If total04.Text = "" Then total04.Text = 0
		If total05.Text = "" Then total05.Text = 0
		If total06.Text = "" Then total06.Text = 0
		If total07.Text = "" Then total07.Text = 0

		If total03.Text <> "" Then t3 = total03.Text
		If total04.Text <> "" Then t4 = total04.Text
		If total05.Text <> "" Then t5 = total05.Text
		If total06.Text <> "" Then t6 = total06.Text
		If total07.Text <> "" Then t7 = total07.Text


		total = t3 + t4 + t5 + t6 + t7 + totalZimmer * TotalNaechte


		'Zwischentotal
		txtZwischenTotal.Text = String.Format("{0:N}", total)
		Dim ZwischenTotal As Double = txtZwischenTotal.Text


		'TotalPersonen
		Dim AnzahlErwachsene As Integer
		Dim AnzahlKinder As Integer
		Dim TotalPersonen As Integer


		'Kinder 0 - 15 Jahre müssen keine Taxe bezahlen
		'If cmb01.Text <> "" Then AnzahlErwachsene = CInt(cmb01.Text)
		'If cmb02.Text <> "" Then AnzahlKinder = CInt(cmb02.Text)
		'TotalPersonen = AnzahlErwachsene + AnzahlKinder

		If cmb01.Text <> "" Then AnzahlErwachsene = CInt(cmb01.Text)
		TotalPersonen = AnzahlErwachsene

		'Taxe
		Dim Taxe As Double = TaxeLesen()
		Dim TotalTaxe As Double
		TotalTaxe = TotalPersonen * Taxe * TotalNaechte

		'Sonstiges
		Dim SonstigesBetrag As Double
		If txtSonstigesBetrag.Text <> "" Then
			SonstigesBetrag = txtSonstigesBetrag.Text
			ZwischenTotal = ZwischenTotal + SonstigesBetrag + TotalTaxe
			txtZwischenTotal.Text = String.Format("{0:N}", ZwischenTotal)
		Else
			ZwischenTotal = ZwischenTotal + TotalTaxe
			txtZwischenTotal.Text = String.Format("{0:N}", ZwischenTotal)
		End If

		'TotalRechnung
		Dim TempTotalRechnung As Double
		TempTotalRechnung = ZwischenTotal


		Dim Rabatt As Double
		Dim RabattArt As String

		If txtRabatt.Text <> "" Then

			Rabatt = CDbl(txtRabatt.Text)

			If txtRabattProzent.Checked = True Then
				RabattArt = "%"
			Else
				If txtRabattCHF.Checked = True Then
					RabattArt = "CHF"
				Else
					RabattArt = ""
				End If
			End If


			If RabattArt = "CHF" Then
				TempTotalRechnung = TempTotalRechnung - Rabatt
				Me.Label38.Text = Rabatt
		End If

			If RabattArt = "%" Then
				Dim Rechnungswert As Double = TempTotalRechnung * Rabatt / 100
				TempTotalRechnung = TempTotalRechnung - Rechnungswert
				Me.Label38.Text = Rechnungswert
			End If

		End If


		Dim Mwst As Double = MwstLesen()
		Dim SummeMwst As Double
		SummeMwst = Math.Round((TempTotalRechnung * Mwst / 100), 1)
		TempTotalRechnung = TempTotalRechnung + SummeMwst

		Dim TotalRechnung As Double
		TotalRechnung = TempTotalRechnung

		'Ausgabe
		txtMwSt.Text = String.Format("{0:N}", SummeMwst)
		txtTaxe.Text = String.Format("{0:N}", TotalTaxe)
		txtTotalRechnung.Text = String.Format("{0:N}", (TotalRechnung))

	End Sub

	Private Sub DateTimePicker2_DatumNeu()

		DateTimePicker2.Value = DateTimePicker2.Value.AddDays(1)

	End Sub

		Private Sub DateTimePicker1_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker1.ValueChanged

		Dim AnzahlNaechte As Long = DateDiff(DateInterval.Day, _
										DateTimePicker1.Value.Date, _
										DateTimePicker2.Value.Date, _
										FirstDayOfWeek.Monday, _
										FirstWeekOfYear.Jan1)

		Me.txtNaechte.Text = AnzahlNaechte.ToString

		Berechnen()
		TotalBerechnen()

	End Sub


	Private Sub DateTimePicker2_ValueChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles DateTimePicker2.ValueChanged

		Dim AnzahlNaechte As Long = DateDiff(DateInterval.Day, _
										DateTimePicker1.Value.Date, _
										DateTimePicker2.Value.Date, _
										FirstDayOfWeek.Monday, _
										FirstWeekOfYear.Jan1)

		Me.txtNaechte.Text = AnzahlNaechte.ToString

		Berechnen()
		TotalBerechnen()

	End Sub

	Private Sub btnSpeichern_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSpeichern.Click

		If StartFlag = "StartBooking" Then

		
			SpeichernBookingDaten()

			If CheckZimmerBuchungen = True Then
				Exit Sub
			End If

			Update_ZimmerBuchungen()

			StartSpeichernBegleitpersonen()
			MsgBox("Buchungsdaten gespeichert", MsgBoxStyle.Exclamation)

			StartMutation(Public_BookingId)

		Else

			MutierenBookingDaten()

			If CheckZimmerBuchungen = True Then
				Exit Sub
			End If

			Update_ZimmerBuchungen()

			StartMutierenBegleitpersonen()
			MsgBox("Buchungsdaten mutiert", MsgBoxStyle.Exclamation)

		End If


	End Sub

		Private Sub btnSpeichernUndSchliessen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSpeichernUndSchliessen.Click

		If StartFlag = "StartBooking" Then

			SpeichernBookingDaten()

			If CheckZimmerBuchungen = True Then
				Exit Sub
			End If

			Update_ZimmerBuchungen()

			StartSpeichernBegleitpersonen()
			MsgBox("Buchungsdaten gespeichert", MsgBoxStyle.Exclamation)

		Else

			MutierenBookingDaten()

			If CheckZimmerBuchungen = True Then
				Exit Sub
			End If

			Update_ZimmerBuchungen()

			StartMutierenBegleitpersonen()
			MsgBox("Buchungsdaten mutiert", MsgBoxStyle.Exclamation)

		End If

		Me.Close()

	End Sub

	Private Sub StartSpeichernBegleitpersonen()

		If txtNachName1.Text <> "" Then SpeichernBegleitpersonen(txtNachName1.Text, txtVorname1.Text, txtGeburtsdatum1.Text, cmbLand1.Text)
		If txtNachName2.Text <> "" Then SpeichernBegleitpersonen(txtNachName2.Text, txtVorname2.Text, txtGeburtsdatum2.Text, cmbLand2.Text)
		If txtNachName3.Text <> "" Then SpeichernBegleitpersonen(txtNachName3.Text, txtVorname3.Text, txtGeburtsdatum3.Text, cmbLand3.Text)
		If txtNachName4.Text <> "" Then SpeichernBegleitpersonen(txtNachName4.Text, txtVorname4.Text, txtGeburtsdatum4.Text, cmbLand4.Text)
		If txtNachName5.Text <> "" Then SpeichernBegleitpersonen(txtNachName5.Text, txtVorname5.Text, txtGeburtsdatum5.Text, cmbLand5.Text)
		If txtNachName6.Text <> "" Then SpeichernBegleitpersonen(txtNachName6.Text, txtVorname6.Text, txtGeburtsdatum6.Text, cmbLand6.Text)
		If txtNachName7.Text <> "" Then SpeichernBegleitpersonen(txtNachName7.Text, txtVorname7.Text, txtGeburtsdatum7.Text, cmbLand7.Text)
		If txtNachName8.Text <> "" Then SpeichernBegleitpersonen(txtNachName8.Text, txtVorname8.Text, txtGeburtsdatum8.Text, cmbLand8.Text)
		If txtNachName9.Text <> "" Then SpeichernBegleitpersonen(txtNachName9.Text, txtVorname9.Text, txtGeburtsdatum9.Text, cmbLand9.Text)
		If txtNachName10.Text <> "" Then SpeichernBegleitpersonen(txtNachName10.Text, txtVorname10.Text, txtGeburtsdatum10.Text, cmbLand10.Text)

	End Sub

	Private Sub SpeichernBegleitpersonen(ByVal strName As String, ByVal strVorname As String, ByVal strGeburtsdatum As String, ByVal strLand As String)

		If strGeburtsdatum = "" Then
			strGeburtsdatum = "01.01.1900"
		End If

		strGeburtsdatum = CDate(strGeburtsdatum)

		Dim strSuchfeld As String = strName & " " & strVorname & " " & strGeburtsdatum

		Dim cmd As New OleDbCommand("Insert Into tbAdress (NachName," _
		& "Vorname, " _
		& "Geburtsdatum, " _
		& "Land, " _
		& "Suchfeld " _
		& ") Values('" & strName & "'" _
		& ", '" & strVorname & "'" _
		& ", '" & strGeburtsdatum & "'" _
		& ", '" & strLand & "'" _
		& ", '" & strSuchfeld & "')", conn)

		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet("tbAdress")

		Try
			da.Fill(ds, "tbAdress")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try

		strSuchfeld = ""
		strGeburtsdatum = ""
		strLand = ""
		strName = ""
		strVorname = ""

		tbZimmertbAdressSchreiben()

	End Sub

	 Private Sub tbZimmertbAdressSchreiben()

		Dim cmd As New OleDbCommand("Select * from tbAdress Order by AdressId ASC", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet()

		Dim k, i As Integer

		da.Fill(ds, "tbAdress")

		k = ds.Tables("tbAdress").Rows.Count
		i = ds.Tables("tbAdress").Rows(k - 1).Item(0)

		'----------------------------------------------------------------

		Dim cmd1 As New OleDbCommand("Select * from tbZimmer Order by BookingId ASC", conn)
		Dim da1 As New OleDbDataAdapter(cmd1)
		Dim ds1 As New DataSet()

		Dim v, w As Integer

		da1.Fill(ds1, "tbZimmer")

		v = ds1.Tables("tbZimmer").Rows.Count
		w = ds1.Tables("tbZimmer").Rows(v - 1).Item(0)

		'Wenn die Buchung besteht, wir die Buchungsnummer für die Mutation der Begl. Personen verwendet
		'Nur bei Neuerfassung wird der letzte Record gelesen

		If StartFlag = "StartMutation" Then
			w = Public_BookingId
		End If

		'----------------------------------------------------------------

		Dim cmd2 As New OleDbCommand("Insert Into tbZimmertbAdress (BookingId, AdressId, Moddate) Values(" & w & ", " & i & ",Now())", conn)
		Dim da2 As New OleDbDataAdapter(cmd2)
		Dim ds2 As New DataSet()

		Try
			da2.Fill(ds2, "tbZimmertbAdress")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try

		End Sub

Private Sub StartMutierenBegleitpersonen()

		If txtNachName1.Text = "" Then
			If lblId1.Text <> "Id1" Then
				MutierenBegleitpersonen(lblId1.Text)
			End If
		Else
			If lblId1.Text = "Id1" Then
				SpeichernBegleitpersonen(txtNachName1.Text, txtVorname1.Text, txtGeburtsdatum1.Text, cmbLand1.Text)
			Else
				UpdateBegleitpersonen(lblId1.Text, txtNachName1.Text, txtVorname1.Text, txtGeburtsdatum1.Text, cmbLand1.Text)
			End If
		End If

		If txtNachName2.Text = "" Then
			If lblId2.Text <> "Id2" Then
				MutierenBegleitpersonen(lblId2.Text)
			End If
		Else
			If lblId2.Text = "Id2" Then
				SpeichernBegleitpersonen(txtNachName2.Text, txtVorname2.Text, txtGeburtsdatum2.Text, cmbLand2.Text)
			Else
				UpdateBegleitpersonen(lblId2.Text, txtNachName2.Text, txtVorname2.Text, txtGeburtsdatum2.Text, cmbLand2.Text)
			End If
		End If

		If txtNachName3.Text = "" Then
			If lblId3.Text <> "Id3" Then
				MutierenBegleitpersonen(lblId3.Text)
			End If
		Else
			If lblId3.Text = "Id3" Then
				SpeichernBegleitpersonen(txtNachName3.Text, txtVorname3.Text, txtGeburtsdatum3.Text, cmbLand3.Text)
			Else
				UpdateBegleitpersonen(lblId3.Text, txtNachName3.Text, txtVorname3.Text, txtGeburtsdatum3.Text, cmbLand3.Text)
			End If
		End If


		If txtNachName4.Text = "" Then
			If lblId4.Text <> "Id4" Then
				MutierenBegleitpersonen(lblId4.Text)
			End If
		Else
			If lblId4.Text = "Id4" Then
				SpeichernBegleitpersonen(txtNachName4.Text, txtVorname4.Text, txtGeburtsdatum4.Text, cmbLand4.Text)
			Else
				UpdateBegleitpersonen(lblId4.Text, txtNachName4.Text, txtVorname4.Text, txtGeburtsdatum4.Text, cmbLand4.Text)
			End If
		End If


		If txtNachName5.Text = "" Then
			If lblId5.Text <> "Id5" Then
				MutierenBegleitpersonen(lblId5.Text)
			End If
		Else
			If lblId5.Text = "Id5" Then
				SpeichernBegleitpersonen(txtNachName5.Text, txtVorname5.Text, txtGeburtsdatum5.Text, cmbLand5.Text)
			Else
				UpdateBegleitpersonen(lblId5.Text, txtNachName5.Text, txtVorname5.Text, txtGeburtsdatum5.Text, cmbLand5.Text)
			End If
		End If

		If txtNachName6.Text = "" Then
			If lblId6.Text <> "Id6" Then
				MutierenBegleitpersonen(lblId6.Text)
			End If
		Else
			If lblId6.Text = "Id6" Then
				SpeichernBegleitpersonen(txtNachName6.Text, txtVorname6.Text, txtGeburtsdatum6.Text, cmbLand6.Text)
			Else
				UpdateBegleitpersonen(lblId6.Text, txtNachName6.Text, txtVorname6.Text, txtGeburtsdatum6.Text, cmbLand6.Text)
			End If
		End If


		If txtNachName7.Text = "" Then
			If lblId7.Text <> "Id7" Then
				MutierenBegleitpersonen(lblId7.Text)
			End If
		Else
			If lblId7.Text = "Id7" Then
				SpeichernBegleitpersonen(txtNachName7.Text, txtVorname7.Text, txtGeburtsdatum7.Text, cmbLand7.Text)
			Else
				UpdateBegleitpersonen(lblId7.Text, txtNachName7.Text, txtVorname7.Text, txtGeburtsdatum7.Text, cmbLand7.Text)
			End If
		End If


		If txtNachName8.Text = "" Then
			If lblId8.Text <> "Id8" Then
				MutierenBegleitpersonen(lblId8.Text)
			End If
		Else
			If lblId8.Text = "Id8" Then
				SpeichernBegleitpersonen(txtNachName8.Text, txtVorname8.Text, txtGeburtsdatum8.Text, cmbLand8.Text)
			Else
				UpdateBegleitpersonen(lblId8.Text, txtNachName8.Text, txtVorname8.Text, txtGeburtsdatum8.Text, cmbLand8.Text)
			End If
		End If


		If txtNachName9.Text = "" Then
			If lblId9.Text <> "Id9" Then
				MutierenBegleitpersonen(lblId9.Text)
			End If
		Else
			If lblId9.Text = "Id9" Then
				SpeichernBegleitpersonen(txtNachName9.Text, txtVorname9.Text, txtGeburtsdatum9.Text, cmbLand9.Text)
			Else
				UpdateBegleitpersonen(lblId9.Text, txtNachName9.Text, txtVorname9.Text, txtGeburtsdatum9.Text, cmbLand9.Text)
			End If
		End If


		If txtNachName10.Text = "" Then
			If lblId10.Text <> "Id10" Then
				MutierenBegleitpersonen(lblId10.Text)
			End If
		Else
			If lblId10.Text = "Id10" Then
				SpeichernBegleitpersonen(txtNachName10.Text, txtVorname10.Text, txtGeburtsdatum10.Text, cmbLand10.Text)
			Else
				UpdateBegleitpersonen(lblId10.Text, txtNachName10.Text, txtVorname10.Text, txtGeburtsdatum10.Text, cmbLand10.Text)
			End If
		End If

	End Sub

	 Private Sub MutierenBegleitpersonen(ByVal Id As Integer)

		Dim cmd As New OleDbCommand("Delete * From tbAdress Where AdressId = " & Id & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet("tbAdress")

		Try
			da.Fill(ds, "tbAdress")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try

		'--------------------------------------

		Dim cmd1 As New OleDbCommand("Delete * From tbBookingtbAdress Where AdressId = " & Id & "", conn)
		Dim da1 As New OleDbDataAdapter(cmd1)
		Dim ds1 As New DataSet("tbBookingtbAdress")

		Try
			da1.Fill(ds, "tbBookingtbAdress")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try

		'--------------------------------------

	End Sub

Private Sub UpdateBegleitpersonen(ByVal intAdressId As Integer, ByVal strName As String, ByVal strVorname As String, ByVal strGeburtsdatum As String, ByVal strLand As String)

		If strGeburtsdatum = "" Then
			strGeburtsdatum = "01.01.1900"
		End If

		strGeburtsdatum = CDate(strGeburtsdatum)

		Dim strSuchfeld As String = strName & " " & strVorname & " " & strGeburtsdatum & " " & strLand

		Dim cmd As New OleDbCommand("UPDATE tbAdress SET NachName = '" & strName & "'" _
			& ", Vorname = '" & strVorname & "'" _
			& ", Land = '" & strLand & "'" _
			& ", Suchfeld = '" & strSuchfeld & "' WHERE AdressId = " & intAdressId & "", conn)

		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet("tbBooking")

		Try
			da.Fill(ds, "tbAdress")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try

		strSuchfeld = ""
		strGeburtsdatum = ""
		strLand = ""
		strName = ""
		strVorname = ""

	End Sub

	Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

		Me.Close()

	End Sub


	Public Function TaxeLesen()

		Dim datName As String = "taxe.ini"
		Dim reader As StreamReader = File.OpenText(datName)

		While (reader.Peek() > -1)

			Return reader.ReadLine()
			Exit While

		End While

		reader.Close()

	End Function

	Public Function MwstLesen()

		Dim datName As String = "mwst.ini"
		Dim reader As StreamReader = File.OpenText(datName)

		While (reader.Peek() > -1)

			Return reader.ReadLine()
			Exit While

		End While

		reader.Close()

	End Function

	Private Sub chkBett11_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett11.CheckedChanged
		If chkZimmer1Komplett.Checked = False And chkZimmer1Einzel.Checked = False Then Berechnen_Zimmer(1, 11)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkBett12_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett12.CheckedChanged
		If chkZimmer1Komplett.Checked = False And chkZimmer1Einzel.Checked = False Then Berechnen_Zimmer(1, 12)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkBett13_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett13.CheckedChanged
		If chkZimmer1Komplett.Checked = False And chkZimmer1Einzel.Checked = False Then Berechnen_Zimmer(1, 13)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkBett14_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett14.CheckedChanged
		If chkZimmer1Komplett.Checked = False And chkZimmer1Einzel.Checked = False Then Berechnen_Zimmer(1, 14)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkZimmer1Komplett_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkZimmer1Komplett.CheckedChanged
		Berechnen_Zimmer_Komplett(1)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkZimmer1Einzel_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkZimmer1Einzel.CheckedChanged
		Berechnen_Zimmer_Einzel(1)
		Berechnen()
		TotalBerechnen()
	End Sub
	Private Sub chkZimmer1Kinderbett_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkZimmer1Kinderbett.CheckedChanged
		Berechnen_Zimmer_Kinderbett(1)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkBett21_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett21.CheckedChanged
		If chkZimmer2Komplett.Checked = False And chkZimmer2Einzel.Checked = False Then Berechnen_Zimmer(2, 21)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkBett22_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett22.CheckedChanged
		If chkZimmer2Komplett.Checked = False And chkZimmer2Einzel.Checked = False Then Berechnen_Zimmer(2, 22)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkBett23_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett23.CheckedChanged
		If chkZimmer2Komplett.Checked = False And chkZimmer2Einzel.Checked = False Then Berechnen_Zimmer(2, 23)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkBett24_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett24.CheckedChanged
		If chkZimmer2Komplett.Checked = False And chkZimmer2Einzel.Checked = False Then Berechnen_Zimmer(2, 24)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkBett25_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett25.CheckedChanged
		If chkZimmer2Komplett.Checked = False And chkZimmer2Einzel.Checked = False Then Berechnen_Zimmer(2, 25)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkBett26_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett26.CheckedChanged
		If chkZimmer2Komplett.Checked = False And chkZimmer2Einzel.Checked = False Then Berechnen_Zimmer(2, 26)
		Berechnen()
		TotalBerechnen()
	End Sub
	Private Sub chkZimmer2Komplett_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkZimmer2Komplett.CheckedChanged
		Berechnen_Zimmer_Komplett(2)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkZimmer2Einzel_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkZimmer2Einzel.CheckedChanged
		Berechnen_Zimmer_Einzel(2)
		Berechnen()
		TotalBerechnen()
	End Sub
	Private Sub chkZimmer2Kinderbett_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkZimmer2Kinderbett.CheckedChanged
		Berechnen_Zimmer_Kinderbett(2)
		Berechnen()
		TotalBerechnen()
	End Sub

Private Sub chkBett31_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett31.CheckedChanged
		If chkZimmer3Komplett.Checked = False And chkZimmer3Einzel.Checked = False Then Berechnen_Zimmer(3, 31)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkBett32_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett32.CheckedChanged
		If chkZimmer3Komplett.Checked = False And chkZimmer3Einzel.Checked = False Then Berechnen_Zimmer(3, 32)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkBett33_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett33.CheckedChanged
		If chkZimmer3Komplett.Checked = False And chkZimmer3Einzel.Checked = False Then Berechnen_Zimmer(3, 33)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkBett34_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkBett34.CheckedChanged
		If chkZimmer3Komplett.Checked = False And chkZimmer3Einzel.Checked = False Then Berechnen_Zimmer(3, 34)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkZimmer3Komplett_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkZimmer3Komplett.CheckedChanged
		Berechnen_Zimmer_Komplett(3)
		Berechnen()
		TotalBerechnen()
	End Sub

	Private Sub chkZimmer3Einzel_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkZimmer3Einzel.CheckedChanged
		Berechnen_Zimmer_Einzel(3)
		Berechnen()
		TotalBerechnen()
	End Sub
	Private Sub chkZimmer3Kinderbett_CheckedChanged(sender As System.Object, e As System.EventArgs) Handles chkZimmer3Kinderbett.CheckedChanged
		Berechnen_Zimmer_Kinderbett(3)
		Berechnen()
		TotalBerechnen()
	End Sub

Private Sub Berechnen_Zimmer_StartMutation()

	If chkBett11.Checked = True And chkZimmer1Komplett.Checked = False And chkZimmer1Einzel.Checked = False Then Berechnen_Zimmer(1, 11)
	If chkBett12.Checked = True And chkZimmer1Komplett.Checked = False And chkZimmer1Einzel.Checked = False Then Berechnen_Zimmer(1, 12)
	If chkBett13.Checked = True And chkZimmer1Komplett.Checked = False And chkZimmer1Einzel.Checked = False Then Berechnen_Zimmer(1, 13)
	If chkBett14.Checked = True And chkZimmer1Komplett.Checked = False And chkZimmer1Einzel.Checked = False Then Berechnen_Zimmer(1, 14)
	If chkZimmer1Komplett.Checked = True Then Berechnen_Zimmer_Komplett(1)
	If chkZimmer1Einzel.Checked = True Then Berechnen_Zimmer_Einzel(1)
	If chkZimmer1Kinderbett.Checked = True Then Berechnen_Zimmer_Kinderbett(1)

	If chkBett21.Checked = True And chkZimmer2Komplett.Checked = False And chkZimmer2Einzel.Checked = False Then Berechnen_Zimmer(2, 21)
	If chkBett22.Checked = True And chkZimmer2Komplett.Checked = False And chkZimmer2Einzel.Checked = False Then Berechnen_Zimmer(2, 22)
	If chkBett23.Checked = True And chkZimmer2Komplett.Checked = False And chkZimmer2Einzel.Checked = False Then Berechnen_Zimmer(2, 23)
	If chkBett24.Checked = True And chkZimmer2Komplett.Checked = False And chkZimmer2Einzel.Checked = False Then Berechnen_Zimmer(2, 24)
	If chkBett25.Checked = True And chkZimmer2Komplett.Checked = False And chkZimmer2Einzel.Checked = False Then Berechnen_Zimmer(2, 25)
	If chkBett26.Checked = True And chkZimmer2Komplett.Checked = False And chkZimmer2Einzel.Checked = False Then Berechnen_Zimmer(2, 26)
	If chkZimmer2Komplett.Checked = True Then Berechnen_Zimmer_Komplett(2)
	If chkZimmer2Einzel.Checked = True Then Berechnen_Zimmer_Einzel(2)
	If chkZimmer2Kinderbett.Checked = True Then Berechnen_Zimmer_Kinderbett(2)

	If chkBett31.Checked = True And chkZimmer3Komplett.Checked = False And chkZimmer3Einzel.Checked = False Then Berechnen_Zimmer(3, 31)
	If chkBett32.Checked = True And chkZimmer3Komplett.Checked = False And chkZimmer3Einzel.Checked = False Then Berechnen_Zimmer(3, 32)
	If chkBett33.Checked = True And chkZimmer3Komplett.Checked = False And chkZimmer3Einzel.Checked = False Then Berechnen_Zimmer(3, 33)
	If chkBett34.Checked = True And chkZimmer3Komplett.Checked = False And chkZimmer3Einzel.Checked = False Then Berechnen_Zimmer(3, 34)
	If chkZimmer3Komplett.Checked = True Then Berechnen_Zimmer_Komplett(3)
	If chkZimmer3Einzel.Checked = True Then Berechnen_Zimmer_Einzel(3)
	If chkZimmer3Kinderbett.Checked = True Then Berechnen_Zimmer_Kinderbett(3)


End Sub


Private Sub Berechnen_Zimmer(ZimmerNummer As Integer, BettNummer As Integer)

	Select Case BettNummer

	Case 11
		If chkBett11.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If
	Case 12
		If chkBett12.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If
	Case 13
		If chkBett13.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If
	Case 14
		If chkBett14.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If
	Case 21
		If chkBett21.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If

	Case 22
		If chkBett22.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If

	Case 23
		If chkBett23.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If

	Case 24
		If chkBett24.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If

	Case 25
		If chkBett25.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If

	Case 26
		If chkBett26.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If
	Case 31
		If chkBett31.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If
	Case 32
		If chkBett32.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If
	Case 33
		If chkBett33.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If
	Case 34
		If chkBett34.Checked = True Then
			Total_Zimmer(ZimmerNummer, "+")
		Else
			Total_Zimmer(ZimmerNummer, "-")
		End If


	End Select

	'Preis_Bett_4er()


End Sub

Private Sub Total_Zimmer(ZimmerNummer As String, RechenArt As String)

	Select Case ZimmerNummer

	Case 1

		If RechenArt = "+" Then
			Preis_Zimmer1_Total = Preis_Zimmer1_Total + Preis_Bett_4er
			txtTotalZimmer1Schreiben(Preis_Zimmer1_Total)
		Else
			Preis_Zimmer1_Total = Preis_Zimmer1_Total - Preis_Bett_4er
			txtTotalZimmer1Schreiben(Preis_Zimmer1_Total)
		End If

	Case 2

		If RechenArt = "+" Then
			Preis_Zimmer2_Total = Preis_Zimmer2_Total + Preis_Bett_6er
			txtTotalZimmer2Schreiben(Preis_Zimmer2_Total)
		Else
			Preis_Zimmer2_Total = Preis_Zimmer2_Total - Preis_Bett_6er
			txtTotalZimmer2Schreiben(Preis_Zimmer2_Total)
		End If

	Case 3

		If RechenArt = "+" Then
			Preis_Zimmer3_Total = Preis_Zimmer3_Total + Preis_Bett_4er
			txtTotalZimmer3Schreiben(Preis_Zimmer3_Total)
		Else
			Preis_Zimmer3_Total = Preis_Zimmer3_Total - Preis_Bett_4er
			txtTotalZimmer3Schreiben(Preis_Zimmer3_Total)
		End If

	End Select

End Sub

Private Sub Berechnen_Zimmer_Komplett(ZimmerNummer As Integer)

	Select Case ZimmerNummer
	Case 1

		If chkZimmer1Komplett.Checked = True Then
			Preis_Zimmer1_Total = 0
			chkZimmer1Einzel.Enabled = False
			chkZimmer1Einzel.Checked = False
			Zimmer1Komplett_on()
			Preis_Zimmer1_Total = Preis_Zimmer1_Total + Preis_Zimmer1_Komplett
			txtTotalZimmer1Schreiben(Preis_Zimmer1_Total)
		Else
			chkZimmer1Einzel.Enabled = True
			Zimmer1Komplett_off()
			'Preis_Zimmer1_Total = Preis_Zimmer1_Total - Preis_Zimmer1_Komplett
			Preis_Zimmer1_Total = 0
			txtTotalZimmer1Schreiben(Preis_Zimmer1_Total)
		End If

	Case 2

		If chkZimmer2Komplett.Checked = True Then
			Preis_Zimmer2_Total = 0
			chkZimmer2Einzel.Enabled = False
			chkZimmer2Einzel.Checked = False
			Zimmer2Komplett_on()
			Preis_Zimmer2_Total = Preis_Zimmer2_Total + Preis_Zimmer2_Komplett
			txtTotalZimmer2Schreiben(Preis_Zimmer2_Total)
		Else
			chkZimmer2Einzel.Enabled = True
			Zimmer2Komplett_off()
			'Preis_Zimmer2_Total = Preis_Zimmer2_Total - Preis_Zimmer2_Komplett
			Preis_Zimmer2_Total = 0
			txtTotalZimmer2Schreiben(Preis_Zimmer2_Total)
		End If

	Case 3

		If chkZimmer3Komplett.Checked = True Then
			Preis_Zimmer3_Total = 0
			chkZimmer3Einzel.Enabled = False
			chkZimmer3Einzel.Checked = False
			Zimmer3Komplett_on()
			Preis_Zimmer3_Total = Preis_Zimmer3_Total + Preis_Zimmer3_Komplett
			txtTotalZimmer3Schreiben(Preis_Zimmer3_Total)
		Else
			chkZimmer3Einzel.Enabled = True
			Zimmer3Komplett_off()
			'Preis_Zimmer3_Total = Preis_Zimmer3_Total - Preis_Zimmer3_Komplett
			Preis_Zimmer3_Total = 0
			txtTotalZimmer3Schreiben(Preis_Zimmer3_Total)
		End If


	End Select



End Sub

Private Sub Berechnen_Zimmer_Einzel(ZimmerNummer As Integer)


	Select Case ZimmerNummer

	Case 1

		If chkZimmer1Einzel.Checked = True Then
			Preis_Zimmer1_Total = 0
			chkZimmer1Komplett.Enabled = False
			chkZimmer1Komplett.Checked = False
			Zimmer1Komplett_on()
			Preis_Zimmer1_Total = Preis_Zimmer1_Total + Preis_Zimmer1_Einzel
			txtTotalZimmer1Schreiben(Preis_Zimmer1_Total)
		Else
			chkZimmer1Komplett.Enabled = True
			Zimmer1Komplett_off()
			'Preis_Zimmer1_Total = Preis_Zimmer1_Total - Preis_Zimmer1_Einzel
			Preis_Zimmer1_Total = 0
			txtTotalZimmer1Schreiben(Preis_Zimmer1_Total)
		End If

	Case 2

		If chkZimmer2Einzel.Checked = True Then
			Preis_Zimmer2_Total = 0
			chkZimmer2Komplett.Enabled = False
			chkZimmer2Komplett.Checked = False
			Zimmer2Komplett_on()
			Preis_Zimmer2_Total = Preis_Zimmer2_Total + Preis_Zimmer2_Einzel
			txtTotalZimmer2Schreiben(Preis_Zimmer2_Total)
		Else
			chkZimmer2Komplett.Enabled = True
			Zimmer2Komplett_off()
			'Preis_Zimmer2_Total = Preis_Zimmer2_Total - Preis_Zimmer2_Einzel
			Preis_Zimmer2_Total = 0
			txtTotalZimmer2Schreiben(Preis_Zimmer2_Total)
		End If


	Case 3

		If chkZimmer3Einzel.Checked = True Then
			Preis_Zimmer3_Total = 0
			chkZimmer3Komplett.Enabled = False
			chkZimmer3Komplett.Checked = False
			Zimmer3Komplett_on()
			Preis_Zimmer3_Total = Preis_Zimmer3_Total + Preis_Zimmer3_Einzel
			txtTotalZimmer3Schreiben(Preis_Zimmer3_Total)
		Else
			chkZimmer3Komplett.Enabled = True
			Zimmer3Komplett_off()
			'Preis_Zimmer3_Total = Preis_Zimmer3_Total - Preis_Zimmer3_Einzel
			Preis_Zimmer3_Total = 0
			txtTotalZimmer3Schreiben(Preis_Zimmer3_Total)
		End If


	End Select


End Sub

Private Sub Berechnen_Zimmer_Kinderbett(ZimmerNummer As Integer)

	Select Case ZimmerNummer

	Case 1

		If chkZimmer1Kinderbett.Checked = True Then
			Preis_Zimmer1_Total = Preis_Zimmer1_Total + Preis_Zimmer1_Kinderbett
			txtTotalZimmer1Schreiben(Preis_Zimmer1_Total)
		Else
			Preis_Zimmer1_Total = Preis_Zimmer1_Total - Preis_Zimmer1_Kinderbett
			txtTotalZimmer1Schreiben(Preis_Zimmer1_Total)
		End If

	Case 2

		If chkZimmer2Kinderbett.Checked = True Then
			Preis_Zimmer2_Total = Preis_Zimmer2_Total + Preis_Zimmer2_Kinderbett
			txtTotalZimmer2Schreiben(Preis_Zimmer2_Total)
		Else
			Preis_Zimmer2_Total = Preis_Zimmer2_Total - Preis_Zimmer2_Kinderbett
			txtTotalZimmer2Schreiben(Preis_Zimmer2_Total)
		End If


	Case 3

		If chkZimmer3Kinderbett.Checked = True Then
			Preis_Zimmer3_Total = Preis_Zimmer3_Total + Preis_Zimmer3_Kinderbett
			txtTotalZimmer3Schreiben(Preis_Zimmer3_Total)
		Else
			Preis_Zimmer3_Total = Preis_Zimmer3_Total - Preis_Zimmer3_Kinderbett
			txtTotalZimmer3Schreiben(Preis_Zimmer3_Total)
		End If


	End Select


End Sub


Private Sub txtTotalZimmer1Schreiben(p As Double)

	txtTotalZimmer1.Text = String.Format("{0:N}", p)

End Sub

Private Sub txtTotalZimmer2Schreiben(p As Double)

	txtTotalZimmer2.Text = String.Format("{0:N}", p)

End Sub

Private Sub txtTotalZimmer3Schreiben(p As Double)

	txtTotalZimmer3.Text = String.Format("{0:N}", p)

End Sub


Private Sub Zimmer1Komplett_on()

		chkBett11.Enabled = False
		chkBett11.Checked = True
		chkBett12.Enabled = False
		chkBett12.Checked = True
		chkBett13.Enabled = False
		chkBett13.Checked = True
		chkBett14.Enabled = False
		chkBett14.Checked = True

End Sub

Private Sub Zimmer1Komplett_off()

		chkBett11.Enabled = True
		chkBett11.Checked = False
		chkBett12.Enabled = True
		chkBett12.Checked = False
		chkBett13.Enabled = True
		chkBett13.Checked = False
		chkBett14.Enabled = True
		chkBett14.Checked = False

End Sub

Private Sub Zimmer2Komplett_on()

		chkBett21.Enabled = False
		chkBett21.Checked = True
		chkBett22.Enabled = False
		chkBett22.Checked = True
		chkBett23.Enabled = False
		chkBett23.Checked = True
		chkBett24.Enabled = False
		chkBett24.Checked = True
		chkBett25.Enabled = False
		chkBett25.Checked = True
		chkBett26.Enabled = False
		chkBett26.Checked = True

End Sub

Private Sub Zimmer2Komplett_off()

		chkBett21.Enabled = True
		chkBett21.Checked = False
		chkBett22.Enabled = True
		chkBett22.Checked = False
		chkBett23.Enabled = True
		chkBett23.Checked = False
		chkBett24.Enabled = True
		chkBett24.Checked = False
		chkBett25.Enabled = True
		chkBett25.Checked = False
		chkBett26.Enabled = True
		chkBett26.Checked = False

End Sub

Private Sub Zimmer3Komplett_on()

		chkBett31.Enabled = False
		chkBett31.Checked = True
		chkBett32.Enabled = False
		chkBett32.Checked = True
		chkBett33.Enabled = False
		chkBett33.Checked = True
		chkBett34.Enabled = False
		chkBett34.Checked = True

End Sub

Private Sub Zimmer3Komplett_off()

		chkBett31.Enabled = True
		chkBett31.Checked = False
		chkBett32.Enabled = True
		chkBett32.Checked = False
		chkBett33.Enabled = True
		chkBett33.Checked = False
		chkBett34.Enabled = True
		chkBett34.Checked = False

End Sub

Private Sub btnBestaetigungOhneMeldeschein_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBestaetigungOhneMeldeschein.Click


		If StartFlag = "StartBooking" Then

			SpeichernBookingDaten()
			StartSpeichernBegleitpersonen()
			MsgBox("Buchungsdaten gespeichert", MsgBoxStyle.Exclamation)

		Else

			MutierenBookingDaten()
			StartMutierenBegleitpersonen()
			MsgBox("Buchungsdaten mutiert", MsgBoxStyle.Exclamation)

		End If


		Dim count As Integer = 1
		Do While count <= AnzahlAusdruckeBuchungsbestaetigung()
			Drucken_Bestätigung_ohneMeldeschein()
			count += 1
		Loop
		count = 1

		Me.Close()

	End Sub

	Private Sub btnBestaetigungMitMeldeschein_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnBestaetigungMitMeldeschein.Click

		If StartFlag = "StartBooking" Then

			SpeichernBookingDaten()
			StartSpeichernBegleitpersonen()
			MsgBox("Buchungsdaten gespeichert", MsgBoxStyle.Exclamation)

		Else

			MutierenBookingDaten()
			StartMutierenBegleitpersonen()
			MsgBox("Buchungsdaten mutiert", MsgBoxStyle.Exclamation)

		End If


		Dim count As Integer = 1
		Do While count <= AnzahlAusdruckeBuchungsbestaetigung()
			Drucken_Bestätigung_mitMeldeschein()
			count += 1
		Loop
		count = 1

		Me.Close()

	End Sub

	Private Sub btnQuittung_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnQuittung.Click


		If MsgBox("Wollen Sie wirklich diese Buchung abrechnen?", vbOKCancel + MsgBoxStyle.Critical) = vbCancel Then
			Exit Sub
		End If


		If StartFlag = "StartBooking" Then

			SpeichernBookingDaten()
			StartSpeichernBegleitpersonen()
			MsgBox("Buchungsdaten gespeichert", MsgBoxStyle.Exclamation)

		Else

			MutierenBookingDaten()
			StartMutierenBegleitpersonen()
			MsgBox("Buchungsdaten mutiert", MsgBoxStyle.Exclamation)

		End If

		AktivFlagSetzen()

		Dim count As Integer = 1
		Do While count <= AnzahlAusdruckeQuittung()
			Drucken_Quittung()
			count += 1
		Loop
		count = 1

		Me.Close()


	End Sub

  Public Sub Drucken_Bestätigung_ohneMeldeschein()

		Dim cmd As New OleDbCommand("Select * from tbAdress where AdressId = " & Public_AdressId & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet()

		da.Fill(ds, "tbAdress")

		Dim oWord As Word.Application
		Dim oDoc As Word.Document

		SprachCode = ds.Tables("tbAdress").Rows(0).Item(14)

		oWord = CreateObject("Word.Application")
		oWord.Visible = True

		oDoc = oWord.Documents.Add("c:\Formulare\Buchung_ohneMeldeschein.dot")

		oDoc.Bookmarks.Item("tmDatum").Range.Text = System.DateTime.Today
		oDoc.Bookmarks.Item("tmBuchungsnummer").Range.Text = Me.lblBookingId.Text
		oDoc.Bookmarks.Item("tmPlatznummer").Range.Text = ""


		oDoc.Bookmarks.Item("tmName").Range.Text = ds.Tables("tbAdress").Rows(0).Item(1)
		oDoc.Bookmarks.Item("tmVorname").Range.Text = ds.Tables("tbAdress").Rows(0).Item(2)
		oDoc.Bookmarks.Item("tmAdresse").Range.Text = ds.Tables("tbAdress").Rows(0).Item(3)
		oDoc.Bookmarks.Item("tmPLZOrt").Range.Text = ds.Tables("tbAdress").Rows(0).Item(4) & " " & ds.Tables("tbAdress").Rows(0).Item(5)


		oDoc.Bookmarks.Item("tmAnreise").Range.Text = DateTimePicker1.Value
		oDoc.Bookmarks.Item("tmAbreise").Range.Text = DateTimePicker2.Value
		oDoc.Bookmarks.Item("tmTotalNaechte").Range.Text = Me.txtNaechte.Text
		oDoc.Bookmarks.Item("tmKennzeichen").Range.Text = Me.txtAutoKennZeichen.Text


		'-----------------------------------------------------------------------------------


		oDoc.Bookmarks.Item("textDokTitel").Range.Text = Form1.TitelErmitteln(1, SprachCode)
		oDoc.Bookmarks.Item("textBegleitperson").Range.Text = Form1.TitelErmitteln(15, SprachCode)

		oDoc.Bookmarks.Item("textName").Range.Text = Form1.TitelErmitteln(18, SprachCode)
		oDoc.Bookmarks.Item("textVorname").Range.Text = Form1.TitelErmitteln(19, SprachCode)


		oDoc.Bookmarks.Item("textName1").Range.Text = Form1.TitelErmitteln(18, SprachCode)
		oDoc.Bookmarks.Item("textVorname1").Range.Text = Form1.TitelErmitteln(19, SprachCode)

		oDoc.Bookmarks.Item("textGeburtsdatum1").Range.Text = Form1.TitelErmitteln(9, SprachCode)
		oDoc.Bookmarks.Item("textLand1").Range.Text = Form1.TitelErmitteln(20, SprachCode)


		oDoc.Bookmarks.Item("textAdresse").Range.Text = Form1.TitelErmitteln(24, SprachCode)
		oDoc.Bookmarks.Item("textPLZOrt").Range.Text = Form1.TitelErmitteln(4, SprachCode)

		oDoc.Bookmarks.Item("textAnreise").Range.Text = Form1.TitelErmitteln(5, SprachCode)
		oDoc.Bookmarks.Item("textAbreise").Range.Text = Form1.TitelErmitteln(6, SprachCode)

		oDoc.Bookmarks.Item("textTotalNächte").Range.Text = Form1.TitelErmitteln(21, SprachCode)
		oDoc.Bookmarks.Item("textKennzeichen").Range.Text = Form1.TitelErmitteln(3, SprachCode)



		oDoc.Bookmarks.Item("textAnzahl").Range.Text = Form1.TitelErmitteln(16, SprachCode)
		oDoc.Bookmarks.Item("textBezeichnung").Range.Text = Form1.TitelErmitteln(10, SprachCode)
		oDoc.Bookmarks.Item("textEinzelpreis").Range.Text = Form1.TitelErmitteln(11, SprachCode)
		oDoc.Bookmarks.Item("textTotal").Range.Text = Form1.TitelErmitteln(23, SprachCode)
		oDoc.Bookmarks.Item("textZwischenTotal").Range.Text = Form1.TitelErmitteln(25, SprachCode)
		oDoc.Bookmarks.Item("textMwst").Range.Text = Form1.TitelErmitteln(17, SprachCode)
		oDoc.Bookmarks.Item("textDatum").Range.Text = Form1.TitelErmitteln(26, SprachCode)
		oDoc.Bookmarks.Item("textBuchungsnummer").Range.Text = Form1.TitelErmitteln(2, SprachCode)
		oDoc.Bookmarks.Item("textPlatznummer").Range.Text = Form1.TitelErmitteln(27, SprachCode)


		'-----------------------------------------------------------------------------------

		If txtNachName1.Text <> "" Then
			oDoc.Bookmarks.Item("bpName1").Range.Text = txtNachName1.Text
			oDoc.Bookmarks.Item("bpVorname1").Range.Text = txtVorname1.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum1").Range.Text = txtGeburtsdatum1.Text
			oDoc.Bookmarks.Item("bpLand1").Range.Text = cmbLand1.Text
		End If


		If txtNachName2.Text <> "" Then
			oDoc.Bookmarks.Item("bpName2").Range.Text = txtNachName2.Text
			oDoc.Bookmarks.Item("bpVorname2").Range.Text = txtVorname2.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum2").Range.Text = txtGeburtsdatum2.Text
			oDoc.Bookmarks.Item("bpLand2").Range.Text = cmbLand2.Text
		End If

		If txtNachName3.Text <> "" Then
			oDoc.Bookmarks.Item("bpName3").Range.Text = txtNachName3.Text
			oDoc.Bookmarks.Item("bpVorname3").Range.Text = txtVorname3.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum3").Range.Text = txtGeburtsdatum3.Text
			oDoc.Bookmarks.Item("bpLand3").Range.Text = cmbLand3.Text
		End If


		If txtNachName4.Text <> "" Then
			oDoc.Bookmarks.Item("bpName4").Range.Text = txtNachName4.Text
			oDoc.Bookmarks.Item("bpVorname4").Range.Text = txtVorname4.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum4").Range.Text = txtGeburtsdatum4.Text
			oDoc.Bookmarks.Item("bpLand4").Range.Text = cmbLand4.Text
		End If

		If txtNachName5.Text <> "" Then
			oDoc.Bookmarks.Item("bpName5").Range.Text = txtNachName5.Text
			oDoc.Bookmarks.Item("bpVorname5").Range.Text = txtVorname5.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum5").Range.Text = txtGeburtsdatum5.Text
			oDoc.Bookmarks.Item("bpLand5").Range.Text = cmbLand5.Text
		End If

		If txtNachName6.Text <> "" Then
			oDoc.Bookmarks.Item("bpName6").Range.Text = txtNachName6.Text
			oDoc.Bookmarks.Item("bpVorname6").Range.Text = txtVorname6.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum6").Range.Text = txtGeburtsdatum6.Text
			oDoc.Bookmarks.Item("bpLand6").Range.Text = cmbLand6.Text
		End If

		If txtNachName7.Text <> "" Then
			oDoc.Bookmarks.Item("bpName7").Range.Text = txtNachName7.Text
			oDoc.Bookmarks.Item("bpVorname7").Range.Text = txtVorname7.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum7").Range.Text = txtGeburtsdatum7.Text
			oDoc.Bookmarks.Item("bpLand7").Range.Text = cmbLand7.Text
		End If

		If txtNachName8.Text <> "" Then
			oDoc.Bookmarks.Item("bpName8").Range.Text = txtNachName8.Text
			oDoc.Bookmarks.Item("bpVorname8").Range.Text = txtVorname8.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum8").Range.Text = txtGeburtsdatum8.Text
			oDoc.Bookmarks.Item("bpLand8").Range.Text = cmbLand8.Text
		End If

		If txtNachName9.Text <> "" Then
			oDoc.Bookmarks.Item("bpName9").Range.Text = txtNachName9.Text
			oDoc.Bookmarks.Item("bpVorname9").Range.Text = txtVorname9.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum9").Range.Text = txtGeburtsdatum9.Text
			oDoc.Bookmarks.Item("bpLand9").Range.Text = cmbLand9.Text
		End If


		If txtNachName10.Text <> "" Then
			oDoc.Bookmarks.Item("bpName10").Range.Text = txtNachName10.Text
			oDoc.Bookmarks.Item("bpVorname10").Range.Text = txtVorname10.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum10").Range.Text = txtGeburtsdatum10.Text
			oDoc.Bookmarks.Item("bpLand10").Range.Text = cmbLand10.Text
		End If

		oDoc.Bookmarks.Item("tmZwischentotal").Range.Text = txtZwischenTotal.Text



		Dim textPart1 As Word.Range = oDoc.Bookmarks.Item("Anzahl").Range
		Dim textPart2 As Word.Range = oDoc.Bookmarks.Item("Bezeichnung").Range
		Dim textPart3 As Word.Range = oDoc.Bookmarks.Item("Einzelpreis").Range
		Dim textPart4 As Word.Range = oDoc.Bookmarks.Item("Totalpreis").Range

		Dim textPartMwst As Word.Range = oDoc.Bookmarks.Item("Mwst").Range
		Dim textPartTotal As Word.Range = oDoc.Bookmarks.Item("Total").Range



		textPart1.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
		textPart2.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
		textPart3.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
		textPart4.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

		textPartMwst.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
		textPartTotal.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight


		Dim CheckString As String = "0"


		Dim AnzahlNaechte As Integer = CInt(txtNaechte.Text)


		'------------------------------------------------------------------------------------------------


		Dim PrintEinzelZimmer1 As Boolean = False
		Dim PrintEinzelZimmer2 As Boolean = False
		Dim PrintEinzelZimmer3 As Boolean = False


		If chkZimmer1Komplett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(32, SprachCode) & " Nr. 1")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Komplett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Komplett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer1 = True

		End If


		If chkZimmer1Einzel.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(30, SprachCode) & " Nr. 1")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Einzel))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Einzel * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer1 = True

		End If


		If PrintEinzelZimmer1 = False Then

			If chkBett11.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.11")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett12.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.12")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett13.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.13")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett14.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.14")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If

		End If



		If chkZimmer1Kinderbett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(34, SprachCode) & " Nr. 1")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Kinderbett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Kinderbett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If


		'------------------------------------------------------------------------------------------------



		If chkZimmer2Komplett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(33, SprachCode) & " Nr. 2")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Komplett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Komplett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer2 = True

		End If


		If chkZimmer2Einzel.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(31, SprachCode) & " Nr. 2")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Einzel))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Einzel * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer2 = True

		End If



		If PrintEinzelZimmer2 = False Then

			If chkBett21.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.21")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett22.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.22")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett23.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.23")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett24.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.24")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett25.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.25")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett26.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.26")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If

		End If


		If chkZimmer2Kinderbett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(34, SprachCode) & " Nr. 2")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Kinderbett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Kinderbett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If



		'-----------------------------------------------------------------------------------------

		If chkZimmer3Komplett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(32, SprachCode) & " Nr. 3")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Komplett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Komplett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer3 = True

		End If


		If chkZimmer3Einzel.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(30, SprachCode) & " Nr. 3")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Einzel))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Einzel * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer3 = True

		End If

		If PrintEinzelZimmer3 = False Then

			If chkBett31.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.31")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett32.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.32")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett33.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.33")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett34.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.34")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If

		End If

		If chkZimmer3Kinderbett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(34, SprachCode) & " Nr. 3")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Kinderbett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Kinderbett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'----------------------------------------------------------------------------------------------

		'------- AUTO -----

		If Me.cmb04.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb04.Text & "/(" & AnzahlNaechte & ")")
			textPart2.InsertAfter(TexteLesen(4))
			textPart3.InsertAfter(Me.txtPreis04.Text)
			textPart4.InsertAfter(Me.total04.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'------- MOTORRAD -----

		If Me.cmb03.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb03.Text & "/(" & AnzahlNaechte & ")")
			textPart2.InsertAfter(TexteLesen(6))
			textPart3.InsertAfter(Me.txtPreis03.Text)
			textPart4.InsertAfter(Me.total03.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If


		'------- HANDTüCHER -----

		If Me.cmb05.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb05.Text)
			textPart2.InsertAfter(TexteLesen(25))
			textPart3.InsertAfter(Me.txtPreis05.Text)
			textPart4.InsertAfter(Me.total05.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'------- FRÜHSTÜCK -----

		If Me.cmb06.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb06.Text)
			textPart2.InsertAfter(TexteLesen(26))
			textPart3.InsertAfter(Me.txtPreis06.Text)
			textPart4.InsertAfter(Me.total06.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'------- Küche -----

		If Me.cmb07.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb07.Text)
			textPart2.InsertAfter(TexteLesen(27))
			textPart3.InsertAfter(Me.txtPreis07.Text)
			textPart4.InsertAfter(Me.total07.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If


		'------- SONSTIGES -----

		If Me.txtSonstigesText.Text <> vbNullString Then

			textPart2.InsertAfter(Me.txtSonstigesText.Text)
			textPart4.InsertAfter(Me.txtSonstigesBetrag.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'------- TAXE -----

		If Me.txtTaxe.Text <> vbNullString Then

			textPart2.InsertAfter(TexteLesen(17))
			textPart4.InsertAfter(Me.txtTaxe.Text)

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If


		'------- RABATT -----

		Dim RabattWert As String
		Dim RabattArt As String
		Dim BezeichnungRabatt As String

		If Me.txtRabatt.Text <> vbNullString Then


			If txtRabattProzent.Checked = True Then
				RabattArt = "%"
				Dim G1 As Double = CDbl(Me.Label38.Text)
				RabattWert = "- " & G1.ToString("N2")
				BezeichnungRabatt = "Rabatt" & " " & Me.txtRabatt.Text & " " & RabattArt
			Else

				Dim H1 As Double = CDbl(Me.txtRabatt.Text)
				RabattWert = "- " & H1.ToString("N2")
				BezeichnungRabatt = "Rabatt"

				If txtRabattCHF.Checked = True Then
					RabattArt = ""
				Else
					RabattArt = ""
				End If
			End If


			'textPart2.InsertAfter(BezeichnungRabatt)
			'textPart4.InsertAfter(RabattWert)

			oDoc.Bookmarks.Item("tmRabattFeld").Range.Text = BezeichnungRabatt
			oDoc.Bookmarks.Item("tmRabattWert").Range.Text = RabattWert

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		textPartMwst.InsertAfter(Me.txtMwSt.Text)
		textPartTotal.InsertAfter(Me.txtTotalRechnung.Text)

	End Sub


 Public Sub Drucken_Bestätigung_mitMeldeschein()

		Dim AnzahlTaxPflichig As Integer = 0

		Dim cmd As New OleDbCommand("Select * from tbAdress where AdressId = " & Public_AdressId & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet()

		da.Fill(ds, "tbAdress")

		Dim oWord As Word.Application
		Dim oDoc As Word.Document
		Dim StrAusweisArt As String
		SprachCode = ds.Tables("tbAdress").Rows(0).Item(14)

		oWord = CreateObject("Word.Application")
		oWord.Visible = True

		oDoc = oWord.Documents.Add("c:\Formulare\Buchung_mitMeldeschein.dot")


		oDoc.Bookmarks.Item("tmNameMS").Range.Text = ds.Tables("tbAdress").Rows(0).Item(1)
		oDoc.Bookmarks.Item("tmVornameMS").Range.Text = ds.Tables("tbAdress").Rows(0).Item(2)
		oDoc.Bookmarks.Item("tmAdresseMS").Range.Text = ds.Tables("tbAdress").Rows(0).Item(3)
		oDoc.Bookmarks.Item("tmLandPLZOrtMS").Range.Text = ds.Tables("tbAdress").Rows(0).Item(4) & " " & ds.Tables("tbAdress").Rows(0).Item(5)
		oDoc.Bookmarks.Item("tmGeburtsdatumMS").Range.Text = ds.Tables("tbAdress").Rows(0).Item(8)
		oDoc.Bookmarks.Item("tmNationMS").Range.Text = ds.Tables("tbAdress").Rows(0).Item(7)

		AnzahlTaxPflichig = AnzahlTaxPflichig + 1

		oDoc.Bookmarks.Item("tmAusweisArtMS").Range.Text = ds.Tables("tbAdress").Rows(0).Item(11)
		oDoc.Bookmarks.Item("tmAusweisNummerMS").Range.Text = ds.Tables("tbAdress").Rows(0).Item(10)


		oDoc.Bookmarks.Item("tmDatumMS").Range.Text = System.DateTime.Today

		oDoc.Bookmarks.Item("tmAnkunftsdatumMS").Range.Text = DateTimePicker1.Value
		oDoc.Bookmarks.Item("tmGeplantesAbreisedatumMS").Range.Text = DateTimePicker2.Value
		oDoc.Bookmarks.Item("tmAbreisedatumMS").Range.Text = DateTimePicker2.Value




		oDoc.Bookmarks.Item("tmDatum").Range.Text = System.DateTime.Today
		oDoc.Bookmarks.Item("tmBuchungsnummer").Range.Text = Me.lblBookingId.Text
		oDoc.Bookmarks.Item("tmPlatznummer").Range.Text = ""


		oDoc.Bookmarks.Item("tmName").Range.Text = ds.Tables("tbAdress").Rows(0).Item(1)
		oDoc.Bookmarks.Item("tmVorname").Range.Text = ds.Tables("tbAdress").Rows(0).Item(2)
		oDoc.Bookmarks.Item("tmAdresse").Range.Text = ds.Tables("tbAdress").Rows(0).Item(3)
		oDoc.Bookmarks.Item("tmPLZOrt").Range.Text = ds.Tables("tbAdress").Rows(0).Item(4) & " " & ds.Tables("tbAdress").Rows(0).Item(5)


		oDoc.Bookmarks.Item("tmAnreise").Range.Text = DateTimePicker1.Value
		oDoc.Bookmarks.Item("tmAbreise").Range.Text = DateTimePicker2.Value
		oDoc.Bookmarks.Item("tmTotalNaechte").Range.Text = Me.txtNaechte.Text
		oDoc.Bookmarks.Item("tmKennzeichen").Range.Text = Me.txtAutoKennZeichen.Text



		'-----------------------------------------------------------------------------------


		oDoc.Bookmarks.Item("textDokTitel").Range.Text = Form1.TitelErmitteln(1, SprachCode)
		oDoc.Bookmarks.Item("textBegleitperson").Range.Text = Form1.TitelErmitteln(15, SprachCode)

		oDoc.Bookmarks.Item("textName").Range.Text = Form1.TitelErmitteln(18, SprachCode)
		oDoc.Bookmarks.Item("textVorname").Range.Text = Form1.TitelErmitteln(19, SprachCode)


		oDoc.Bookmarks.Item("textName1").Range.Text = Form1.TitelErmitteln(18, SprachCode)
		oDoc.Bookmarks.Item("textVorname1").Range.Text = Form1.TitelErmitteln(19, SprachCode)

		oDoc.Bookmarks.Item("textGeburtsdatum1").Range.Text = Form1.TitelErmitteln(9, SprachCode)
		oDoc.Bookmarks.Item("textLand1").Range.Text = Form1.TitelErmitteln(20, SprachCode)


		oDoc.Bookmarks.Item("textAdresse").Range.Text = Form1.TitelErmitteln(24, SprachCode)
		oDoc.Bookmarks.Item("textPLZOrt").Range.Text = Form1.TitelErmitteln(4, SprachCode)

		oDoc.Bookmarks.Item("textAnreise").Range.Text = Form1.TitelErmitteln(5, SprachCode)
		oDoc.Bookmarks.Item("textAbreise").Range.Text = Form1.TitelErmitteln(6, SprachCode)

		oDoc.Bookmarks.Item("textTotalNächte").Range.Text = Form1.TitelErmitteln(21, SprachCode)
		oDoc.Bookmarks.Item("textKennzeichen").Range.Text = Form1.TitelErmitteln(3, SprachCode)



		oDoc.Bookmarks.Item("textAnzahl").Range.Text = Form1.TitelErmitteln(16, SprachCode)
		oDoc.Bookmarks.Item("textBezeichnung").Range.Text = Form1.TitelErmitteln(10, SprachCode)
		oDoc.Bookmarks.Item("textEinzelpreis").Range.Text = Form1.TitelErmitteln(11, SprachCode)
		oDoc.Bookmarks.Item("textTotal").Range.Text = Form1.TitelErmitteln(23, SprachCode)
		oDoc.Bookmarks.Item("textZwischenTotal").Range.Text = Form1.TitelErmitteln(25, SprachCode)
		oDoc.Bookmarks.Item("textMwst").Range.Text = Form1.TitelErmitteln(17, SprachCode)
		oDoc.Bookmarks.Item("textDatum").Range.Text = Form1.TitelErmitteln(26, SprachCode)
		oDoc.Bookmarks.Item("textBuchungsnummer").Range.Text = Form1.TitelErmitteln(2, SprachCode)
		oDoc.Bookmarks.Item("textPlatznummer").Range.Text = Form1.TitelErmitteln(27, SprachCode)


		'-----------------------------------------------------------------------------------

		If txtNachName1.Text <> "" Then
			oDoc.Bookmarks.Item("bpName1").Range.Text = txtNachName1.Text
			oDoc.Bookmarks.Item("bpVorname1").Range.Text = txtVorname1.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum1").Range.Text = txtGeburtsdatum1.Text
			oDoc.Bookmarks.Item("bpLand1").Range.Text = cmbLand1.Text

			oDoc.Bookmarks.Item("xbpName1").Range.Text = txtNachName1.Text
			oDoc.Bookmarks.Item("xbpVorname1").Range.Text = txtVorname1.Text
			oDoc.Bookmarks.Item("xbpGeburtsdatum1").Range.Text = txtGeburtsdatum1.Text
			oDoc.Bookmarks.Item("xbpLand1").Range.Text = cmbLand1.Text

			AnzahlTaxPflichig = AnzahlTaxPflichig + 1

		End If


		If txtNachName2.Text <> "" Then
			oDoc.Bookmarks.Item("bpName2").Range.Text = txtNachName2.Text
			oDoc.Bookmarks.Item("bpVorname2").Range.Text = txtVorname2.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum2").Range.Text = txtGeburtsdatum2.Text
			oDoc.Bookmarks.Item("bpLand2").Range.Text = cmbLand2.Text

			oDoc.Bookmarks.Item("xbpName2").Range.Text = txtNachName2.Text
			oDoc.Bookmarks.Item("xbpVorname2").Range.Text = txtVorname2.Text
			oDoc.Bookmarks.Item("xbpGeburtsdatum2").Range.Text = txtGeburtsdatum2.Text
			oDoc.Bookmarks.Item("xbpLand2").Range.Text = cmbLand2.Text

			AnzahlTaxPflichig = AnzahlTaxPflichig + 1

		End If

		If txtNachName3.Text <> "" Then
			oDoc.Bookmarks.Item("bpName3").Range.Text = txtNachName3.Text
			oDoc.Bookmarks.Item("bpVorname3").Range.Text = txtVorname3.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum3").Range.Text = txtGeburtsdatum3.Text
			oDoc.Bookmarks.Item("bpLand3").Range.Text = cmbLand3.Text

			oDoc.Bookmarks.Item("xbpName3").Range.Text = txtNachName3.Text
			oDoc.Bookmarks.Item("xbpVorname3").Range.Text = txtVorname3.Text
			oDoc.Bookmarks.Item("xbpGeburtsdatum3").Range.Text = txtGeburtsdatum3.Text
			oDoc.Bookmarks.Item("xbpLand3").Range.Text = cmbLand3.Text

			AnzahlTaxPflichig = AnzahlTaxPflichig + 1

		End If


		If txtNachName4.Text <> "" Then
			oDoc.Bookmarks.Item("bpName4").Range.Text = txtNachName4.Text
			oDoc.Bookmarks.Item("bpVorname4").Range.Text = txtVorname4.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum4").Range.Text = txtGeburtsdatum4.Text
			oDoc.Bookmarks.Item("bpLand4").Range.Text = cmbLand4.Text

			oDoc.Bookmarks.Item("xbpName4").Range.Text = txtNachName4.Text
			oDoc.Bookmarks.Item("xbpVorname4").Range.Text = txtVorname4.Text
			oDoc.Bookmarks.Item("xbpGeburtsdatum4").Range.Text = txtGeburtsdatum4.Text
			oDoc.Bookmarks.Item("xbpLand4").Range.Text = cmbLand4.Text

			AnzahlTaxPflichig = AnzahlTaxPflichig + 1

		End If

		If txtNachName5.Text <> "" Then
			oDoc.Bookmarks.Item("bpName5").Range.Text = txtNachName5.Text
			oDoc.Bookmarks.Item("bpVorname5").Range.Text = txtVorname5.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum5").Range.Text = txtGeburtsdatum5.Text
			oDoc.Bookmarks.Item("bpLand5").Range.Text = cmbLand5.Text

			oDoc.Bookmarks.Item("xbpName5").Range.Text = txtNachName5.Text
			oDoc.Bookmarks.Item("xbpVorname5").Range.Text = txtVorname5.Text
			oDoc.Bookmarks.Item("xbpGeburtsdatum5").Range.Text = txtGeburtsdatum5.Text
			oDoc.Bookmarks.Item("xbpLand5").Range.Text = cmbLand5.Text

			AnzahlTaxPflichig = AnzahlTaxPflichig + 1

		End If

		If txtNachName6.Text <> "" Then
			oDoc.Bookmarks.Item("bpName6").Range.Text = txtNachName6.Text
			oDoc.Bookmarks.Item("bpVorname6").Range.Text = txtVorname6.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum6").Range.Text = txtGeburtsdatum6.Text
			oDoc.Bookmarks.Item("bpLand6").Range.Text = cmbLand6.Text

			oDoc.Bookmarks.Item("xbpName6").Range.Text = txtNachName6.Text
			oDoc.Bookmarks.Item("xbpVorname6").Range.Text = txtVorname6.Text
			oDoc.Bookmarks.Item("xbpGeburtsdatum6").Range.Text = txtGeburtsdatum6.Text
			oDoc.Bookmarks.Item("xbpLand6").Range.Text = cmbLand6.Text

			AnzahlTaxPflichig = AnzahlTaxPflichig + 1

		End If

		If txtNachName7.Text <> "" Then
			oDoc.Bookmarks.Item("bpName7").Range.Text = txtNachName7.Text
			oDoc.Bookmarks.Item("bpVorname7").Range.Text = txtVorname7.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum7").Range.Text = txtGeburtsdatum7.Text
			oDoc.Bookmarks.Item("bpLand7").Range.Text = cmbLand7.Text

			oDoc.Bookmarks.Item("xbpName7").Range.Text = txtNachName7.Text
			oDoc.Bookmarks.Item("xbpVorname7").Range.Text = txtVorname7.Text
			oDoc.Bookmarks.Item("xbpGeburtsdatum7").Range.Text = txtGeburtsdatum7.Text
			oDoc.Bookmarks.Item("xbpLand7").Range.Text = cmbLand7.Text

			AnzahlTaxPflichig = AnzahlTaxPflichig + 1

		End If

		If txtNachName8.Text <> "" Then
			oDoc.Bookmarks.Item("bpName8").Range.Text = txtNachName8.Text
			oDoc.Bookmarks.Item("bpVorname8").Range.Text = txtVorname8.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum8").Range.Text = txtGeburtsdatum8.Text
			oDoc.Bookmarks.Item("bpLand8").Range.Text = cmbLand8.Text

			oDoc.Bookmarks.Item("xbpName8").Range.Text = txtNachName8.Text
			oDoc.Bookmarks.Item("xbpVorname8").Range.Text = txtVorname8.Text
			oDoc.Bookmarks.Item("xbpGeburtsdatum8").Range.Text = txtGeburtsdatum8.Text
			oDoc.Bookmarks.Item("xbpLand8").Range.Text = cmbLand8.Text

			AnzahlTaxPflichig = AnzahlTaxPflichig + 1

		End If

		If txtNachName9.Text <> "" Then
			oDoc.Bookmarks.Item("bpName9").Range.Text = txtNachName9.Text
			oDoc.Bookmarks.Item("bpVorname9").Range.Text = txtVorname9.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum9").Range.Text = txtGeburtsdatum9.Text
			oDoc.Bookmarks.Item("bpLand9").Range.Text = cmbLand9.Text

			oDoc.Bookmarks.Item("xbpName9").Range.Text = txtNachName9.Text
			oDoc.Bookmarks.Item("xbpVorname9").Range.Text = txtVorname9.Text
			oDoc.Bookmarks.Item("xbpGeburtsdatum9").Range.Text = txtGeburtsdatum9.Text
			oDoc.Bookmarks.Item("xbpLand9").Range.Text = cmbLand9.Text

			AnzahlTaxPflichig = AnzahlTaxPflichig + 1

		End If


		If txtNachName10.Text <> "" Then
			oDoc.Bookmarks.Item("bpName10").Range.Text = txtNachName10.Text
			oDoc.Bookmarks.Item("bpVorname10").Range.Text = txtVorname10.Text
			oDoc.Bookmarks.Item("bpGeburtsdatum10").Range.Text = txtGeburtsdatum10.Text
			oDoc.Bookmarks.Item("bpLand10").Range.Text = cmbLand10.Text

			oDoc.Bookmarks.Item("xbpName10").Range.Text = txtNachName10.Text
			oDoc.Bookmarks.Item("xbpVorname10").Range.Text = txtVorname10.Text
			oDoc.Bookmarks.Item("xbpGeburtsdatum10").Range.Text = txtGeburtsdatum10.Text
			oDoc.Bookmarks.Item("xbpLand10").Range.Text = cmbLand10.Text

			AnzahlTaxPflichig = AnzahlTaxPflichig + 1

		End If

		oDoc.Bookmarks.Item("tmAnzahlTaxpflichtigePersonenMS").Range.Text = AnzahlTaxPflichig
		AnzahlTaxPflichig = 0


		oDoc.Bookmarks.Item("tmZwischentotal").Range.Text = txtZwischenTotal.Text


		Dim textPart1 As Word.Range = oDoc.Bookmarks.Item("Anzahl").Range
		Dim textPart2 As Word.Range = oDoc.Bookmarks.Item("Bezeichnung").Range
		Dim textPart3 As Word.Range = oDoc.Bookmarks.Item("Einzelpreis").Range
		Dim textPart4 As Word.Range = oDoc.Bookmarks.Item("Totalpreis").Range

		Dim textPartMwst As Word.Range = oDoc.Bookmarks.Item("Mwst").Range
		Dim textPartTotal As Word.Range = oDoc.Bookmarks.Item("Total").Range

		textPart1.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
		textPart2.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
		textPart3.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
		textPart4.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

		textPartMwst.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
		textPartTotal.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight


		Dim CheckString As String = "0"


		Dim AnzahlNaechte As Integer = CInt(txtNaechte.Text)


		'------------------------------------------------------------------------------------------------


		Dim PrintEinzelZimmer1 As Boolean = False
		Dim PrintEinzelZimmer2 As Boolean = False
		Dim PrintEinzelZimmer3 As Boolean = False


		If chkZimmer1Komplett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(32, SprachCode) & " Nr. 1")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Komplett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Komplett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer1 = True

		End If


		If chkZimmer1Einzel.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(30, SprachCode) & " Nr. 1")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Einzel))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Einzel * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer1 = True

		End If


		If PrintEinzelZimmer1 = False Then

			If chkBett11.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.11")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett12.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.12")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett13.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.13")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett14.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.14")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If

		End If



		If chkZimmer1Kinderbett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(34, SprachCode) & " Nr. 1")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Kinderbett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Kinderbett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If


		'------------------------------------------------------------------------------------------------



		If chkZimmer2Komplett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(33, SprachCode) & " Nr. 2")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Komplett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Komplett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer2 = True

		End If


		If chkZimmer2Einzel.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(31, SprachCode) & " Nr. 2")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Einzel))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Einzel * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer2 = True

		End If



		If PrintEinzelZimmer2 = False Then

			If chkBett21.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.21")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett22.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.22")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett23.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.23")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett24.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.24")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett25.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.25")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett26.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.26")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If

		End If


		If chkZimmer2Kinderbett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(34, SprachCode) & " Nr. 2")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Kinderbett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Kinderbett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If



		'-----------------------------------------------------------------------------------------

		If chkZimmer3Komplett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(32, SprachCode) & " Nr. 3")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Komplett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Komplett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer3 = True

		End If


		If chkZimmer3Einzel.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(30, SprachCode) & " Nr. 3")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Einzel))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Einzel * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer3 = True

		End If

		If PrintEinzelZimmer3 = False Then

			If chkBett31.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.31")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett32.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.32")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett33.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.33")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett34.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.34")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If

		End If

		If chkZimmer3Kinderbett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(34, SprachCode) & " Nr. 3")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Kinderbett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Kinderbett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'----------------------------------------------------------------------------------------------

		'------- AUTO -----

		If Me.cmb04.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb04.Text & "/(" & AnzahlNaechte & ")")
			textPart2.InsertAfter(TexteLesen(4))
			textPart3.InsertAfter(Me.txtPreis04.Text)
			textPart4.InsertAfter(Me.total04.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'------- MOTORRAD -----

		If Me.cmb03.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb03.Text & "/(" & AnzahlNaechte & ")")
			textPart2.InsertAfter(TexteLesen(6))
			textPart3.InsertAfter(Me.txtPreis03.Text)
			textPart4.InsertAfter(Me.total03.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If


		'------- HANDTüCHER -----

		If Me.cmb05.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb05.Text)
			textPart2.InsertAfter(TexteLesen(25))
			textPart3.InsertAfter(Me.txtPreis05.Text)
			textPart4.InsertAfter(Me.total05.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'------- FRÜHSTÜCK -----

		If Me.cmb06.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb06.Text)
			textPart2.InsertAfter(TexteLesen(26))
			textPart3.InsertAfter(Me.txtPreis06.Text)
			textPart4.InsertAfter(Me.total06.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'------- KÜCHE -----

		If Me.cmb07.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb07.Text)
			textPart2.InsertAfter(TexteLesen(27))
			textPart3.InsertAfter(Me.txtPreis07.Text)
			textPart4.InsertAfter(Me.total07.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'------- SONSTIGES -----

		If Me.txtSonstigesText.Text <> vbNullString Then

			textPart2.InsertAfter(Me.txtSonstigesText.Text)
			textPart4.InsertAfter(Me.txtSonstigesBetrag.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'------- TAXE -----

		If Me.txtTaxe.Text <> vbNullString Then

			textPart2.InsertAfter("Taxe")
			textPart4.InsertAfter(Me.txtTaxe.Text)

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If


		'------- RABATT -----

		Dim RabattWert As String
		Dim RabattArt As String
		Dim BezeichnungRabatt As String

		If Me.txtRabatt.Text <> vbNullString Then


			If txtRabattProzent.Checked = True Then
				RabattArt = "%"
				Dim G1 As Double = CDbl(Me.Label38.Text)
				RabattWert = "- " & G1.ToString("N2")
				BezeichnungRabatt = "Rabatt" & " " & Me.txtRabatt.Text & " " & RabattArt
			Else

				Dim H1 As Double = CDbl(Me.txtRabatt.Text)
				RabattWert = "- " & H1.ToString("N2")
				BezeichnungRabatt = "Rabatt"

				If txtRabattCHF.Checked = True Then
					RabattArt = ""
				Else
					RabattArt = ""
				End If
			End If

			'textPart2.InsertAfter(BezeichnungRabatt)
			'textPart4.InsertAfter(RabattWert)

			oDoc.Bookmarks.Item("tmRabattFeld").Range.Text = BezeichnungRabatt
			oDoc.Bookmarks.Item("tmRabattWert").Range.Text = RabattWert

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		textPartMwst.InsertAfter(Me.txtMwSt.Text)
		textPartTotal.InsertAfter(Me.txtTotalRechnung.Text)

	End Sub


Public Sub Drucken_Quittung()

		Dim cmd As New OleDbCommand("Select * from tbAdress where AdressId = " & Public_AdressId & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet()

		da.Fill(ds, "tbAdress")

		Dim oWord As Word.Application
		Dim oDoc As Word.Document

		SprachCode = ds.Tables("tbAdress").Rows(0).Item(14)

		oWord = CreateObject("Word.Application")
		oWord.Visible = True

		oDoc = oWord.Documents.Add("c:\Formulare\Quittung.dot")

		oDoc.Bookmarks.Item("tmDatum").Range.Text = System.DateTime.Today
		'oDoc.Bookmarks.Item("tmBuchungsnummer").Range.Text = Me.lblBookingId.Text
		'oDoc.Bookmarks.Item("tmPlatznummer").Range.Text = Me.txtPlatznummer.Text


		Dim strNameVorname As String = ds.Tables("tbAdress").Rows(0).Item(1) & " " & ds.Tables("tbAdress").Rows(0).Item(2)

		oDoc.Bookmarks.Item("tmNameVorname").Range.Text = strNameVorname.ToString
		oDoc.Bookmarks.Item("tmAdresse").Range.Text = ds.Tables("tbAdress").Rows(0).Item(3)
		oDoc.Bookmarks.Item("tmPLZOrt").Range.Text = ds.Tables("tbAdress").Rows(0).Item(4) & " " & ds.Tables("tbAdress").Rows(0).Item(5)


		oDoc.Bookmarks.Item("tmAnreise").Range.Text = DateTimePicker1.Value
		oDoc.Bookmarks.Item("tmAbreise").Range.Text = DateTimePicker2.Value
		oDoc.Bookmarks.Item("tmTotalNaechte").Range.Text = Me.txtNaechte.Text
		oDoc.Bookmarks.Item("tmKennzeichen").Range.Text = Me.txtAutoKennZeichen.Text


		'-----------------------------------------------------------------------------------


		oDoc.Bookmarks.Item("textDokTitel").Range.Text = Form1.TitelErmitteln(7, SprachCode)

		oDoc.Bookmarks.Item("textAnreise").Range.Text = Form1.TitelErmitteln(5, SprachCode)
		oDoc.Bookmarks.Item("textAbreise").Range.Text = Form1.TitelErmitteln(6, SprachCode)

		oDoc.Bookmarks.Item("textTotalNächte").Range.Text = Form1.TitelErmitteln(21, SprachCode)
		oDoc.Bookmarks.Item("textKennzeichen").Range.Text = Form1.TitelErmitteln(3, SprachCode)



		oDoc.Bookmarks.Item("textAnzahl").Range.Text = Form1.TitelErmitteln(16, SprachCode)
		oDoc.Bookmarks.Item("textBezeichnung").Range.Text = Form1.TitelErmitteln(10, SprachCode)
		oDoc.Bookmarks.Item("textEinzelpreis").Range.Text = Form1.TitelErmitteln(11, SprachCode)
		oDoc.Bookmarks.Item("textTotal").Range.Text = Form1.TitelErmitteln(23, SprachCode)
		oDoc.Bookmarks.Item("textZwischenTotal").Range.Text = Form1.TitelErmitteln(25, SprachCode)
		oDoc.Bookmarks.Item("textMwst").Range.Text = Form1.TitelErmitteln(17, SprachCode)
		'oDoc.Bookmarks.Item("textDatum").Range.Text = Form1.TitelErmitteln(26, SprachCode)
		'oDoc.Bookmarks.Item("textBuchungsnummer").Range.Text = Form1.TitelErmitteln(2, SprachCode)
		'oDoc.Bookmarks.Item("textPlatznummer").Range.Text = Form1.TitelErmitteln(27, SprachCode)


		'-----------------------------------------------------------------------------------



		Dim textPart1 As Word.Range = oDoc.Bookmarks.Item("Anzahl").Range
		Dim textPart2 As Word.Range = oDoc.Bookmarks.Item("Bezeichnung").Range
		Dim textPart3 As Word.Range = oDoc.Bookmarks.Item("Einzelpreis").Range
		Dim textPart4 As Word.Range = oDoc.Bookmarks.Item("Totalpreis").Range

		Dim textPartMwst As Word.Range = oDoc.Bookmarks.Item("Mwst").Range
		Dim textPartTotal As Word.Range = oDoc.Bookmarks.Item("Total").Range

		oDoc.Bookmarks.Item("tmZwischentotal").Range.Text = txtZwischenTotal.Text

		textPart1.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter
		textPart2.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphLeft
		textPart3.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
		textPart4.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight

		textPartMwst.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight
		textPartTotal.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight


		Dim CheckString As String = "0"


		Dim AnzahlNaechte As Integer = CInt(txtNaechte.Text)


		'------------------------------------------------------------------------------------------------


		Dim PrintEinzelZimmer1 As Boolean = False
		Dim PrintEinzelZimmer2 As Boolean = False
		Dim PrintEinzelZimmer3 As Boolean = False


		If chkZimmer1Komplett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(32, SprachCode) & " Nr. 1")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Komplett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Komplett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer1 = True

		End If


		If chkZimmer1Einzel.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(30, SprachCode) & " Nr. 1")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Einzel))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Einzel * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer1 = True

		End If

		If PrintEinzelZimmer1 = False Then

			If chkBett11.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.11")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett12.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.12")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett13.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.13")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett14.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.14")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If

		End If



		If chkZimmer1Kinderbett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(34, SprachCode) & " Nr. 1")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Kinderbett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer1_Kinderbett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If


		'------------------------------------------------------------------------------------------------



		If chkZimmer2Komplett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(33, SprachCode) & " Nr. 2")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Komplett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Komplett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer2 = True

		End If


		If chkZimmer2Einzel.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(31, SprachCode) & " Nr. 2")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Einzel))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Einzel * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer2 = True

		End If



		If PrintEinzelZimmer2 = False Then

			If chkBett21.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.21")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett22.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.22")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett23.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.23")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett24.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.24")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett25.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.25")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett26.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(29, SprachCode) & " Nr.26")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_6er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_6er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If

		End If


		If chkZimmer2Kinderbett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(34, SprachCode) & " Nr. 2")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Kinderbett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer2_Kinderbett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If



		'-----------------------------------------------------------------------------------------

		If chkZimmer3Komplett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(32, SprachCode) & " Nr. 3")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Komplett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Komplett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer3 = True

		End If


		If chkZimmer3Einzel.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(30, SprachCode) & " Nr. 3")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Einzel))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Einzel * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

			PrintEinzelZimmer3 = True

		End If

		If PrintEinzelZimmer3 = False Then

			If chkBett31.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.31")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett32.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.32")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett33.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.33")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If


			If chkBett34.Checked = True Then

				textPart1.InsertAfter(Me.txtNaechte.Text)
				textPart2.InsertAfter(Form1.TitelErmitteln(28, SprachCode) & " Nr.34")
				textPart3.InsertAfter(String.Format("{0:N}", Preis_Bett_4er))
				textPart4.InsertAfter(String.Format("{0:N}", Preis_Bett_4er * AnzahlNaechte))

				textPart1.InsertParagraphAfter()
				textPart2.InsertParagraphAfter()
				textPart3.InsertParagraphAfter()
				textPart4.InsertParagraphAfter()

			End If

		End If

		If chkZimmer3Kinderbett.Checked = True Then

			textPart1.InsertAfter(Me.txtNaechte.Text)
			textPart2.InsertAfter(Form1.TitelErmitteln(34, SprachCode) & " Nr. 3")
			textPart3.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Kinderbett))
			textPart4.InsertAfter(String.Format("{0:N}", Preis_Zimmer3_Kinderbett * AnzahlNaechte))

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'----------------------------------------------------------------------------------------------

		'------- AUTO -----

		If Me.cmb04.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb04.Text & "/(" & AnzahlNaechte & ")")
			textPart2.InsertAfter(TexteLesen(4))
			textPart3.InsertAfter(Me.txtPreis04.Text)
			textPart4.InsertAfter(Me.total04.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'------- MOTORRAD -----

		If Me.cmb03.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb03.Text & "/(" & AnzahlNaechte & ")")
			textPart2.InsertAfter(TexteLesen(6))
			textPart3.InsertAfter(Me.txtPreis03.Text)
			textPart4.InsertAfter(Me.total03.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If


		'------- HANDTüCHER -----

		If Me.cmb05.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb05.Text)
			textPart2.InsertAfter(TexteLesen(25))
			textPart3.InsertAfter(Me.txtPreis05.Text)
			textPart4.InsertAfter(Me.total05.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'------- FRÜHSTÜCK -----

		If Me.cmb06.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb06.Text)
			textPart2.InsertAfter(TexteLesen(26))
			textPart3.InsertAfter(Me.txtPreis06.Text)
			textPart4.InsertAfter(Me.total06.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If


		'------- KÜCHE -----

		If Me.cmb07.Text <> CheckString Then

			textPart1.InsertAfter(Me.cmb07.Text)
			textPart2.InsertAfter(TexteLesen(27))
			textPart3.InsertAfter(Me.txtPreis07.Text)
			textPart4.InsertAfter(Me.total07.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If







		'------- SONSTIGES -----

		If Me.txtSonstigesText.Text <> vbNullString Then

			textPart2.InsertAfter(Me.txtSonstigesText.Text)
			textPart4.InsertAfter(Me.txtSonstigesBetrag.Text)


			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		'------- TAXE -----

		If Me.txtTaxe.Text <> vbNullString Then

			textPart2.InsertAfter("Taxe")
			textPart4.InsertAfter(Me.txtTaxe.Text)

			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If


		'------- RABATT -----

		Dim RabattWert As String
		Dim RabattArt As String
		Dim BezeichnungRabatt As String

		If Me.txtRabatt.Text <> vbNullString Then


			If txtRabattProzent.Checked = True Then
				RabattArt = "%"
				Dim G1 As Double = CDbl(Me.Label38.Text)
				RabattWert = "- " & G1.ToString("N2")
				BezeichnungRabatt = "Rabatt" & " " & Me.txtRabatt.Text & " " & RabattArt
			Else

				Dim H1 As Double = CDbl(Me.txtRabatt.Text)
				RabattWert = "- " & H1.ToString("N2")
				BezeichnungRabatt = "Rabatt"

				If txtRabattCHF.Checked = True Then
					RabattArt = ""
				Else
					RabattArt = ""
				End If
			End If

			'textPart2.InsertAfter(BezeichnungRabatt)
			'textPart4.InsertAfter(RabattWert)

			oDoc.Bookmarks.Item("tmRabattFeld").Range.Text = BezeichnungRabatt
			oDoc.Bookmarks.Item("tmRabattWert").Range.Text = RabattWert



			textPart1.InsertParagraphAfter()
			textPart2.InsertParagraphAfter()
			textPart3.InsertParagraphAfter()
			textPart4.InsertParagraphAfter()

		End If

		textPartMwst.InsertAfter(Me.txtMwSt.Text)
		textPartTotal.InsertAfter(Me.txtTotalRechnung.Text)

	End Sub

	Function TexteLesen(ByVal intPreisId As Integer)

		Dim cmd As New OleDbCommand("Select * from tbPreis where PreisId = " & intPreisId & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet()

		da.Fill(ds, "tbPreis")

		Select Case SprachCode
			Case "D"
				Return ds.Tables("tbPreis").Rows(0).Item(1)
			Case "E"
				Return ds.Tables("tbPreis").Rows(0).Item(2)
			Case "F"
				Return ds.Tables("tbPreis").Rows(0).Item(3)
			Case "I"
				Return ds.Tables("tbPreis").Rows(0).Item(4)
		End Select


	End Function

		Public Function AnzahlAusdruckeBuchungsbestaetigung()

		Dim datName As String = "PrintBuchung.ini"
		Dim reader As StreamReader = File.OpenText(datName)

		While (reader.Peek() > -1)

			Return reader.ReadLine()
			Exit While

		End While

		reader.Close()

	End Function


Public Function AnzahlAusdruckeQuittung()

		Dim datName As String = "PrintQuittung.ini"
		Dim reader As StreamReader = File.OpenText(datName)

		While (reader.Peek() > -1)

			Return reader.ReadLine()
			Exit While

		End While

		reader.Close()

	End Function

		Private Sub AktivFlagSetzen()

		Dim cmd As New OleDbCommand("UPDATE tbZimmer SET Aktiv  = " & False & " WHERE BookingId = " & Public_BookingId & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet("tbZimmer")

		Try
			da.Fill(ds, "tbZimmer")
		Catch ex As Exception
			MessageBox.Show(ex.Message)
		End Try

	End Sub

End Class

			'Dim CountRun As Integer
			'CountRun = 10

			'For ii = 3 To 25

			'	CountRun = CountRun + 1

				'Debug.WriteLine("--------------------------")
				'Debug.WriteLine(ii)
				'Debug.WriteLine(ds.Tables("tbZimmerGebucht").Rows(0).Item(ii))
				'Debug.WriteLine("--------------------------")

					'Dim matches() As Control
					'matches = Me.Controls.Find("chkBett" & CountRun, True)

					'Debug.WriteLine(CountRun)
					'If matches.Length > 0 AndAlso TypeOf matches(0) Is CheckBox Then
				'		Dim cb As CheckBox = DirectCast(matches(0), CheckBox)
					'	Debug.WriteLine(cb.CheckState)
					'End If
					'Debug.WriteLine("--------------------------")


					'If matches.Length > 0 AndAlso TypeOf matches(0) Is CheckBox Then
					'	Dim cb As CheckBox = DirectCast(matches(0), CheckBox)
					'	If cb.Checked Then
						'Debug.WriteLine("--------------------------------------------------------------------------")

						'Debug.WriteLine(ii)
						'Debug.WriteLine(ds.Tables("tbZimmerGebucht").Rows(0).Item(ii))
						'Debug.WriteLine(cb.Name)
						'Debug.WriteLine(cb.CheckState)
						'Debug.WriteLine("--------------------------------------------------------------------------")

						'If cb.Checked = True And ds.Tables("tbZimmerGebucht").Rows(0).Item(ii) = True Then

						'Debug.WriteLine("Zimmer gebucht")

						'End If

						'Debug.WriteLine("--------------------------------------------------------------------------")

					'End If
				'End If


			'Next