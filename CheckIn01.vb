
Imports System.Data.OleDb
Imports System.Console
Imports System.Drawing.Printing
Imports System.IO
Imports System.Math
Imports Microsoft.Office.Interop

Public Class CheckIn01

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


    Dim conn As New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=camping.mdb")
    

    Private Sub CheckIn01_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

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

        PreisLaden()

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

    Private Sub PreisLaden()

        Dim cmd As New OleDbCommand("Select * from tbPreis", conn)
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet()

        da.Fill(ds, "tbPreis")

        txtPreis01.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(0).Item(5))
        txtPreis02.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(1).Item(5))
        txtPreis03.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(2).Item(5))
        txtPreis04.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(3).Item(5))
        txtPreis05.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(4).Item(5))
        txtPreis06.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(5).Item(5))
        txtPreis07.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(6).Item(5))
        txtPreis08.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(7).Item(5))
        txtPreis09.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(8).Item(5))
        txtPreis10.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(9).Item(5))
        txtPreis11.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(10).Item(5))
        txtPreis12.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(11).Item(5))
        txtPreis13.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(12).Item(5))
        txtPreis14.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(13).Item(5))
        txtPreis15.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(14).Item(5))
        txtPreis16.Text = String.Format("{0:N}", ds.Tables("tbPreis").Rows(15).Item(5))

		Berechnen()
        TotalBerechnen()


    End Sub


    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        Me.Close()

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

    Public Sub StartMutation(ByVal i As Integer)

        StartFlag = "StartMutation"

        Dim BookingId As Integer
        BookingId = i
        Public_BookingId = i

        lblBookingId.Text = BookingId.ToString

        Dim da As New OleDbDataAdapter("SELECT * FROM tbBooking " _
                & "INNER JOIN (tbAdress INNER JOIN tbBookingtbAdress ON tbAdress.AdressId = tbBookingtbAdress.AdressId) " _
                & "ON tbBooking.BookingId = tbBookingtbAdress.BookingId where tbBookingtbAdress.Leader = true and tbBooking.BookingId = " & BookingId & "", conn)

        Dim ds As New DataSet()

        da.Fill(ds, "tbBooking")

        Public_AdressId = ds.Tables("tbBooking").Rows(0).Item(28).ToString()

        lblNameKunde.Text = (ds.Tables("tbBooking").Rows(0).Item(29).ToString()) & " " & (ds.Tables("tbBooking").Rows(0).Item(30).ToString())

        DateTimePicker1.Value = ds.Tables("tbBooking").Rows(0).Item(1).ToString()

        DateTimePicker2.Value = ds.Tables("tbBooking").Rows(0).Item(2).ToString()

        txtNaechte.Text = ds.Tables("tbBooking").Rows(0).Item(3).ToString()

        txtAutoKennZeichen.Text = ds.Tables("tbBooking").Rows(0).Item(20).ToString()

        txtPlatznummer.Text = ds.Tables("tbBooking").Rows(0).Item(21).ToString()

        txtBemerkungen.Text = ds.Tables("tbBooking").Rows(0).Item(22).ToString()

        txtRabatt.Text = ds.Tables("tbBooking").Rows(0).Item(23).ToString()


        If AdressenErfassen.CheckDBNull(ds.Tables("tbBooking").Rows(0).Item(24)) <> "" Then
            If ds.Tables("tbBooking").Rows(0).Item(24) = "%" Then
                txtRabattProzent.Checked = True
            Else
                txtRabattCHF.Checked = True
            End If
        End If


        txtSonstigesText.Text = ds.Tables("tbBooking").Rows(0).Item(25).ToString()
        txtSonstigesBetrag.Text = ds.Tables("tbBooking").Rows(0).Item(26).ToString()

        cmb01.Text = ds.Tables("tbBooking").Rows(0).Item(5)
        cmb02.Text = ds.Tables("tbBooking").Rows(0).Item(6)
        cmb03.Text = ds.Tables("tbBooking").Rows(0).Item(7)
        cmb04.Text = ds.Tables("tbBooking").Rows(0).Item(8)
        cmb05.Text = ds.Tables("tbBooking").Rows(0).Item(9)
        cmb06.Text = ds.Tables("tbBooking").Rows(0).Item(10)
        cmb07.Text = ds.Tables("tbBooking").Rows(0).Item(11)
        cmb08.Text = ds.Tables("tbBooking").Rows(0).Item(12)
        cmb09.Text = ds.Tables("tbBooking").Rows(0).Item(13)
        cmb10.Text = ds.Tables("tbBooking").Rows(0).Item(14)
        cmb11.Text = ds.Tables("tbBooking").Rows(0).Item(15)
        cmb12.Text = ds.Tables("tbBooking").Rows(0).Item(16)
        cmb13.Text = ds.Tables("tbBooking").Rows(0).Item(17)
        cmb14.Text = ds.Tables("tbBooking").Rows(0).Item(18)
        cmb15.Text = ds.Tables("tbBooking").Rows(0).Item(19)
        cmb16.Text = ds.Tables("tbBooking").Rows(0).Item(27)


        '-----------------------------------------------------------------------

        Dim da1 As New OleDbDataAdapter("SELECT tbAdress.AdressId, tbAdress.NachName, tbAdress.VorName,tbAdress.Geburtsdatum, tbAdress.Land FROM tbBookingtbAdress " _
                    & "INNER JOIN tbAdress ON tbBookingtbAdress.AdressId = tbAdress.AdressId " _
                    & "where tbBookingtbAdress.BookingId = " & BookingId & "", conn)

        Dim ds1 As New DataSet()

        da1.Fill(ds1, "tbBooking")

        Dim rc As Integer = ds1.Tables("tbBooking").Rows.Count()
        rc = rc - 1

        If rc >= 1 Then
            txtNachName1.Text = ds1.Tables("tbBooking").Rows(1).Item(1).ToString()
            txtVorname1.Text = ds1.Tables("tbBooking").Rows(1).Item(2).ToString()
            txtGeburtsdatum1.Text = ds1.Tables("tbBooking").Rows(1).Item(3)
            LandErfasst1 = ds1.Tables("tbBooking").Rows(1).Item(4).ToString()
            lblId1.Text = ds1.Tables("tbBooking").Rows(1).Item(0).ToString()
        End If

        If rc >= 2 Then
            txtNachName2.Text = ds1.Tables("tbBooking").Rows(2).Item(1).ToString()
            txtVorname2.Text = ds1.Tables("tbBooking").Rows(2).Item(2).ToString()
            txtGeburtsdatum2.Text = ds1.Tables("tbBooking").Rows(2).Item(3).ToString()
            LandErfasst2 = ds1.Tables("tbBooking").Rows(2).Item(4).ToString()
            lblId2.Text = ds1.Tables("tbBooking").Rows(2).Item(0).ToString()
        End If

        If rc >= 3 Then
            txtNachName3.Text = ds1.Tables("tbBooking").Rows(3).Item(1).ToString()
            txtVorname3.Text = ds1.Tables("tbBooking").Rows(3).Item(2).ToString()
            txtGeburtsdatum3.Text = ds1.Tables("tbBooking").Rows(3).Item(3).ToString()
            LandErfasst3 = ds1.Tables("tbBooking").Rows(3).Item(4).ToString()
            lblId3.Text = ds1.Tables("tbBooking").Rows(3).Item(0).ToString()
        End If

        If rc >= 4 Then
            txtNachName4.Text = ds1.Tables("tbBooking").Rows(4).Item(1).ToString()
            txtVorname4.Text = ds1.Tables("tbBooking").Rows(4).Item(2).ToString()
            txtGeburtsdatum4.Text = ds1.Tables("tbBooking").Rows(4).Item(3).ToString()
            LandErfasst4 = ds1.Tables("tbBooking").Rows(4).Item(4).ToString()
            lblId4.Text = ds1.Tables("tbBooking").Rows(4).Item(0).ToString()
        End If

        If rc >= 5 Then
            txtNachName5.Text = ds1.Tables("tbBooking").Rows(5).Item(1).ToString()
            txtVorname5.Text = ds1.Tables("tbBooking").Rows(5).Item(2).ToString()
            txtGeburtsdatum5.Text = ds1.Tables("tbBooking").Rows(5).Item(3).ToString()
            LandErfasst5 = ds1.Tables("tbBooking").Rows(5).Item(4).ToString()
            lblId5.Text = ds1.Tables("tbBooking").Rows(5).Item(0).ToString()
		End If

		If rc >= 6 Then
			txtNachName6.Text = ds1.Tables("tbBooking").Rows(6).Item(1).ToString()
			txtVorname6.Text = ds1.Tables("tbBooking").Rows(6).Item(2).ToString()
			txtGeburtsdatum6.Text = ds1.Tables("tbBooking").Rows(6).Item(3).ToString()
			LandErfasst6 = ds1.Tables("tbBooking").Rows(6).Item(4).ToString()
			lblId6.Text = ds1.Tables("tbBooking").Rows(6).Item(0).ToString()
		End If

		If rc >= 7 Then
			txtNachName7.Text = ds1.Tables("tbBooking").Rows(7).Item(1).ToString()
			txtVorname7.Text = ds1.Tables("tbBooking").Rows(7).Item(2).ToString()
			txtGeburtsdatum7.Text = ds1.Tables("tbBooking").Rows(7).Item(3).ToString()
			LandErfasst7 = ds1.Tables("tbBooking").Rows(7).Item(4).ToString()
			lblId7.Text = ds1.Tables("tbBooking").Rows(7).Item(0).ToString()
		End If

		If rc >= 8 Then
			txtNachName8.Text = ds1.Tables("tbBooking").Rows(8).Item(1).ToString()
			txtVorname8.Text = ds1.Tables("tbBooking").Rows(8).Item(2).ToString()
			txtGeburtsdatum8.Text = ds1.Tables("tbBooking").Rows(8).Item(3).ToString()
			LandErfasst8 = ds1.Tables("tbBooking").Rows(8).Item(4).ToString()
			lblId8.Text = ds1.Tables("tbBooking").Rows(8).Item(0).ToString()
		End If

		If rc >= 9 Then
			txtNachName9.Text = ds1.Tables("tbBooking").Rows(9).Item(1).ToString()
			txtVorname9.Text = ds1.Tables("tbBooking").Rows(9).Item(2).ToString()
			txtGeburtsdatum9.Text = ds1.Tables("tbBooking").Rows(9).Item(3).ToString()
			LandErfasst9 = ds1.Tables("tbBooking").Rows(9).Item(4).ToString()
			lblId9.Text = ds1.Tables("tbBooking").Rows(9).Item(0).ToString()
		End If

		If rc >= 10 Then
			txtNachName10.Text = ds1.Tables("tbBooking").Rows(10).Item(1).ToString()
			txtVorname10.Text = ds1.Tables("tbBooking").Rows(10).Item(2).ToString()
			txtGeburtsdatum10.Text = ds1.Tables("tbBooking").Rows(10).Item(3).ToString()
			LandErfasst10 = ds1.Tables("tbBooking").Rows(10).Item(4).ToString()
			lblId10.Text = ds1.Tables("tbBooking").Rows(10).Item(0).ToString()
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


        If ds.Tables("tbBooking").Rows(0).Item(4).ToString() = False Then

            btnSpeichern.Enabled = False
            btnSpeichernUndSchliessen.Enabled = False
            btnBestaetigungMitMeldeschein.Enabled = False
            btnBestaetigungOhneMeldeschein.Enabled = False
            btnQuittung.Enabled = False
        End If



    End Sub


    Private Sub btnSpeichern_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSpeichern.Click

        If StartFlag = "StartBooking" Then

            SpeichernBookingDaten()
            StartSpeichernBegleitpersonen()
            MsgBox("Buchungsdaten gespeichert", MsgBoxStyle.Exclamation)
            
            StartMutation(Public_BookingId)

        Else

            MutierenBookingDaten()
            StartMutierenBegleitpersonen()
            MsgBox("Buchungsdaten mutiert", MsgBoxStyle.Exclamation)

        End If


    End Sub


    Private Sub btnSpeichernUndSchliessen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSpeichernUndSchliessen.Click

        If StartFlag = "StartBooking" Then

            SpeichernBookingDaten()
            StartSpeichernBegleitpersonen()
            MsgBox("Buchungsdaten gespeichert", MsgBoxStyle.Exclamation)

        Else

            MutierenBookingDaten()
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

    Private Sub MutierenBookingDaten()

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

        Dim cmd As New OleDbCommand("UPDATE tbBooking SET Anreise = '" & DateTimePicker1.Value.Date & "'" _
            & ", Abreise = '" & DateTimePicker2.Value.Date & "'" _
            & ", AnzahlNaechte = '" & txtNaechte.Text & "'" _
            & ", Aktiv = " & True & "" _
            & ", 1 = '" & cmb01.Text & "'" _
            & ", 2 = '" & cmb02.Text & "'" _
            & ", 3 = '" & cmb03.Text & "'" _
            & ", 4 = '" & cmb04.Text & "'" _
            & ", 5 = '" & cmb05.Text & "'" _
            & ", 6 = '" & cmb06.Text & "'" _
            & ", 7 = '" & cmb07.Text & "'" _
            & ", 8 = '" & cmb08.Text & "'" _
            & ", 9 = '" & cmb09.Text & "'" _
            & ", 10 = '" & cmb10.Text & "'" _
            & ", 11 = '" & cmb11.Text & "'" _
            & ", 12 = '" & cmb12.Text & "'" _
            & ", 13 = '" & cmb13.Text & "'" _
            & ", 14 = '" & cmb14.Text & "'" _
            & ", 15 = '" & cmb15.Text & "'" _
            & ", 16 = '" & cmb16.Text & "'" _
            & ", Autokennzeichen = '" & txtAutoKennZeichen.Text & "'" _
            & ", PlatzNummer = '" & txtPlatznummer.Text & "'" _
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

        '---------------------------------



    End Sub


    Private Sub SpeichernBookingDaten()

        Dim strRabattArt As String

        If txtRabattProzent.Checked = True Then
            strRabattArt = "%"
        Else
            strRabattArt = ""
        End If

        If txtRabattCHF.Checked = True Then
            strRabattArt = "CHF"
        Else
            strRabattArt = ""
        End If


        InitCMB()

        Dim cmd As New OleDbCommand("Insert Into tbBooking (Anreise, Abreise, AnzahlNaechte,Aktiv" _
                & ",1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,AutoKennZeichen,Platznummer,Bemerkungen,RabattWert,RabattArt,SonstigesText,SonstigesBetrag)" _
                & "Values('" & DateTimePicker1.Value.Date & "'" _
                & ", '" & DateTimePicker2.Value.Date & "'" _
                & ", '" & txtNaechte.Text & "'" _
                & ", " & True & "" _
                & ", '" & cmb01.Text & "'" _
                & ", '" & cmb02.Text & "'" _
                & ", '" & cmb03.Text & "'" _
                & ", '" & cmb04.Text & "'" _
                & ", '" & cmb05.Text & "'" _
                & ", '" & cmb06.Text & "'" _
                & ", '" & cmb07.Text & "'" _
                & ", '" & cmb08.Text & "'" _
                & ", '" & cmb09.Text & "'" _
                & ", '" & cmb10.Text & "'" _
                & ", '" & cmb11.Text & "'" _
                & ", '" & cmb12.Text & "'" _
                & ", '" & cmb13.Text & "'" _
                & ", '" & cmb14.Text & "'" _
                & ", '" & cmb15.Text & "'" _
                & ", '" & cmb16.Text & "'" _
                & ", '" & txtAutoKennZeichen.Text & "'" _
                & ", '" & txtPlatznummer.Text & "','" & txtBemerkungen.Text & "','" & txtRabatt.Text & "','" & strRabattArt.ToString & "','" & txtSonstigesText.Text & "','" & txtSonstigesBetrag.Text & "')", conn)

        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet("tbBooking")

        Try
            da.Fill(ds, "tbBooking")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


        '---------------------------------

        Dim cmd1 As New OleDbCommand("Select * from tbBooking Order by BookingId ASC", conn)
        Dim da1 As New OleDbDataAdapter(cmd1)
        Dim ds1 As New DataSet()

        Dim d, e As Integer

        da1.Fill(ds1, "tbBooking1")

        d = ds1.Tables("tbBooking1").Rows.Count
        e = ds1.Tables("tbBooking1").Rows(d - 1).Item(0)

        Public_BookingId = e

        '---------------------------------

        Dim cmd2 As New OleDbCommand("Insert Into tbBookingtbAdress (BookingId, AdressId, Leader, Moddate) Values(" & e & ", " & Public_AdressId & ",true,Now())", conn)
        Dim da2 As New OleDbDataAdapter(cmd2)
        Dim ds2 As New DataSet()

        Try
            da2.Fill(ds2, "tbBookingtbAdress")
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
        If cmb08.Text = "" Then cmb08.Text = 0
        If cmb09.Text = "" Then cmb09.Text = 0
        If cmb10.Text = "" Then cmb10.Text = 0
        If cmb11.Text = "" Then cmb11.Text = 0
        If cmb12.Text = "" Then cmb12.Text = 0
        If cmb13.Text = "" Then cmb13.Text = 0
        If cmb14.Text = "" Then cmb14.Text = 0
        If cmb15.Text = "" Then cmb15.Text = 0
        If cmb16.Text = "" Then cmb16.Text = 0

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

        tbBookingtbAdressSchreiben()

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

    Private Sub tbBookingtbAdressSchreiben()

        Dim cmd As New OleDbCommand("Select * from tbAdress Order by AdressId ASC", conn)
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet()

        Dim k, i As Integer

        da.Fill(ds, "tbAdress")

        k = ds.Tables("tbAdress").Rows.Count
        i = ds.Tables("tbAdress").Rows(k - 1).Item(0)

        '----------------------------------------------------------------

        Dim cmd1 As New OleDbCommand("Select * from tbBooking Order by BookingId ASC", conn)
        Dim da1 As New OleDbDataAdapter(cmd1)
        Dim ds1 As New DataSet()

        Dim v, w As Integer

        da1.Fill(ds1, "tbBooking")

        v = ds1.Tables("tbBooking").Rows.Count
        w = ds1.Tables("tbBooking").Rows(v - 1).Item(0)

        'Wenn die Buchung besteht, wir die Buchungsnummer für die Mutation der Begl. Personen verwendet
        'Nur bei Neuerfassung wird der letzte Record gelesen

        If StartFlag = "StartMutation" Then
            w = Public_BookingId
        End If

        '----------------------------------------------------------------

        Dim cmd2 As New OleDbCommand("Insert Into tbBookingtbAdress (BookingId, AdressId, Moddate) Values(" & w & ", " & i & ",Now())", conn)
        Dim da2 As New OleDbDataAdapter(cmd2)
        Dim ds2 As New DataSet()

        Try
            da2.Fill(ds2, "tbBookingtbAdress")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


    End Sub


    Private Sub Berechnen()

        Dim s1, s2, s3, s4, s5, s6, s7, s8, s9, s10, s11, s12, s13, s14, s15, s16 As Double
        Dim w1, w2, w3, w4, w5, w6, w7, w8, w9, w10, w11, w12, w13, w14, w15, w16 As Double
        Dim v1, v2, v3, v4, v5, v6, v7, v8, v9, v10, v11, v12, v13, v14, v15, v16 As Integer

        Dim AnzahlNaechte As Integer = CInt(txtNaechte.Text)


        If txtPreis01.Text <> "" Then w1 = txtPreis01.Text
        If txtPreis02.Text <> "" Then w2 = txtPreis02.Text
        If txtPreis03.Text <> "" Then w3 = txtPreis03.Text
        If txtPreis01.Text <> "" Then w1 = txtPreis01.Text
        If txtPreis02.Text <> "" Then w2 = txtPreis02.Text
        If txtPreis03.Text <> "" Then w3 = txtPreis03.Text
        If txtPreis04.Text <> "" Then w4 = txtPreis04.Text
        If txtPreis05.Text <> "" Then w5 = txtPreis05.Text
        If txtPreis06.Text <> "" Then w6 = txtPreis06.Text
        If txtPreis07.Text <> "" Then w7 = txtPreis07.Text
        If txtPreis08.Text <> "" Then w8 = txtPreis08.Text
        If txtPreis09.Text <> "" Then w9 = txtPreis09.Text
        If txtPreis10.Text <> "" Then w10 = txtPreis10.Text
        If txtPreis11.Text <> "" Then w11 = txtPreis11.Text
        If txtPreis12.Text <> "" Then w12 = txtPreis12.Text
        If txtPreis13.Text <> "" Then w13 = txtPreis13.Text
        If txtPreis14.Text <> "" Then w14 = txtPreis14.Text
        If txtPreis15.Text <> "" Then w15 = txtPreis15.Text
        If txtPreis16.Text <> "" Then w16 = txtPreis16.Text


        If cmb01.Text <> "" Then v1 = cmb01.Text
        If cmb02.Text <> "" Then v2 = cmb02.Text
        If cmb03.Text <> "" Then v3 = cmb03.Text
        If cmb04.Text <> "" Then v4 = cmb04.Text
        If cmb05.Text <> "" Then v5 = cmb05.Text
        If cmb06.Text <> "" Then v6 = cmb06.Text
        If cmb07.Text <> "" Then v7 = cmb07.Text
        If cmb08.Text <> "" Then v8 = cmb08.Text
        If cmb09.Text <> "" Then v9 = cmb09.Text
        If cmb10.Text <> "" Then v10 = cmb10.Text
        If cmb11.Text <> "" Then v11 = cmb11.Text
        If cmb12.Text <> "" Then v12 = cmb12.Text
        If cmb13.Text <> "" Then v13 = cmb13.Text
        If cmb14.Text <> "" Then v14 = cmb14.Text
        If cmb15.Text <> "" Then v15 = cmb15.Text
        If cmb16.Text <> "" Then v16 = cmb16.Text


        If BungalowCheck = True Then
            If v15 = 0 Then
                s1 = w1 * v1 * AnzahlNaechte
                s2 = w2 * v2 * AnzahlNaechte
            Else
                s1 = 0
                s2 = 0
            End If

        Else
            If v15 = 0 Then
                s1 = w1 * v1 * AnzahlNaechte
                s2 = w2 * v2 * AnzahlNaechte
            End If
        End If



        s3 = w3 * v3 * AnzahlNaechte
        s4 = w4 * v4 * AnzahlNaechte
        s5 = w5 * v5 * AnzahlNaechte
        s6 = w6 * v6 * AnzahlNaechte
        s7 = w7 * v7 * AnzahlNaechte
        s8 = w8 * v8 * AnzahlNaechte
        s9 = w9 * v9 * AnzahlNaechte
        s10 = w10 * v10 * AnzahlNaechte
        s11 = w11 * v11 * AnzahlNaechte
        s12 = w12 * v12 * AnzahlNaechte
        s13 = w13 * v13 * AnzahlNaechte
        s14 = w14 * v14 * AnzahlNaechte
        s15 = w15 * v15 * AnzahlNaechte
        s16 = w16 * v16 * AnzahlNaechte


        total01.Text = String.Format("{0:N}", s1)
        total02.Text = String.Format("{0:N}", s2)
        total03.Text = String.Format("{0:N}", s3)
        total04.Text = String.Format("{0:N}", s4)
        total05.Text = String.Format("{0:N}", s5)
        total06.Text = String.Format("{0:N}", s6)
        total07.Text = String.Format("{0:N}", s7)
        total08.Text = String.Format("{0:N}", s8)
        total09.Text = String.Format("{0:N}", s9)
        total10.Text = String.Format("{0:N}", s10)
        total11.Text = String.Format("{0:N}", s11)
        total12.Text = String.Format("{0:N}", s12)
        total13.Text = String.Format("{0:N}", s13)
        total14.Text = String.Format("{0:N}", s14)
        total15.Text = String.Format("{0:N}", s15)
        total16.Text = String.Format("{0:N}", s16)

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

    Private Sub cmb08_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb08.SelectedIndexChanged
        Berechnen()
        TotalBerechnen()
    End Sub

    Private Sub cmb09_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb09.SelectedIndexChanged
        Berechnen()
        TotalBerechnen()
    End Sub

    Private Sub cmb10_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb10.SelectedIndexChanged
        Berechnen()
        TotalBerechnen()
    End Sub
    Private Sub cmb11_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb11.SelectedIndexChanged
        Berechnen()
        TotalBerechnen()
    End Sub

    Private Sub cmb12_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb12.SelectedIndexChanged
        Berechnen()
        TotalBerechnen()
    End Sub

    Private Sub cmb13_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb13.SelectedIndexChanged
        Berechnen()
        TotalBerechnen()
    End Sub

    Private Sub cmb14_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb14.SelectedIndexChanged
        Berechnen()
        TotalBerechnen()
    End Sub

    Private Sub cmb15_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb15.SelectedIndexChanged

        BungalowCheck = True
        Berechnen()
        TotalBerechnen()
        BungalowCheck = False

    End Sub

    Private Sub cmb16_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles cmb16.SelectedIndexChanged
        Berechnen()
        TotalBerechnen()
    End Sub

	Private Sub TotalBerechnen()

		Dim total As Double
		Dim t1, t2, t3, t4, t5, t6, t7, t8, t9, t10, t11, t12, t13, t14, t15, t16 As Double


		If total01.Text = "" Then total01.Text = 0
		If total02.Text = "" Then total02.Text = 0
		If total03.Text = "" Then total03.Text = 0
		If total04.Text = "" Then total04.Text = 0
		If total05.Text = "" Then total05.Text = 0
		If total06.Text = "" Then total06.Text = 0
		If total07.Text = "" Then total07.Text = 0
		If total08.Text = "" Then total08.Text = 0
		If total09.Text = "" Then total09.Text = 0
		If total10.Text = "" Then total10.Text = 0
		If total11.Text = "" Then total11.Text = 0
		If total12.Text = "" Then total12.Text = 0
		If total12.Text = "" Then total12.Text = 0
		If total13.Text = "" Then total13.Text = 0
		If total14.Text = "" Then total14.Text = 0
		If total15.Text = "" Then total15.Text = 0
		If total16.Text = "" Then total16.Text = 0


		If total01.Text <> "" Then t1 = total01.Text
		If total02.Text <> "" Then t2 = total02.Text
		If total03.Text <> "" Then t3 = total03.Text
		If total04.Text <> "" Then t4 = total04.Text
		If total05.Text <> "" Then t5 = total05.Text
		If total06.Text <> "" Then t6 = total06.Text
		If total07.Text <> "" Then t7 = total07.Text
		If total08.Text <> "" Then t8 = total08.Text
		If total09.Text <> "" Then t9 = total09.Text
		If total10.Text <> "" Then t10 = total10.Text
		If total11.Text <> "" Then t11 = total11.Text
		If total12.Text <> "" Then t12 = total12.Text
		If total13.Text <> "" Then t13 = total13.Text
		If total14.Text <> "" Then t14 = total14.Text
		If total15.Text <> "" Then t15 = total15.Text
		If total16.Text <> "" Then t16 = total16.Text


		total = t1 + t2 + t3 + t4 + t5 + t6 + t7 + t8 + t9 + t10 + t11 + t12 + t13 + t14 + t15 + t16

		'Zwischentotal
		txtZwischenTotal.Text = String.Format("{0:N}", total)
		Dim ZwischenTotal As Double = txtZwischenTotal.Text

		'TotalNächte
		Dim TotalNaechte As Integer = CInt(txtNaechte.Text)

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


    Private Sub AktivFlagSetzen()

        Dim cmd As New OleDbCommand("UPDATE tbBooking SET Aktiv  = " & False & " WHERE BookingId = " & Public_BookingId & "", conn)
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet("tbBooking")

        Try
            da.Fill(ds, "tbBooking")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub

    Public Function MwstLesen()

        Dim datName As String = "mwst.ini"
        Dim reader As StreamReader = File.OpenText(datName)

        While (reader.Peek() > -1)

            Return reader.ReadLine()
            Exit While

        End While

        reader.Close()

    End Function

    Public Function TaxeLesen()

        Dim datName As String = "taxe.ini"
        Dim reader As StreamReader = File.OpenText(datName)

        While (reader.Peek() > -1)

            Return reader.ReadLine()
            Exit While

        End While

        reader.Close()

    End Function

    Public Function DruckerLesen()

        Dim datName As String = "drucker.ini"
        Dim reader As StreamReader = File.OpenText(datName)

        While (reader.Peek() > -1)

            Return reader.ReadLine()
            Exit While

        End While

        reader.Close()

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
        oDoc.Bookmarks.Item("tmPlatznummer").Range.Text = Me.txtPlatznummer.Text


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


        '------- ERWACHSENE -----

        If Me.cmb01.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb01.Text)
            textPart2.InsertAfter(TexteLesen(1))
            textPart3.InsertAfter(Me.txtPreis01.Text)
            textPart4.InsertAfter(Me.total01.Text)

            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- KINDER -----

        If Me.cmb02.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb02.Text)
            textPart2.InsertAfter(TexteLesen(2))
            textPart3.InsertAfter(Me.txtPreis02.Text)
            textPart4.InsertAfter(Me.total02.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- HUND -----

        If Me.cmb11.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb11.Text)
            textPart2.InsertAfter(TexteLesen(11))
            textPart3.InsertAfter(Me.txtPreis11.Text)
            textPart4.InsertAfter(Me.total11.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- AUTO -----

        If Me.cmb04.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb04.Text)
            textPart2.InsertAfter(TexteLesen(4))
            textPart3.InsertAfter(Me.txtPreis04.Text)
            textPart4.InsertAfter(Me.total04.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- MOTORRAD -----

        If Me.cmb06.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb06.Text)
            textPart2.InsertAfter(TexteLesen(6))
            textPart3.InsertAfter(Me.txtPreis06.Text)
            textPart4.InsertAfter(Me.total06.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- BEHERBERUNG ERWACHSENE -----

        If Me.cmb03.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb03.Text)
            textPart2.InsertAfter(TexteLesen(3))
            textPart3.InsertAfter(Me.txtPreis03.Text)
            textPart4.InsertAfter(Me.total03.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- BEHERBERUNG KINDER -----

        If Me.cmb16.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb16.Text)
            textPart2.InsertAfter(TexteLesen(16))
            textPart3.InsertAfter(Me.txtPreis16.Text)
            textPart4.InsertAfter(Me.total16.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- BUNGALOW -----

        If Me.cmb15.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb15.Text)
            textPart2.InsertAfter(TexteLesen(15))
            textPart3.InsertAfter(Me.txtPreis15.Text)
            textPart4.InsertAfter(Me.total15.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- UEBERNACHTUNG AUTO -----

        If Me.cmb05.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb05.Text)
            textPart2.InsertAfter(TexteLesen(5))
            textPart3.InsertAfter(Me.txtPreis05.Text)
            textPart4.InsertAfter(Me.total05.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- ZELT KLEIN -----

        If Me.cmb07.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb07.Text)
            textPart2.InsertAfter(TexteLesen(7))
            textPart3.InsertAfter(Me.txtPreis07.Text)
            textPart4.InsertAfter(Me.total07.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- ZELT GROSS -----

        If Me.cmb08.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb08.Text)
            textPart2.InsertAfter(TexteLesen(8))
            textPart3.InsertAfter(Me.txtPreis08.Text)
            textPart4.InsertAfter(Me.total08.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If




        '------- WOHNWAGEN -----

        If Me.cmb09.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb09.Text)
            textPart2.InsertAfter(TexteLesen(9))
            textPart3.InsertAfter(Me.txtPreis09.Text)
            textPart4.InsertAfter(Me.total09.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- WOHNMOBIL -----

        If Me.cmb10.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb10.Text)
            textPart2.InsertAfter(TexteLesen(10))
            textPart3.InsertAfter(Me.txtPreis10.Text)
            textPart4.InsertAfter(Me.total10.Text)

            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- STROM -----

        If Me.cmb12.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb12.Text)
            textPart2.InsertAfter(TexteLesen(12))
            textPart3.InsertAfter(Me.txtPreis12.Text)
            textPart4.InsertAfter(Me.total12.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- STROM FUNKER -----

        If Me.cmb13.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb13.Text)
            textPart2.InsertAfter(TexteLesen(13))
            textPart3.InsertAfter(Me.txtPreis13.Text)
            textPart4.InsertAfter(Me.total13.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- GEBUEHREN -----

        If Me.cmb14.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb14.Text)
            textPart2.InsertAfter(TexteLesen(14))
            textPart3.InsertAfter(Me.txtPreis14.Text)
            textPart4.InsertAfter(Me.total14.Text)


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
        oDoc.Bookmarks.Item("tmPlatznummer").Range.Text = Me.txtPlatznummer.Text


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


        '------- ERWACHSENE -----

        If Me.cmb01.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb01.Text)
            textPart2.InsertAfter(TexteLesen(1))
            textPart3.InsertAfter(Me.txtPreis01.Text)
            textPart4.InsertAfter(Me.total01.Text)

            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- KINDER -----

        If Me.cmb02.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb02.Text)
            textPart2.InsertAfter(TexteLesen(2))
            textPart3.InsertAfter(Me.txtPreis02.Text)
            textPart4.InsertAfter(Me.total02.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- HUND -----

        If Me.cmb11.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb11.Text)
            textPart2.InsertAfter(TexteLesen(11))
            textPart3.InsertAfter(Me.txtPreis11.Text)
            textPart4.InsertAfter(Me.total11.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- AUTO -----

        If Me.cmb04.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb04.Text)
            textPart2.InsertAfter(TexteLesen(4))
            textPart3.InsertAfter(Me.txtPreis04.Text)
            textPart4.InsertAfter(Me.total04.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- MOTORRAD -----

        If Me.cmb06.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb06.Text)
            textPart2.InsertAfter(TexteLesen(6))
            textPart3.InsertAfter(Me.txtPreis06.Text)
            textPart4.InsertAfter(Me.total06.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- BEHERBERUNG ERWACHSENE -----

        If Me.cmb03.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb03.Text)
            textPart2.InsertAfter(TexteLesen(3))
            textPart3.InsertAfter(Me.txtPreis03.Text)
            textPart4.InsertAfter(Me.total03.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- BEHERBERUNG KINDER -----

        If Me.cmb16.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb16.Text)
            textPart2.InsertAfter(TexteLesen(16))
            textPart3.InsertAfter(Me.txtPreis16.Text)
            textPart4.InsertAfter(Me.total16.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- BUNGALOW -----

        If Me.cmb15.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb15.Text)
            textPart2.InsertAfter(TexteLesen(15))
            textPart3.InsertAfter(Me.txtPreis15.Text)
            textPart4.InsertAfter(Me.total15.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- UEBERNACHTUNG AUTO -----

        If Me.cmb05.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb05.Text)
            textPart2.InsertAfter(TexteLesen(5))
            textPart3.InsertAfter(Me.txtPreis05.Text)
            textPart4.InsertAfter(Me.total05.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- ZELT KLEIN -----

        If Me.cmb07.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb07.Text)
            textPart2.InsertAfter(TexteLesen(7))
            textPart3.InsertAfter(Me.txtPreis07.Text)
            textPart4.InsertAfter(Me.total07.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- ZELT GROSS -----

        If Me.cmb08.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb08.Text)
            textPart2.InsertAfter(TexteLesen(8))
            textPart3.InsertAfter(Me.txtPreis08.Text)
            textPart4.InsertAfter(Me.total08.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If




        '------- WOHNWAGEN -----

        If Me.cmb09.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb09.Text)
            textPart2.InsertAfter(TexteLesen(9))
            textPart3.InsertAfter(Me.txtPreis09.Text)
            textPart4.InsertAfter(Me.total09.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- WOHNMOBIL -----

        If Me.cmb10.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb10.Text)
            textPart2.InsertAfter(TexteLesen(10))
            textPart3.InsertAfter(Me.txtPreis10.Text)
            textPart4.InsertAfter(Me.total10.Text)

            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- STROM -----

        If Me.cmb12.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb12.Text)
            textPart2.InsertAfter(TexteLesen(12))
            textPart3.InsertAfter(Me.txtPreis12.Text)
            textPart4.InsertAfter(Me.total12.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- STROM FUNKER -----

        If Me.cmb13.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb13.Text)
            textPart2.InsertAfter(TexteLesen(13))
            textPart3.InsertAfter(Me.txtPreis13.Text)
            textPart4.InsertAfter(Me.total13.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- GEBUEHREN -----

        If Me.cmb14.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb14.Text)
            textPart2.InsertAfter(TexteLesen(14))
            textPart3.InsertAfter(Me.txtPreis14.Text)
            textPart4.InsertAfter(Me.total14.Text)


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


        '------- ERWACHSENE -----

        If Me.cmb01.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb01.Text)
            textPart2.InsertAfter(TexteLesen(1))
            textPart3.InsertAfter(Me.txtPreis01.Text)
            textPart4.InsertAfter(Me.total01.Text)

            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- KINDER -----

        If Me.cmb02.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb02.Text)
            textPart2.InsertAfter(TexteLesen(2))
            textPart3.InsertAfter(Me.txtPreis02.Text)
            textPart4.InsertAfter(Me.total02.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- HUND -----

        If Me.cmb11.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb11.Text)
            textPart2.InsertAfter(TexteLesen(11))
            textPart3.InsertAfter(Me.txtPreis11.Text)
            textPart4.InsertAfter(Me.total11.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- AUTO -----

        If Me.cmb04.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb04.Text)
            textPart2.InsertAfter(TexteLesen(4))
            textPart3.InsertAfter(Me.txtPreis04.Text)
            textPart4.InsertAfter(Me.total04.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- MOTORRAD -----

        If Me.cmb06.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb06.Text)
            textPart2.InsertAfter(TexteLesen(6))
            textPart3.InsertAfter(Me.txtPreis06.Text)
            textPart4.InsertAfter(Me.total06.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- BEHERBERUNG ERWACHSENE -----

        If Me.cmb03.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb03.Text)
            textPart2.InsertAfter(TexteLesen(3))
            textPart3.InsertAfter(Me.txtPreis03.Text)
            textPart4.InsertAfter(Me.total03.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- BEHERBERUNG KINDER -----

        If Me.cmb16.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb16.Text)
            textPart2.InsertAfter(TexteLesen(16))
            textPart3.InsertAfter(Me.txtPreis16.Text)
            textPart4.InsertAfter(Me.total16.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- BUNGALOW -----

        If Me.cmb15.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb15.Text)
            textPart2.InsertAfter(TexteLesen(15))
            textPart3.InsertAfter(Me.txtPreis15.Text)
            textPart4.InsertAfter(Me.total15.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- UEBERNACHTUNG AUTO -----

        If Me.cmb05.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb05.Text)
            textPart2.InsertAfter(TexteLesen(5))
            textPart3.InsertAfter(Me.txtPreis05.Text)
            textPart4.InsertAfter(Me.total05.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- ZELT KLEIN -----

        If Me.cmb07.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb07.Text)
            textPart2.InsertAfter(TexteLesen(7))
            textPart3.InsertAfter(Me.txtPreis07.Text)
            textPart4.InsertAfter(Me.total07.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- ZELT GROSS -----

        If Me.cmb08.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb08.Text)
            textPart2.InsertAfter(TexteLesen(8))
            textPart3.InsertAfter(Me.txtPreis08.Text)
            textPart4.InsertAfter(Me.total08.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If




        '------- WOHNWAGEN -----

        If Me.cmb09.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb09.Text)
            textPart2.InsertAfter(TexteLesen(9))
            textPart3.InsertAfter(Me.txtPreis09.Text)
            textPart4.InsertAfter(Me.total09.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- WOHNMOBIL -----

        If Me.cmb10.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb10.Text)
            textPart2.InsertAfter(TexteLesen(10))
            textPart3.InsertAfter(Me.txtPreis10.Text)
            textPart4.InsertAfter(Me.total10.Text)

            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- STROM -----

        If Me.cmb12.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb12.Text)
            textPart2.InsertAfter(TexteLesen(12))
            textPart3.InsertAfter(Me.txtPreis12.Text)
            textPart4.InsertAfter(Me.total12.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If


        '------- STROM FUNKER -----

        If Me.cmb13.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb13.Text)
            textPart2.InsertAfter(TexteLesen(13))
            textPart3.InsertAfter(Me.txtPreis13.Text)
            textPart4.InsertAfter(Me.total13.Text)


            textPart1.InsertParagraphAfter()
            textPart2.InsertParagraphAfter()
            textPart3.InsertParagraphAfter()
            textPart4.InsertParagraphAfter()

        End If

        '------- GEBUEHREN -----

        If Me.cmb14.Text <> CheckString Then

            textPart1.InsertAfter(Me.cmb14.Text)
            textPart2.InsertAfter(TexteLesen(14))
            textPart3.InsertAfter(Me.txtPreis14.Text)
            textPart4.InsertAfter(Me.total14.Text)


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



End Class