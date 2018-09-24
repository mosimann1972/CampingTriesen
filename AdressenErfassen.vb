
Imports System.Data.OleDb
Imports System.Console
Imports System.Object
Imports System.IO
Imports Microsoft.Office.Interop


Public Class AdressenErfassen

    Inherits System.Windows.Forms.Form
    Dim conn As New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=camping.mdb")

    Public AufrufVon As String = ""
    Public AdressIdPublic As Integer
    Public CheckAdressOK As Boolean

    Public MerkenLandName1 As String
    Public MerkenLandName2 As String

    Public SprachCode As String


    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        If AdressIdPublic = 0 Then
            RadioButton1.Checked = True
        End If

        If (Form1.ErfassenMutierenLoeschen) = "ER" Then               'Erfassen
            btnSpeichern.Text = "Speichern"
        End If

        If (Form1.ErfassenMutierenLoeschen) = "LO" Then               'Loschen
            btnSpeichern.Text = "Löschen"
        End If

		If (Form1.ErfassenMutierenLoeschen) = "ERBOCA" Then				'Speichern und Camping
			btnSpeichern.Text = "weiter ==>"
		End If

		If (Form1.ErfassenMutierenLoeschen) = "ERBOZI" Then			  'Speichern und Zimmer
			btnSpeichern.Text = "weiter ==>"
		End If


        '-------- LAND

        Dim dt1 As DataTable
        dt1 = Form1.Laenderliste()
        cmbLand.DataSource = dt1
        cmbLand.DisplayMember = "Land"

        If MerkenLandName1 <> "" Then
            cmbLand.Text = MerkenLandName1
        End If
        MerkenLandName1 = ""


        '-------- NATION

        Dim dt2 As DataTable
        dt2 = Form1.Laenderliste()
        cmbNation.DataSource = dt2
        cmbNation.DisplayMember = "Land"

        If MerkenLandName2 <> "" Then
            cmbNation.Text = MerkenLandName2
        End If
        MerkenLandName2 = ""


    End Sub


    Private Sub btnSpeichern_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSpeichern.Click

        If (Form1.ErfassenMutierenLoeschen) = "ER" Then
            Speichern()
        End If

        If (Form1.ErfassenMutierenLoeschen) = "MU" Then
            updaten()
        End If

        If (Form1.ErfassenMutierenLoeschen) = "LO" Then
            loeschen()
        End If

		If (Form1.ErfassenMutierenLoeschen) = "ERBOCA" Then
			CheckAdressOK = True
			Speichern()
			If CheckAdressOK = True Then weiter()
		End If

		If (Form1.ErfassenMutierenLoeschen) = "ERBOZI" Then
			CheckAdressOK = True
			Speichern()
			If CheckAdressOK = True Then weiter()
		End If

        Me.Close()

    End Sub


    Private Sub Speichern()

        If txtNachName.Text = "" Then
            MsgBox("Das Feld Nachname darf nicht leer sein!", MsgBoxStyle.Critical)
            txtNachName.Focus()
            CheckAdressOK = False
            Exit Sub
        End If

        If txtVorname.Text = "" Then
            MsgBox("Das Feld Vorname darf nicht leer sein!", MsgBoxStyle.Critical)
            txtVorname.Focus()
            CheckAdressOK = False
            Exit Sub
        End If

        Dim strAusweisArt As String = ""

        If txtGeburtsdatum.Text = "" Then
            txtGeburtsdatum.Text = "01.01.1900"
        End If

        If txtId.Checked = True Then
            strAusweisArt = "Id"
        End If

        If txtPass.Checked = True Then
            strAusweisArt = "Pass"
        End If

        Dim strSuchfeld As String = txtNachName.Text & " " & txtVorname.Text & " " & txtAdresse.Text & " " & txtPLZ.Text & " " & txtOrt.Text & " " & txtGeburtsdatum.Text & "  " & txtEMailAdresse.Text

        Dim cmd As New OleDbCommand("Insert Into tbAdress (NachName," _
        & "Vorname, " _
        & "Adresse, " _
        & "PLZ, " _
        & "Ort, " _
        & "Land, " _
        & "Nation, " _
        & "Geburtsdatum, " _
        & "EMailAdresse, " _
        & "AusweisArt, " _
        & "AusweisNummer, " _
        & "Suchfeld, " _
        & "Leader, " _
        & "Sprache " _
        & ") Values('" & txtNachName.Text & "'" _
        & ", '" & txtVorname.Text & "'" _
        & ", '" & txtAdresse.Text & "'" _
        & ", '" & txtPLZ.Text & "'" _
        & ", '" & txtOrt.Text & "'" _
        & ", '" & cmbLand.Text & "'" _
        & ", '" & cmbNation.Text & "'" _
        & ", '" & txtGeburtsdatum.Text & "'" _
        & ", '" & txtEMailAdresse.Text & "'" _
        & ", '" & strAusweisArt.ToString & "'" _
        & ", '" & txtAusweisNummer.Text & "'" _
        & ", '" & strSuchfeld.ToString & "',true,'" & SpracheFestlegen() & "')", conn)

        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet("tbAdress")

        Try
            da.Fill(ds, "tbAdress")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        MsgBox("Adresse erfasst", MsgBoxStyle.Exclamation)

    End Sub

    Private Sub updaten()

        Dim strAusweisArt As String = ""

        If txtGeburtsdatum.Text = "" Then
            txtGeburtsdatum.Text = "01.01.1900"
        End If

        If txtId.Checked = True Then
            strAusweisArt = "Id"
        End If

        If txtPass.Checked = True Then
            strAusweisArt = "Pass"
        End If

        Dim strSuchfeld As String = txtNachName.Text & " " & txtVorname.Text & " " & txtAdresse.Text & " " & txtPLZ.Text & " " & txtOrt.Text & " " & txtGeburtsdatum.Text & "  " & txtEMailAdresse.Text


        Dim cmd As New OleDbCommand("UPDATE tbAdress SET NachName = '" & txtNachName.Text & "'" _
                    & ", Vorname = '" & txtVorname.Text & "'" _
                    & ", Adresse = '" & txtAdresse.Text & "'" _
                    & ", PLZ = '" & txtPLZ.Text & "'" _
                    & ", Ort = '" & txtOrt.Text & "'" _
                    & ", Land = '" & cmbLand.Text & "'" _
                    & ", Nation = '" & cmbNation.Text & "'" _
                    & ", Geburtsdatum = '" & txtGeburtsdatum.Text & "'" _
                    & ", EMailAdresse = '" & txtEMailAdresse.Text & "'" _
                    & ", AusweisNummer = '" & txtAusweisNummer.Text & "'" _
                    & ", AusweisArt = '" & strAusweisArt.ToString & "'" _
                    & ", Sprache = '" & SpracheFestlegen() & "'" _
                    & ", Suchfeld = '" & strSuchfeld.ToString & "'" _
            & " WHERE AdressId = " & AdressIdPublic & "", conn)

        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet("tbAdress")

        Try
            da.Fill(ds, "tbAdress")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        MsgBox("Adresse mutiert", MsgBoxStyle.Exclamation)

    End Sub

    Private Sub loeschen()

        If MsgBox("Wollen Sie wirklich löschen?", vbOKCancel + MsgBoxStyle.Critical) = vbCancel Then
            Exit Sub
        End If

        Dim E As String

		Check_BookingVorhanden(E)
		Check_ZimmerVorhanden(E)

        If E <> "-1" Then
			MsgBox("Adresse kann nicht gelöscht werden, da Buchungen/Zimmer vorhanden!", MsgBoxStyle.Exclamation)
            Exit Sub
        End If

        Dim cmd As New OleDbCommand("Delete * From tbAdress Where AdressId = " & AdressIdPublic & "", conn)
        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet("tbAdress")

        Try
            da.Fill(ds, "tbAdress")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

        MsgBox("Adresse gelöscht", MsgBoxStyle.Exclamation)

        Me.Close()
    End Sub

	Private Sub Check_BookingVorhanden(ByRef E As String)

		Dim da As New OleDbDataAdapter("SELECT * FROM tbBookingtbAdress where tbBookingtbAdress.AdressId = " & AdressIdPublic & "", conn)
		Dim ds As New DataSet()
		da.Fill(ds, "tbBookingtbAdress")

		Me.BindingContext(ds, "tbBookingtbAdress").Position = Me.BindingContext(ds, "tbBookingtbAdress").Count - 1

		Debug.Print(Me.BindingContext(ds, "tbBookingtbAdress").Position)

		E = Me.BindingContext(ds, "tbBookingtbAdress").Position

	End Sub

	Private Sub Check_ZimmerVorhanden(ByRef E As String)

		Dim da As New OleDbDataAdapter("SELECT * FROM tbZimmertbAdress where tbZimmertbAdress.AdressId = " & AdressIdPublic & "", conn)
		Dim ds As New DataSet()
		da.Fill(ds, "tbZimmertbAdress")

		Me.BindingContext(ds, "tbZimmertbAdress").Position = Me.BindingContext(ds, "tbZimmertbAdress").Count - 1

		Debug.Print(Me.BindingContext(ds, "tbZimmertbAdress").Position)

		E = Me.BindingContext(ds, "tbZimmertbAdress").Position

	End Sub


	Private Sub btnAbbrechen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnAbbrechen.Click

		Me.Close()

	End Sub


	Private Sub weiter()

		Dim cmd As New OleDbCommand("Select * from tbAdress Order by AdressId", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet()

		Dim k As Integer
		Dim w As Integer

		da.Fill(ds, "tbAdress")

		Me.BindingContext(ds, "tbAdress").Position = Me.BindingContext(ds, "tbAdress").Count - 1

		k = Me.BindingContext(ds, "tbAdress").Position
		w = ds.Tables("tbAdress").Rows(k).Item(0)

		If Form1.ErfassenMutierenLoeschen = "ERBOCA" Then

			Dim MDIChildForm As New CheckIn01()
			MDIChildForm.MdiParent = Form1
			MDIChildForm.WindowState = FormWindowState.Maximized

			MDIChildForm.StartBooking(w)
			Me.Close()
			MDIChildForm.Show()

		End If


		If Form1.ErfassenMutierenLoeschen = "ERBOZI" Then

			Dim MDIChildForm As New CheckIn02()
			MDIChildForm.MdiParent = Form1
			MDIChildForm.WindowState = FormWindowState.Maximized

			MDIChildForm.StartBooking(w)
			Me.Close()
			MDIChildForm.Show()

		End If

	End Sub


	Public Sub AdresseMutieren(ByVal AdressId As Integer)

		Dim cmd As New OleDbCommand("Select * from tbAdress where AdressId = " & AdressId & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet()

		AdressIdPublic = AdressId

		conn.Open()
		da.Fill(ds, "tbAdress")

		txtNachName.Text = ds.Tables("tbAdress").Rows(0).Item(1)
		txtVorname.Text = ds.Tables("tbAdress").Rows(0).Item(2)

		If CheckDBNull(ds.Tables("tbAdress").Rows(0).Item(3)) <> "" Then txtAdresse.Text = ds.Tables("tbAdress").Rows(0).Item(3)
		If CheckDBNull(ds.Tables("tbAdress").Rows(0).Item(4)) <> "" Then txtPLZ.Text = ds.Tables("tbAdress").Rows(0).Item(4)
		If CheckDBNull(ds.Tables("tbAdress").Rows(0).Item(5)) <> "" Then txtOrt.Text = ds.Tables("tbAdress").Rows(0).Item(5)

		'------ LAND

		If CheckDBNull(ds.Tables("tbAdress").Rows(0).Item(6)) <> "" Then
			cmbLand.Text = ds.Tables("tbAdress").Rows(0).Item(6)
			MerkenLandName1 = ds.Tables("tbAdress").Rows(0).Item(6)
		End If

		'------ NATION

		If CheckDBNull(ds.Tables("tbAdress").Rows(0).Item(7)) <> "" Then
			cmbNation.Text = ds.Tables("tbAdress").Rows(0).Item(7)
			MerkenLandName2 = ds.Tables("tbAdress").Rows(0).Item(7)
		End If


		txtGeburtsdatum.Text = ds.Tables("tbAdress").Rows(0).Item(8)
		If CheckDBNull(ds.Tables("tbAdress").Rows(0).Item(9)) <> "" Then txtEMailAdresse.Text = ds.Tables("tbAdress").Rows(0).Item(9)
		If CheckDBNull(ds.Tables("tbAdress").Rows(0).Item(10)) <> "" Then txtAusweisNummer.Text = ds.Tables("tbAdress").Rows(0).Item(10)

		If CheckDBNull(ds.Tables("tbAdress").Rows(0).Item(11)) <> "" Then
			If ds.Tables("tbAdress").Rows(0).Item(11) = "Id" Then
				txtId.Checked = True
			Else
				txtPass.Checked = True
			End If
		End If


		If CheckDBNull(ds.Tables("tbAdress").Rows(0).Item(14)) <> "" Then

			If ds.Tables("tbAdress").Rows(0).Item(14) = "D" Then
				RadioButton1.Checked = True
			End If
			If ds.Tables("tbAdress").Rows(0).Item(14) = "E" Then
				RadioButton2.Checked = True
			End If

			If ds.Tables("tbAdress").Rows(0).Item(14) = "F" Then
				RadioButton3.Checked = True
			End If

			If ds.Tables("tbAdress").Rows(0).Item(14) = "I" Then
				RadioButton4.Checked = True
			End If

		Else

			RadioButton1.Checked = True

		End If



	End Sub

	Public Function CheckDBNull(ByVal obj As Object, Optional ByVal ObjectType As enumObjectType = enumObjectType.StrType) As Object
		Dim objReturn As Object
		objReturn = obj
		If ObjectType = enumObjectType.StrType And IsDBNull(obj) Then
			objReturn = ""
		ElseIf ObjectType = enumObjectType.IntType And IsDBNull(obj) Then
			objReturn = 0
		ElseIf ObjectType = enumObjectType.DblType And IsDBNull(obj) Then
			objReturn = 0.0
		End If
		Return objReturn
	End Function

	Enum enumObjectType
		StrType = 0
		IntType = 1
		DblType = 2
	End Enum

	Public Function SpracheFestlegen()

		Dim strSprache As String

		If RadioButton1.Checked = True Then
			strSprache = "D"
		End If

		If RadioButton2.Checked = True Then
			strSprache = "E"
		End If

		If RadioButton3.Checked = True Then
			strSprache = "F"
		End If

		If RadioButton4.Checked = True Then
			strSprache = "I"
		End If

		Return strSprache

	End Function


	Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click

		If txtNachName.Text = "" Then
			MsgBox("Das Feld Nachname darf nicht leer sein!", MsgBoxStyle.Critical)
			txtNachName.Focus()
			Exit Sub
		End If

		If txtVorname.Text = "" Then
			MsgBox("Das Feld Vorname darf nicht leer sein!", MsgBoxStyle.Critical)
			txtVorname.Focus()
			Exit Sub
		End If

		If (Form1.ErfassenMutierenLoeschen) = "MU" Then
			updaten()
		Else
			Speichern()
		End If

		If DruckerLesen() <> "no" Then
			WordDrucken()
		Else
			Me.Close()
			Exit Sub
		End If

		Me.Close()


	End Sub


	Public Function DruckerLesen()

		Dim datName As String = "drucker.ini"
		Dim reader As StreamReader = File.OpenText(datName)

		While (reader.Peek() > -1)

			Return reader.ReadLine()
			Exit While

		End While

		reader.Close()

	End Function


	Private Sub WordDrucken()

		Dim oWord As Word.Application
		Dim oDoc As Word.Document
		Dim strAusweisArt As String
		oWord = CreateObject("Word.Application")
		oWord.Visible = True


		If RadioButton1.Checked = True Then oDoc = oWord.Documents.Add("c:\Formulare\Meldeschein_mitAnmeldungD.dot")
		If RadioButton2.Checked = True Then oDoc = oWord.Documents.Add("c:\Formulare\Meldeschein_mitAnmeldungE.dot")
		If RadioButton3.Checked = True Then oDoc = oWord.Documents.Add("c:\Formulare\Meldeschein_mitAnmeldungF.dot")
		If RadioButton4.Checked = True Then oDoc = oWord.Documents.Add("c:\Formulare\Meldeschein_mitAnmeldungI.dot")


		oDoc.Bookmarks.Item("tmName").Range.Text = Me.txtNachName.Text
		oDoc.Bookmarks.Item("tmVorname").Range.Text = Me.txtVorname.Text
		oDoc.Bookmarks.Item("tmAdresse").Range.Text = Me.txtAdresse.Text
		oDoc.Bookmarks.Item("tmLandPLZOrt").Range.Text = Me.cmbLand.Text & ", " & Me.txtPLZ.Text & " " & Me.txtOrt.Text
		oDoc.Bookmarks.Item("tmGeburtsdatum").Range.Text = Me.txtGeburtsdatum.Text
		oDoc.Bookmarks.Item("tmNation").Range.Text = Me.cmbNation.Text

		If txtId.Checked = True Then
			strAusweisArt = "Id"
		Else
			strAusweisArt = "***"
		End If

		If txtPass.Checked = True Then
			strAusweisArt = "Pass"
		Else
			strAusweisArt = "***"
		End If

		oDoc.Bookmarks.Item("tmAusweisart").Range.Text = strAusweisArt.ToString

		oDoc.Bookmarks.Item("tmAusweisNummer").Range.Text = Me.txtAusweisNummer.Text
		oDoc.Bookmarks.Item("tmDatum").Range.Text = System.DateTime.Today

		oDoc.Bookmarks.Item("tmBegleitpersonen").Range.Text = Form1.TitelErmitteln("15", SpracheFestlegen())
		oDoc.Bookmarks.Item("tmAnmeldung").Range.Text = Form1.TitelErmitteln("14", SpracheFestlegen())
		oDoc.Bookmarks.Item("tmKontrollschild").Range.Text = Form1.TitelErmitteln("3", SpracheFestlegen())
		oDoc.Bookmarks.Item("tmBemerkungen").Range.Text = Form1.TitelErmitteln("12", SpracheFestlegen())


	End Sub



End Class