
Imports System.Data.OleDb
Imports System.Object
Imports System.Console

Public Class AdressenSuchen

    Dim conn As New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=camping.mdb")
    Dim AufrufArt As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        Me.AcceptButton = btnSuchen

        If Form1.ErfassenMutierenLoeschen = "MU" Then
            lblSuchTitel.Text = "Adressen suchen für Mutation"
            txtBookingId.Enabled = False
            txtPlatznummer.Enabled = False
        End If

        If Form1.ErfassenMutierenLoeschen = "LO" Then
            lblSuchTitel.Text = "Adressen suchen für Löschung"
            txtBookingId.Enabled = False
            txtPlatznummer.Enabled = False
        End If


		If Form1.ErfassenMutierenLoeschen = "SUBU" Then lblSuchTitel.Text = "Buchung suchen"
		
		If Form1.ErfassenMutierenLoeschen = "SUBOCA" Then
			lblSuchTitel.Text = "Adresse suchen für Camping"
			txtBookingId.Enabled = False
			txtPlatznummer.Enabled = False
		End If

		If Form1.ErfassenMutierenLoeschen = "SUBOZI" Then
			lblSuchTitel.Text = "Adresse suchen für Zimmer"
			txtBookingId.Enabled = False
			txtPlatznummer.Enabled = False
		End If

		If Form1.ErfassenMutierenLoeschen = "SUCHCA" Then lblSuchTitel.Text = "Camping suchen"
		If Form1.ErfassenMutierenLoeschen = "SUCHZI" Then lblSuchTitel.Text = "Zimmer suchen"
		If Form1.ErfassenMutierenLoeschen = "LOSCCA" Then lblSuchTitel.Text = "Camping suchen für Löschung"
		If Form1.ErfassenMutierenLoeschen = "LOSCZI" Then lblSuchTitel.Text = "Zimmer suchen für Löschung"

	End Sub

    
    Private Sub btnSuchen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSuchen.Click

		'Mutieren
		If Form1.ErfassenMutierenLoeschen = "SUCHCA" Then Suchen_Booking()
		If Form1.ErfassenMutierenLoeschen = "SUCHZI" Then Suchen_Zimmer()

		'Löschen
		If Form1.ErfassenMutierenLoeschen = "CHLOCA" Then Suchen_Booking()
		If Form1.ErfassenMutierenLoeschen = "CHLOZI" Then Suchen_Zimmer()
		

		If Form1.ErfassenMutierenLoeschen = "SUBOCA" Then Suchen_Adressen()
		If Form1.ErfassenMutierenLoeschen = "SUBOZI" Then Suchen_Adressen()


		If Form1.ErfassenMutierenLoeschen = "MU" Then Suchen_Adressen()
		If Form1.ErfassenMutierenLoeschen = "LO" Then Suchen_Adressen()


	End Sub


    Private Sub Suchen_Adressen()

        Dim begleitperson As Boolean = chkBegleitpersonen.Checked

        Dim strSQL As String

        If begleitperson = True Then
            strSQL = "SELECT AdressId,Nachname,Vorname,Adresse,Ort,Geburtsdatum FROM tbAdress where Suchfeld Like '%" & txtSearchString.Text & "%' Order by Nachname,Vorname,Ort "
        Else
            strSQL = "SELECT AdressId,Nachname,Vorname,Adresse,Ort,Geburtsdatum FROM tbAdress where Suchfeld Like '%" & txtSearchString.Text & "%' and Leader = true Order by Nachname,Vorname,Ort "
        End If

        Dim da As New OleDbDataAdapter(strSQL, conn)
        Dim ds As New DataSet()


        DataGridView1.Refresh()
        DataGridView1.AllowUserToResizeColumns = False
        DataGridView1.GridColor = Color.Black

        da.Fill(ds, "tbAdress")
        DataGridView1.DataSource = ds
        DataGridView1.DataMember = "tbAdress"


    End Sub

    Private Sub Suchen_Booking()


        Dim TrueFlag As Boolean

        'Aktiv = true   => alle aktiven Buchungen werden angezeigt
        'Aktiv = false  => alle verrechneten Buchungen werden angezeigt

        If chkAktiv.Checked = True Then
            TrueFlag = False
        Else
            TrueFlag = True
        End If


        If txtPlatznummer.Text <> "" Then
            Dim da2 As New OleDbDataAdapter("SELECT tbBookingtbAdress.BookingId,Nachname," _
                & "Vorname,Ort,Anreise,Abreise,tbBooking.Aktiv " _
                & "FROM tbBooking INNER JOIN (tbAdress INNER JOIN tbBookingtbAdress ON " _
                & "tbAdress.AdressId = tbBookingtbAdress.AdressId) ON " _
                & "tbBooking.BookingId = tbBookingtbAdress.BookingId where " _
                & "tbBooking.Platznummer = '" & txtPlatznummer.Text & "'" _
                & " and tbBooking.Aktiv = " & TrueFlag & "" _
                & " and tbBookingtbAdress.Leader = true", conn)
            Dim ds2 As New DataSet()

            DataGridView1.Refresh()
            DataGridView1.AllowUserToResizeColumns = False
            DataGridView1.GridColor = Color.Black

            da2.Fill(ds2, "tbBooking")
            DataGridView1.DataSource = ds2
            DataGridView1.DataMember = "tbBooking"
            Exit Sub
        End If


        If txtBookingId.Text <> "" Then
            Dim da1 As New OleDbDataAdapter("SELECT tbBookingtbAdress.BookingId,Nachname," _
                & "Vorname,Ort,Anreise,Abreise,tbBooking.Aktiv " _
                & "FROM tbBooking INNER JOIN (tbAdress INNER JOIN tbBookingtbAdress ON " _
                & "tbAdress.AdressId = tbBookingtbAdress.AdressId) ON " _
                & "tbBooking.BookingId = tbBookingtbAdress.BookingId where " _
                & " tbBooking.BookingId = " & txtBookingId.Text & "" _
                & " and tbBooking.Aktiv = " & TrueFlag & "" _
                & " and tbBookingtbAdress.Leader = true", conn)
            Dim ds1 As New DataSet()

            DataGridView1.Refresh()
            DataGridView1.AllowUserToResizeColumns = False
            DataGridView1.GridColor = Color.Black

            da1.Fill(ds1, "tbBooking")
            DataGridView1.DataSource = ds1
            DataGridView1.DataMember = "tbBooking"
            Exit Sub
        End If



        'If txtSearchString.Text <> "" Then
        Dim da As New OleDbDataAdapter("SELECT tbBookingtbAdress.BookingId,Nachname," _
                & "Vorname,Ort,Anreise,Abreise,tbBooking.Aktiv " _
                & "FROM tbBooking INNER JOIN (tbAdress INNER JOIN tbBookingtbAdress ON " _
                & "tbAdress.AdressId = tbBookingtbAdress.AdressId) ON " _
                & "tbBooking.BookingId = tbBookingtbAdress.BookingId where Suchfeld " _
                & "Like '%" & txtSearchString.Text & "%' and tbBookingtbAdress.Leader = true and tbBooking.Aktiv = " & TrueFlag & "", conn)
        Dim ds As New DataSet()


        DataGridView1.Refresh()
        DataGridView1.AllowUserToResizeColumns = False
        DataGridView1.GridColor = Color.Black

        da.Fill(ds, "tbBooking")
        DataGridView1.DataSource = ds
		DataGridView1.DataMember = "tbBooking"

		'End If



    End Sub

	Private Sub Suchen_Zimmer()

		Dim TrueFlag As Boolean

		'Aktiv = true   => alle aktiven Buchungen werden angezeigt
		'Aktiv = false  => alle verrechneten Buchungen werden angezeigt

		If chkAktiv.Checked = True Then
			TrueFlag = False
		Else
			TrueFlag = True
		End If


		If txtPlatznummer.Text <> "" Then
			Dim da2 As New OleDbDataAdapter("SELECT tbZimmertbAdress.BookingId,Nachname," _
				& "Vorname,Ort,Anreise,Abreise,tbZimmer.Aktiv " _
				& "FROM tbZimmer INNER JOIN (tbAdress INNER JOIN tbZimmertbAdress ON " _
				& "tbAdress.AdressId = tbZimmertbAdress.AdressId) ON " _
				& "tbZimmer.BookingId = tbZimmertbAdress.BookingId where " _
				& " and tbZimmer.Aktiv = " & TrueFlag & "" _
				& " and tbZimmertbAdress.Leader = true", conn)
			Dim ds2 As New DataSet()

			DataGridView1.Refresh()
			DataGridView1.AllowUserToResizeColumns = False
			DataGridView1.GridColor = Color.Black

			da2.Fill(ds2, "tbZimmer")
			DataGridView1.DataSource = ds2
			DataGridView1.DataMember = "tbZimmer"
			Exit Sub
		End If


		If txtBookingId.Text <> "" Then
			Dim da1 As New OleDbDataAdapter("SELECT tbZimmertbAdress.BookingId,Nachname," _
				& "Vorname,Ort,Anreise,Abreise,tbZimmer.Aktiv " _
				& "FROM tbZimmer INNER JOIN (tbAdress INNER JOIN tbZimmertbAdress ON " _
				& "tbAdress.AdressId = tbZimmertbAdress.AdressId) ON " _
				& "tbZimmer.BookingId = tbZimmertbAdress.BookingId where " _
				& " tbZimmer.BookingId = " & txtBookingId.Text & "" _
				& " and tbZimmer.Aktiv = " & TrueFlag & "" _
				& " and tbZimmertbAdress.Leader = true", conn)
			Dim ds1 As New DataSet()

			DataGridView1.Refresh()
			DataGridView1.AllowUserToResizeColumns = False
			DataGridView1.GridColor = Color.Black

			da1.Fill(ds1, "tbZimmer")
			DataGridView1.DataSource = ds1
			DataGridView1.DataMember = "tbZimmer"
			Exit Sub
		End If



		'If txtSearchString.Text <> "" Then
		Dim da As New OleDbDataAdapter("SELECT tbZimmertbAdress.BookingId,Nachname," _
				& "Vorname,Ort,Anreise,Abreise,tbZimmer.Aktiv " _
				& "FROM tbZimmer INNER JOIN (tbAdress INNER JOIN tbZimmertbAdress ON " _
				& "tbAdress.AdressId = tbZimmertbAdress.AdressId) ON " _
				& "tbZimmer.BookingId = tbZimmertbAdress.BookingId where Suchfeld " _
				& "Like '%" & txtSearchString.Text & "%' and tbZimmertbAdress.Leader = true and tbZimmer.Aktiv = " & TrueFlag & "", conn)
		Dim ds As New DataSet()


		DataGridView1.Refresh()
		DataGridView1.AllowUserToResizeColumns = False
		DataGridView1.GridColor = Color.Black

		da.Fill(ds, "tbZimmer")
		DataGridView1.DataSource = ds
		DataGridView1.DataMember = "tbZimmer"
		'End If



	End Sub


    Private Sub btnLoeschen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoeschen.Click

        Me.Close()

    End Sub

    Private Sub DataGridView1_CellClick(ByVal sender As Object, ByVal e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellClick

        Dim i As Integer = DataGridView1.CurrentRow.Index
        Dim Id As String = Me.DataGridView1.Item(0, i).Value.ToString   'Entweder AdressId oder BookingId


        If Id = "" Then
            MsgBox("Kein Datensatz ausgewählt", MsgBoxStyle.Exclamation)
            Exit Sub
        End If


		If Form1.ErfassenMutierenLoeschen = "CHLOCA" Then	  'Camping löschen

			Dim ClassLoeschen As New Class1
			ClassLoeschen.StartLoeschenCamping(Id)
			Me.Close()
			Exit Sub

		End If


		If Form1.ErfassenMutierenLoeschen = "CHLOZI" Then	  'Zimmer löschen

			Dim ClassLoeschen As New Class1
			ClassLoeschen.StartLoeschenZimmer(Id)
			Me.Close()
			Exit Sub

		End If


		If Form1.ErfassenMutierenLoeschen = "SUBOCA" Then	  'Adresse suchen und camping

			Dim MDIChildForm1 As New CheckIn01()
			MDIChildForm1.MdiParent = Form1
			MDIChildForm1.WindowState = FormWindowState.Maximized

			MDIChildForm1.StartBooking(Id)
			Me.Close()
			MDIChildForm1.Show()
			Exit Sub

		End If


		If Form1.ErfassenMutierenLoeschen = "SUBOZI" Then		'Adresse suchen und zimmer

			Dim MDIChildForm1 As New CheckIn02()
			MDIChildForm1.MdiParent = Form1
			MDIChildForm1.WindowState = FormWindowState.Maximized

			MDIChildForm1.StartBooking(Id)
			Me.Close()
			MDIChildForm1.Show()
			Exit Sub

		End If



		If Form1.ErfassenMutierenLoeschen = "SUCHCA" Then  'Eingecheckte Camping Suchen

			Dim MDIChildForm2 As New CheckIn01()
			MDIChildForm2.MdiParent = Form1
			MDIChildForm2.WindowState = FormWindowState.Maximized

			MDIChildForm2.StartMutation(Id)
			Me.Close()
			MDIChildForm2.Show()
			Exit Sub

		End If



		If Form1.ErfassenMutierenLoeschen = "SUCHZI" Then	 'Eingecheckte Zimmer Suchen

			Dim MDIChildForm2 As New CheckIn02()
			MDIChildForm2.MdiParent = Form1
			MDIChildForm2.WindowState = FormWindowState.Maximized

			MDIChildForm2.StartMutation(Id)
			Me.Close()
			MDIChildForm2.Show()
			Exit Sub

		End If


        Dim MDIChildForm As New AdressenErfassen()          'Adresse mutieren
        MDIChildForm.MdiParent = Form1
        MDIChildForm.WindowState = FormWindowState.Maximized

        MDIChildForm.AdresseMutieren(Id)
        Me.Close()
        MDIChildForm.Show()


    End Sub

End Class