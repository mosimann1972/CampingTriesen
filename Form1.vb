
Imports System.Drawing.Printing
Imports System.IO
Imports System.Data.OleDb
Imports Microsoft.Office.Interop


Public Class Form1

    Public Conn As New OleDbConnection

    Public ErfassenMutierenLoeschen As String
    Public AnAbreiseSuchText As String

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        'MyBase.WindowState = FormWindowState.Maximized

    End Sub

		Private Function Verdoppeln(ByVal Zahl As Integer) As Integer
		Dim Doppel As Integer = 2 * Zahl
		Return Doppel
	End Function


    Private Sub ProgrammSchliessenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ProgrammSchliessenToolStripMenuItem.Click

        Me.Close()

    End Sub


	'Private Sub CheckInToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckInToolStripMenuItem.Click

	'    ErfassenMutierenLoeschen = "SUCH"

	'	Dim MDIChildForm As New AdressenSuchen()
	'    MDIChildForm.MdiParent = Me
	'    MDIChildForm.WindowState = FormWindowState.Maximized
	'    MDIChildForm.Show()

	'End Sub

	Private Sub CampingSuchenToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CampingSuchenToolStripMenuItem.Click

		ErfassenMutierenLoeschen = "SUCHCA"

		Dim MDIChildForm As New AdressenSuchen()
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()

	End Sub

	Private Sub ZimmerSuchenToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ZimmerSuchenToolStripMenuItem.Click

		ErfassenMutierenLoeschen = "SUCHZI"

		Dim MDIChildForm As New AdressenSuchen()
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()

	End Sub


	'Private Sub ErfassenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ErfassenToolStripMenuItem.Click

	'    ErfassenMutierenLoeschen = "ERBO"

	'	Dim MDIChildForm As New AdressenErfassen()
	'    MDIChildForm.MdiParent = Me
	'   MDIChildForm.WindowState = FormWindowState.Maximized
	'    MDIChildForm.Show()

	'End Sub

	Private Sub AdresseUndCampingToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AdresseUndCampingToolStripMenuItem.Click

		ErfassenMutierenLoeschen = "ERBOCA"

		Dim MDIChildForm As New AdressenErfassen()
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()


	End Sub


	Private Sub AdresseUndZimmerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AdresseUndZimmerToolStripMenuItem.Click

		ErfassenMutierenLoeschen = "ERBOZI"

		Dim MDIChildForm As New AdressenErfassen()
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()



	End Sub


	'Private Sub SuchenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles SuchenToolStripMenuItem.Click

	'	ErfassenMutierenLoeschen = "SUBO"

	'	Dim MDIChildForm As New AdressenSuchen()
	'	MDIChildForm.MdiParent = Me
	'	MDIChildForm.WindowState = FormWindowState.Maximized
	'	MDIChildForm.Show()

	'End Sub

	Private Sub AdresseSuchenFuerCampingToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AdresseSuchenFuerCampingToolStripMenuItem.Click

		ErfassenMutierenLoeschen = "SUBOCA"

		Dim MDIChildForm As New AdressenSuchen()
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()

	End Sub

	Private Sub AdresseSuchenFuerZimmerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AdresseSuchenFuerZimmerToolStripMenuItem.Click

		ErfassenMutierenLoeschen = "SUBOZI"

		Dim MDIChildForm As New AdressenSuchen()
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()


	End Sub

	Private Sub ErfassenToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ErfassenToolStripMenuItem1.Click

		ErfassenMutierenLoeschen = "ER"

		Dim MDIChildForm As New AdressenErfassen()
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()

	End Sub


	Private Sub MutierenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MutierenToolStripMenuItem.Click

		ErfassenMutierenLoeschen = "MU"

		Dim MDIChildForm As New AdressenSuchen()
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()


	End Sub

	Private Sub LöschenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles LöschenToolStripMenuItem.Click

		ErfassenMutierenLoeschen = "LO"

		Dim MDIChildForm As New AdressenSuchen
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()


	End Sub

	'Private Sub CheckInLöschenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CheckInLöschenToolStripMenuItem.Click

	'	ErfassenMutierenLoeschen = "CHLO"		'Check In löschen

	'	Dim MDIChildForm As New AdressenSuchen
	'	MDIChildForm.MdiParent = Me
	'	MDIChildForm.WindowState = FormWindowState.Maximized
	'	MDIChildForm.Show()

	'End Sub

	Private Sub CampingLoeschenToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CampingLoeschenToolStripMenuItem.Click

		ErfassenMutierenLoeschen = "CHLOCA"		'Check In löschen

		Dim MDIChildForm As New AdressenSuchen
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()


	End Sub

	Private Sub ZimmerLoeschenToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ZimmerLoeschenToolStripMenuItem.Click

		ErfassenMutierenLoeschen = "CHLOZI"		'Check In löschen

		Dim MDIChildForm As New AdressenSuchen
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()


	End Sub


	Private Sub InfoToolStripMenuItem1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles InfoToolStripMenuItem1.Click

		Dim MDIChildForm As New AboutBox1
		MDIChildForm.MdiParent = Me
		'MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()

	End Sub

	Private Sub MeldelisteErstellenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MeldelisteErstellenToolStripMenuItem.Click

		Dim MDIChildForm As New Meldewesen
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()

	End Sub

	Private Sub AnreiseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AnreiseToolStripMenuItem.Click

		AnAbreiseSuchText = "AnreiseCamping"

		Dim MDIChildForm As New AnAbreise
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()

	End Sub

	Private Sub AbreiseToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles AbreiseToolStripMenuItem.Click

		AnAbreiseSuchText = "AbreiseCamping"

		Dim MDIChildForm As New AnAbreise
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()



	End Sub


	Private Sub AnreiseZimmerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AnreiseZimmerToolStripMenuItem.Click

		AnAbreiseSuchText = "AnreiseZimmer"

		Dim MDIChildForm As New AnAbreise
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()


	End Sub

	Private Sub AbreiseZimmerToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AbreiseZimmerToolStripMenuItem.Click

		AnAbreiseSuchText = "AbreiseZimmer"

		Dim MDIChildForm As New AnAbreise
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()


	End Sub


	Private Sub BelegungToolStripMenuItem_Click_1(sender As System.Object, e As System.EventArgs) Handles BelegungToolStripMenuItem.Click

		Dim MDIChildForm As New Belegung
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()

	End Sub

	Private Sub AdminToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles AdminToolStripMenuItem.Click

		Dim MDIChildForm As New Admin
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()

	End Sub



	Public Function Laenderliste() As DataTable
		Dim objConn As New OleDbConnection
		Dim dtAdapter As OleDbDataAdapter
		Dim dt As New DataTable

		Dim strConnString As String

		strConnString = "provider=microsoft.jet.oledb.4.0;data source=camping.mdb"
		objConn = New OleDbConnection(strConnString)
		objConn.Open()

		Dim strSQL As String

		strSQL = "SELECT Land FROM tbLand"

		dtAdapter = New OleDbDataAdapter(strSQL, objConn)
		dtAdapter.Fill(dt)

		dtAdapter = Nothing

		objConn.Close()
		objConn = Nothing

		Return (dt)	'*** Return DataTable ***'   

	End Function

	Private Sub PreiseMutierenToolStripMenuItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PreiseMutierenToolStripMenuItem.Click

		Dim MDIChildForm As New PreiseMutieren
		MDIChildForm.MdiParent = Me
		MDIChildForm.WindowState = FormWindowState.Maximized
		MDIChildForm.Show()

	End Sub

	Public Function TitelErmitteln(ByVal TitelId As Integer, ByVal Sprachcode As String)


		Dim conn As New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=camping.mdb")
		Dim cmd As New OleDbCommand("Select * from tbTitel where Id = " & TitelId & "", conn)
		Dim da As New OleDbDataAdapter(cmd)
		Dim ds As New DataSet()

		da.Fill(ds, "tbTitel")

		Select Case Sprachcode
			Case "D"
				Return ds.Tables("tbTitel").Rows(0).Item(1)
			Case "E"
				Return ds.Tables("tbTitel").Rows(0).Item(2)
			Case "F"
				Return ds.Tables("tbTitel").Rows(0).Item(3)
			Case "I"
				Return ds.Tables("tbTitel").Rows(0).Item(4)
		End Select

	End Function


	Private Sub MeldeFormAnmeldungDruckenToolStripMenuItemD_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MeldeFormAnmeldungDruckenToolStripMenuItemD.Click

		Dim oWord As Word.Application
		Dim oDoc As Word.Document
		oWord = CreateObject("Word.Application")
		oWord.Visible = True
		oDoc = oWord.Documents.Add("c:\Formulare\Meldeschein_mitAnmeldungD.dot")

	End Sub

	Private Sub MeldeFormUndAnmeldungDruckenEToolStripMenuItemE_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MeldeFormUndAnmeldungDruckenEToolStripMenuItemE.Click

		Dim oWord As Word.Application
		Dim oDoc As Word.Document
		oWord = CreateObject("Word.Application")
		oWord.Visible = True
		oDoc = oWord.Documents.Add("c:\Formulare\Meldeschein_mitAnmeldungE.dot")

	End Sub

	Private Sub MeldeFormUndAnmeldungDruckenFToolStripMenuItemF_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MeldeFormUndAnmeldungDruckenFToolStripMenuItemF.Click

		Dim oWord As Word.Application
		Dim oDoc As Word.Document
		oWord = CreateObject("Word.Application")
		oWord.Visible = True
		oDoc = oWord.Documents.Add("c:\Formulare\Meldeschein_mitAnmeldungF.dot")

	End Sub

	Private Sub MeldeFormUndAnmeldungDruckenIToolStripMenuItemI_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MeldeFormUndAnmeldungDruckenIToolStripMenuItemI.Click

		Dim oWord As Word.Application
		Dim oDoc As Word.Document
		oWord = CreateObject("Word.Application")
		oWord.Visible = True
		oDoc = oWord.Documents.Add("c:\Formulare\Meldeschein_mitAnmeldungI.dot")

	End Sub


End Class
















