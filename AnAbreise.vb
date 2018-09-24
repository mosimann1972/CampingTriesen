
Imports System.Data.OleDb
Imports System.Object
Imports System.Console
Imports System.Drawing.Printing
Imports System.IO
Imports System.Math


Public Class AnAbreise

    Private Const FontAdjustmentFactor = 1.1

    Dim conn As New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=camping.mdb")
    Dim AufrufArt As String


    Private Sub AnAbreise_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

		If Form1.AnAbreiseSuchText = "AnreiseCamping" Then
			AufrufArt = "Camping"
			lblSuchTitel1.Text = "Anreise Camping"
		End If

		If Form1.AnAbreiseSuchText = "AbreiseCamping" Then
			AufrufArt = "Camping"
			lblSuchTitel1.Text = "Abreise Camping"
		End If

		If Form1.AnAbreiseSuchText = "AnreiseZimmer" Then
			AufrufArt = "Zimmer"
			lblSuchTitel1.Text = "Anreise Zimmer"
		End If

		If Form1.AnAbreiseSuchText = "AbreiseZimmer" Then
			AufrufArt = "Zimmer"
			lblSuchTitel1.Text = "Abreise Zimmer"
		End If



    End Sub

    Private Sub btnSuchen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSuchen.Click


		If Form1.AnAbreiseSuchText = "AnreiseCamping" Then
			Suchen_Anreise_Camping()
		End If

		If Form1.AnAbreiseSuchText = "AbreiseCamping" Then
			Suchen_Abreise_Camping()
		End If


		If Form1.AnAbreiseSuchText = "AnreiseZimmer" Then
			Suchen_Anreise_Zimmer()
		End If

		If Form1.AnAbreiseSuchText = "AbreiseZimmer" Then
			Suchen_Abreise_Zimmer()
		End If


    End Sub

	Private Sub Suchen_Anreise_Camping()


		Dim strSQL As String
		Dim d As Date

		d = DateTimePicker1.Value.Date

		Dim SuchDatum As String

		SuchDatum = "#" & DateTimePicker1.Value.Month & "/" & DateTimePicker1.Value.Day & "/" & DateTimePicker1.Value.Year & "#"

		strSQL = (" SELECT tbBooking.BookingId, tbBooking.Anreise, tbBooking.Abreise, tbBooking.AnzahlNaechte, tbAdress.NachName," _
		& "tbAdress.VorName, tbAdress.Ort FROM tbBooking INNER JOIN (tbAdress INNER JOIN tbBookingtbAdress ON tbAdress.AdressId = tbBookingtbAdress.AdressId)" _
		& "ON tbBooking.BookingId = tbBookingtbAdress.BookingId WHERE (((tbBooking.Anreise)= " & SuchDatum & "));")


		Dim da As New OleDbDataAdapter(strSQL, conn)
		Dim ds As New DataSet()

		DataGridView1.Refresh()
		DataGridView1.AllowUserToResizeColumns = False
		DataGridView1.GridColor = Color.Black

		da.Fill(ds, "tbBooking")
		DataGridView1.DataSource = ds
		DataGridView1.DataMember = "tbBooking"

	End Sub


	Private Sub Suchen_Abreise_Camping()


		Dim strSQL As String
		Dim d As Date

		d = DateTimePicker1.Value.Date

		Dim SuchDatum As String

		SuchDatum = "#" & DateTimePicker1.Value.Month & "/" & DateTimePicker1.Value.Day & "/" & DateTimePicker1.Value.Year & "#"

		strSQL = (" SELECT tbBooking.BookingId, tbBooking.Anreise, tbBooking.Abreise, tbBooking.AnzahlNaechte, tbAdress.NachName," _
		& "tbAdress.VorName, tbAdress.Ort FROM tbBooking INNER JOIN (tbAdress INNER JOIN tbBookingtbAdress ON tbAdress.AdressId = tbBookingtbAdress.AdressId)" _
		& "ON tbBooking.BookingId = tbBookingtbAdress.BookingId WHERE (((tbBooking.Abreise)= " & SuchDatum & "));")


		Dim da As New OleDbDataAdapter(strSQL, conn)
		Dim ds As New DataSet()

		DataGridView1.Refresh()
		DataGridView1.AllowUserToResizeColumns = False
		DataGridView1.GridColor = Color.Black

		da.Fill(ds, "tbBooking")
		DataGridView1.DataSource = ds
		DataGridView1.DataMember = "tbBooking"

	End Sub

	Private Sub Suchen_Anreise_Zimmer()


		Dim strSQL As String
		Dim d As Date

		d = DateTimePicker1.Value.Date

		Dim SuchDatum As String

		SuchDatum = "#" & DateTimePicker1.Value.Month & "/" & DateTimePicker1.Value.Day & "/" & DateTimePicker1.Value.Year & "#"

		strSQL = (" SELECT tbZimmer.BookingId, tbZimmer.Anreise, tbZimmer.Abreise, tbZimmer.AnzahlNaechte, tbAdress.NachName," _
		& "tbAdress.VorName, tbAdress.Ort FROM tbZimmer INNER JOIN (tbAdress INNER JOIN tbZimmertbAdress ON tbAdress.AdressId = tbZimmertbAdress.AdressId)" _
		& "ON tbZimmer.BookingId = tbZimmertbAdress.BookingId WHERE (((tbZimmer.Anreise)= " & SuchDatum & "));")


		Dim da As New OleDbDataAdapter(strSQL, conn)
		Dim ds As New DataSet()

		DataGridView1.Refresh()
		DataGridView1.AllowUserToResizeColumns = False
		DataGridView1.GridColor = Color.Black

		da.Fill(ds, "tbZimmer")
		DataGridView1.DataSource = ds
		DataGridView1.DataMember = "tbZimmer"

	End Sub

	Private Sub Suchen_Abreise_Zimmer()

		Dim strSQL As String
		Dim d As Date

		d = DateTimePicker1.Value.Date

		Dim SuchDatum As String

		SuchDatum = "#" & DateTimePicker1.Value.Month & "/" & DateTimePicker1.Value.Day & "/" & DateTimePicker1.Value.Year & "#"

		strSQL = (" SELECT tbZimmer.BookingId, tbZimmer.Anreise, tbZimmer.Abreise, tbZimmer.AnzahlNaechte, tbAdress.NachName," _
		& "tbAdress.VorName, tbAdress.Ort FROM tbZimmer INNER JOIN (tbAdress INNER JOIN tbZimmertbAdress ON tbAdress.AdressId = tbZimmertbAdress.AdressId)" _
		& "ON tbZimmer.BookingId = tbZimmertbAdress.BookingId WHERE (((tbZimmer.Abreise)= " & SuchDatum & "));")


		Dim da As New OleDbDataAdapter(strSQL, conn)
		Dim ds As New DataSet()

		DataGridView1.Refresh()
		DataGridView1.AllowUserToResizeColumns = False
		DataGridView1.GridColor = Color.Black

		da.Fill(ds, "tbZimmer")
		DataGridView1.DataSource = ds
		DataGridView1.DataMember = "tbZimmer"

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

		If AufrufArt = "Camping" Then

			Dim MDIChildForm2 As New CheckIn01()
			MDIChildForm2.MdiParent = Form1
			MDIChildForm2.WindowState = FormWindowState.Maximized

			MDIChildForm2.StartMutation(Id)
			Me.Close()
			MDIChildForm2.Show()
			Exit Sub

		End If

		If AufrufArt = "Zimmer" Then

			Dim MDIChildForm2 As New CheckIn02()
			MDIChildForm2.MdiParent = Form1
			MDIChildForm2.WindowState = FormWindowState.Maximized

			MDIChildForm2.StartMutation(Id)
			Me.Close()
			MDIChildForm2.Show()
			Exit Sub

		End If



	End Sub


    Private Sub btnPrint_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnPrint.Click

        StartDruckenAnAbreise()

    End Sub


    Public Sub StartDruckenAnAbreise()

        'DokumentTitel = "Buchungsbestätigung"

        If DruckerLesen() <> "no" Then
            PrintDocument2.DefaultPageSettings.PrinterSettings.PrinterName = DruckerLesen()
        Else
            Exit Sub
        End If

        Try
            PrintDocument2.Print()
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try


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


    Private Sub PrintDocument2_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument2.PrintPage

        Dim bm As New Bitmap(Me.DataGridView1.Width, Me.DataGridView1.Height)

        DataGridView1.DrawToBitmap(bm, New Rectangle(0, 0, Me.DataGridView1.Width, Me.DataGridView1.Height))

        Dim g As Graphics = e.Graphics

        g.DrawString(lblSuchTitel1.Text, New Font("Arial", 6, FontStyle.Bold, GraphicsUnit.Millimeter), Brushes.Black, 40, 10)
        g.DrawString(DateTimePicker1.Value.Date, New Font("Arial", 6, FontStyle.Bold, GraphicsUnit.Millimeter), Brushes.Black, 160, 10)

        e.Graphics.DrawImage(bm, 0, 70)

    End Sub



    'Private Sub GridViewPrintDocument_PrintPage(ByVal sender As Object, ByVal e As PrintPageEventArgs)
    '    PrintGridView(e.Graphics)
    'End Sub

    'Private Sub btnPrint_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPrint.Click

    'Dim GridViewPrintDocument As New Printing.PrintDocument
    '    AddHandler GridViewPrintDocument.PrintPage, _
    '    AddressOf GridViewPrintDocument_PrintPage
    '    GridViewPrintDocument.Print()

    'End Sub
    'Private Sub PrintGridView(ByVal GxPrint As Graphics)

    'DrawGridViewBox(GxPrint)
    'DrawGridViewHeader(GxPrint)
    '    DrawGridViewRows(GxPrint)
    'End Sub
    'Private Sub DrawGridViewHeader(ByVal GxPrint As Graphics)

    'Dim CellText = String.Empty
    'Dim StartTop = DataGridView1.Top
    'Dim PrintFont As New Font(New FontFamily("Microsoft Sans Serif"), 10, FontStyle.Bold)
    'Dim StartLeft = DataGridView1.Left
    '   For Each PrintCol As DataGridViewColumn In DataGridView1.Columns
    '      CellText = PrintCol.HeaderText
    '     GxPrint.DrawString(CellText, PrintFont, Brushes.Gray, _
    '                           StartLeft, StartTop)
    '    GxPrint.DrawLine(Pens.Black, StartLeft, StartTop, _
    '                     StartLeft, StartTop + DataGridView1.Rows(0).Height)
    '   StartLeft += PrintCol.Width * FontAdjustmentFactor
    'Next
    'GxPrint.DrawLine(Pens.Black, DataGridView1.Left, _
    '                StartTop + DataGridView1.Rows(0).Height, _
    '               CInt(DataGridView1.Width * FontAdjustmentFactor), _
    '              StartTop + DataGridView1.Rows(0).Height)
    'End Sub

    'Private Sub DrawGridViewRows(ByVal GxPrint As Graphics)
    'Dim RowIndex = 1 'since header is used so we start with 1
    'Dim PrintFont As New Font(New FontFamily("Microsoft Sans Serif"), _
    '                                  10, FontStyle.Regular)
    '    For Each PrintRow As DataGridViewRow In DataGridView1.Rows
    'Dim StartTop = DataGridView1.Top + (RowIndex * PrintRow.Height)
    'Dim ColIndex = 0
    'Dim StartLeft = DataGridView1.Left
    '        For Each PrintCell As DataGridViewCell In PrintRow.Cells
    '
    '               StartLeft *= FontAdjustmentFactor
    '  Dim CellText = String.Empty
    '            If (Not IsDBNull(PrintCell.Value)) Then CellText = PrintCell.Value
    '
    '           GxPrint.DrawString(CellText, PrintFont, Brushes.Gray, _
    '                              StartLeft, StartTop)
    '         GxPrint.DrawLine(Pens.Black, StartLeft, StartTop, _
    '                           StartLeft, StartTop + PrintRow.Height)
    '           StartLeft += DataGridView1.Columns(ColIndex).Width
    '           ColIndex += 1
    '       Next
    '       GxPrint.DrawLine(Pens.Black, DataGridView1.Left, _
    '                        StartTop + PrintRow.Height, _
    '                        CInt(DataGridView1.Width * FontAdjustmentFactor), _
    '                        StartTop + PrintRow.Height)
    '       RowIndex += 1
    '   Next
    'End Sub

    'Private Sub DrawGridViewBox(ByVal GxPrint As Graphics)
    ' Dim GridViewRect As New Rectangle( _
    '             DataGridView1.Left, DataGridView1.Top, _
    '            DataGridView1.Width * FontAdjustmentFactor, _
    '             DataGridView1.Height * FontAdjustmentFactor)
    '
    '    GxPrint.DrawRectangle(Pens.Black, GridViewRect)
    'End Sub












End Class








