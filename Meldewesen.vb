
Imports System.Data.OleDb
Imports System.Object
Imports System.Console
Imports System.Drawing.Printing
Imports System.IO
Imports System.Math



Public Class Meldewesen

    Dim conn As New OleDbConnection("provider=microsoft.jet.oledb.4.0;data source=camping.mdb")

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

        lblSuchTitel.Text = "Meldeliste erstellen"

    End Sub


    Private Sub btnLoeschen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoeschen.Click

        Me.Close()

    End Sub

    Private Sub btnSuchen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnSuchen.Click

        Dim strSQL As String

        strSQL = "SELECT tbAdress.NachName, tbAdress.VorName, tbAdress.Geburtsdatum, tbAdress.Nation FROM tbAdress where MeldescheinPrint = false"

        Dim da As New OleDbDataAdapter(strSQL, conn)
        Dim ds As New DataSet()


        DataGridView1.Refresh()
        DataGridView1.AllowUserToResizeColumns = False
        DataGridView1.GridColor = Color.Black

        da.Fill(ds, "tbAdress")
        DataGridView1.DataSource = ds
        DataGridView1.DataMember = "tbAdress"


    End Sub


    Private Sub btnMeldescheinDrucken_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnMeldescheinDrucken.Click

        StartDruckenMeldeliste()

        MeldescheinUpdaten()

        Me.Close()

    End Sub


    Public Sub StartDruckenMeldeliste()

        If DruckerLesen() <> "no" Then
            PrintDocument3.DefaultPageSettings.PrinterSettings.PrinterName = DruckerLesen()
        Else
            Exit Sub
        End If

        Try
            PrintDocument3.Print()
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


    Private Sub PrintDocument3_PrintPage(ByVal sender As System.Object, ByVal e As System.Drawing.Printing.PrintPageEventArgs) Handles PrintDocument3.PrintPage

        Dim bm As New Bitmap(Me.DataGridView1.Width, Me.DataGridView1.Height)

        Dim DokumentTitel As String = "Meldeliste"

        DataGridView1.DrawToBitmap(bm, New Rectangle(0, 0, Me.DataGridView1.Width, Me.DataGridView1.Height))

        Dim g As Graphics = e.Graphics

        g.DrawString(DokumentTitel.ToString, New Font("Arial", 6, FontStyle.Bold, GraphicsUnit.Millimeter), Brushes.Black, 40, 10)
        g.DrawString("Gedruckt am:", New Font("Arial", 3, FontStyle.Bold, GraphicsUnit.Millimeter), Brushes.Black, 200, 10)
        g.DrawString(Now(), New Font("Arial", 3, FontStyle.Bold, GraphicsUnit.Millimeter), Brushes.Black, 280, 10)

        e.Graphics.DrawImage(bm, 0, 70)

    End Sub


    Private Sub MeldescheinUpdaten()

        Dim D As Date

        D = Now()

        Dim cmd As New OleDbCommand("UPDATE tbAdress SET MeldescheinPrint = " & True & ", MeldescheinPrintDatum = '" & D & "' WHERE MeldescheinPrint = false ", conn)

        Dim da As New OleDbDataAdapter(cmd)
        Dim ds As New DataSet("tbAdress")

        Try
            da.Fill(ds, "tbAdress")
        Catch ex As Exception
            MessageBox.Show(ex.Message)
        End Try

    End Sub



End Class