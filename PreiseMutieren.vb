Public Class PreiseMutieren

    Private Sub TbPreisBindingNavigatorSaveItem_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TbPreisBindingNavigatorSaveItem.Click
        Me.Validate()
        Me.TbPreisBindingSource.EndEdit()
        Me.TableAdapterManager.UpdateAll(Me.CampingDataSet)

    End Sub

    Private Sub PreiseMutieren_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'TODO: Diese Codezeile lädt Daten in die Tabelle "CampingDataSet.tbPreis". Sie können sie bei Bedarf verschieben oder entfernen.
        Me.TbPreisTableAdapter.Fill(Me.CampingDataSet.tbPreis)

    End Sub

    Private Sub btnLoeschen_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles btnLoeschen.Click

        Me.Close()

    End Sub
End Class