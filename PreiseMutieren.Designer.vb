<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class PreiseMutieren
    Inherits System.Windows.Forms.Form

    'Das Formular überschreibt den Löschvorgang, um die Komponentenliste zu bereinigen.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Wird vom Windows Form-Designer benötigt.
    Private components As System.ComponentModel.IContainer

    'Hinweis: Die folgende Prozedur ist für den Windows Form-Designer erforderlich.
    'Das Bearbeiten ist mit dem Windows Form-Designer möglich.  
    'Das Bearbeiten mit dem Code-Editor ist nicht möglich.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(PreiseMutieren))
        Me.CampingDataSet = New CampingTriesen.campingDataSet()
        Me.TbPreisBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.TbPreisTableAdapter = New CampingTriesen.campingDataSetTableAdapters.tbPreisTableAdapter()
        Me.TableAdapterManager = New CampingTriesen.campingDataSetTableAdapters.TableAdapterManager()
        Me.TbPreisBindingNavigator = New System.Windows.Forms.BindingNavigator(Me.components)
        Me.BindingNavigatorMoveFirstItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMovePreviousItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorPositionItem = New System.Windows.Forms.ToolStripTextBox()
        Me.BindingNavigatorCountItem = New System.Windows.Forms.ToolStripLabel()
        Me.BindingNavigatorSeparator1 = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorMoveNextItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorMoveLastItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.BindingNavigatorAddNewItem = New System.Windows.Forms.ToolStripButton()
        Me.BindingNavigatorDeleteItem = New System.Windows.Forms.ToolStripButton()
        Me.TbPreisBindingNavigatorSaveItem = New System.Windows.Forms.ToolStripButton()
        Me.TbPreisDataGridView = New System.Windows.Forms.DataGridView()
        Me.DataGridViewTextBoxColumn1 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn2 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn3 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn4 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn5 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.DataGridViewTextBoxColumn6 = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.lblSuchTitel1 = New System.Windows.Forms.Label()
        Me.btnLoeschen = New System.Windows.Forms.Button()
        CType(Me.CampingDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TbPreisBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.TbPreisBindingNavigator, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.TbPreisBindingNavigator.SuspendLayout()
        CType(Me.TbPreisDataGridView, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'CampingDataSet
        '
        Me.CampingDataSet.DataSetName = "campingDataSet"
        Me.CampingDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'TbPreisBindingSource
        '
        Me.TbPreisBindingSource.DataMember = "tbPreis"
        Me.TbPreisBindingSource.DataSource = Me.CampingDataSet
        '
        'TbPreisTableAdapter
        '
        Me.TbPreisTableAdapter.ClearBeforeFill = True
        '
        'TableAdapterManager
        '
        Me.TableAdapterManager.BackupDataSetBeforeUpdate = False
        Me.TableAdapterManager.tbPreisTableAdapter = Me.TbPreisTableAdapter
        Me.TableAdapterManager.UpdateOrder = CampingTriesen.campingDataSetTableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete
        '
        'TbPreisBindingNavigator
        '
        Me.TbPreisBindingNavigator.AddNewItem = Me.BindingNavigatorAddNewItem
        Me.TbPreisBindingNavigator.BindingSource = Me.TbPreisBindingSource
        Me.TbPreisBindingNavigator.CountItem = Me.BindingNavigatorCountItem
        Me.TbPreisBindingNavigator.DeleteItem = Me.BindingNavigatorDeleteItem
        Me.TbPreisBindingNavigator.Dock = System.Windows.Forms.DockStyle.None
        Me.TbPreisBindingNavigator.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.BindingNavigatorMoveFirstItem, Me.BindingNavigatorMovePreviousItem, Me.BindingNavigatorSeparator, Me.BindingNavigatorPositionItem, Me.BindingNavigatorCountItem, Me.BindingNavigatorSeparator1, Me.BindingNavigatorMoveNextItem, Me.BindingNavigatorMoveLastItem, Me.BindingNavigatorSeparator2, Me.BindingNavigatorAddNewItem, Me.BindingNavigatorDeleteItem, Me.TbPreisBindingNavigatorSaveItem})
        Me.TbPreisBindingNavigator.Location = New System.Drawing.Point(534, 503)
        Me.TbPreisBindingNavigator.MoveFirstItem = Me.BindingNavigatorMoveFirstItem
        Me.TbPreisBindingNavigator.MoveLastItem = Me.BindingNavigatorMoveLastItem
        Me.TbPreisBindingNavigator.MoveNextItem = Me.BindingNavigatorMoveNextItem
        Me.TbPreisBindingNavigator.MovePreviousItem = Me.BindingNavigatorMovePreviousItem
        Me.TbPreisBindingNavigator.Name = "TbPreisBindingNavigator"
        Me.TbPreisBindingNavigator.PositionItem = Me.BindingNavigatorPositionItem
        Me.TbPreisBindingNavigator.Size = New System.Drawing.Size(287, 25)
        Me.TbPreisBindingNavigator.TabIndex = 0
        Me.TbPreisBindingNavigator.Text = "BindingNavigator1"
        '
        'BindingNavigatorMoveFirstItem
        '
        Me.BindingNavigatorMoveFirstItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveFirstItem.Image = CType(resources.GetObject("BindingNavigatorMoveFirstItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveFirstItem.Name = "BindingNavigatorMoveFirstItem"
        Me.BindingNavigatorMoveFirstItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveFirstItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMoveFirstItem.Text = "Erste verschieben"
        '
        'BindingNavigatorMovePreviousItem
        '
        Me.BindingNavigatorMovePreviousItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMovePreviousItem.Image = CType(resources.GetObject("BindingNavigatorMovePreviousItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMovePreviousItem.Name = "BindingNavigatorMovePreviousItem"
        Me.BindingNavigatorMovePreviousItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMovePreviousItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorMovePreviousItem.Text = "Vorherige verschieben"
        '
        'BindingNavigatorSeparator
        '
        Me.BindingNavigatorSeparator.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator.Size = New System.Drawing.Size(6, 25)
        '
        'BindingNavigatorPositionItem
        '
        Me.BindingNavigatorPositionItem.AccessibleName = "Position"
        Me.BindingNavigatorPositionItem.AutoSize = False
        Me.BindingNavigatorPositionItem.Name = "BindingNavigatorPositionItem"
        Me.BindingNavigatorPositionItem.Size = New System.Drawing.Size(50, 23)
        Me.BindingNavigatorPositionItem.Text = "0"
        Me.BindingNavigatorPositionItem.ToolTipText = "Aktuelle Position"
        '
        'BindingNavigatorCountItem
        '
        Me.BindingNavigatorCountItem.Name = "BindingNavigatorCountItem"
        Me.BindingNavigatorCountItem.Size = New System.Drawing.Size(44, 15)
        Me.BindingNavigatorCountItem.Text = "von {0}"
        Me.BindingNavigatorCountItem.ToolTipText = "Die Gesamtanzahl der Elemente."
        '
        'BindingNavigatorSeparator1
        '
        Me.BindingNavigatorSeparator1.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator1.Size = New System.Drawing.Size(6, 6)
        '
        'BindingNavigatorMoveNextItem
        '
        Me.BindingNavigatorMoveNextItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveNextItem.Image = CType(resources.GetObject("BindingNavigatorMoveNextItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveNextItem.Name = "BindingNavigatorMoveNextItem"
        Me.BindingNavigatorMoveNextItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveNextItem.Size = New System.Drawing.Size(23, 20)
        Me.BindingNavigatorMoveNextItem.Text = "Nächste verschieben"
        '
        'BindingNavigatorMoveLastItem
        '
        Me.BindingNavigatorMoveLastItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorMoveLastItem.Image = CType(resources.GetObject("BindingNavigatorMoveLastItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorMoveLastItem.Name = "BindingNavigatorMoveLastItem"
        Me.BindingNavigatorMoveLastItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorMoveLastItem.Size = New System.Drawing.Size(23, 20)
        Me.BindingNavigatorMoveLastItem.Text = "Letzte verschieben"
        '
        'BindingNavigatorSeparator2
        '
        Me.BindingNavigatorSeparator2.Name = "BindingNavigatorSeparator"
        Me.BindingNavigatorSeparator2.Size = New System.Drawing.Size(6, 6)
        '
        'BindingNavigatorAddNewItem
        '
        Me.BindingNavigatorAddNewItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorAddNewItem.Image = CType(resources.GetObject("BindingNavigatorAddNewItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorAddNewItem.Name = "BindingNavigatorAddNewItem"
        Me.BindingNavigatorAddNewItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorAddNewItem.Size = New System.Drawing.Size(23, 22)
        Me.BindingNavigatorAddNewItem.Text = "Neu hinzufügen"
        '
        'BindingNavigatorDeleteItem
        '
        Me.BindingNavigatorDeleteItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.BindingNavigatorDeleteItem.Image = CType(resources.GetObject("BindingNavigatorDeleteItem.Image"), System.Drawing.Image)
        Me.BindingNavigatorDeleteItem.Name = "BindingNavigatorDeleteItem"
        Me.BindingNavigatorDeleteItem.RightToLeftAutoMirrorImage = True
        Me.BindingNavigatorDeleteItem.Size = New System.Drawing.Size(23, 20)
        Me.BindingNavigatorDeleteItem.Text = "Löschen"
        '
        'TbPreisBindingNavigatorSaveItem
        '
        Me.TbPreisBindingNavigatorSaveItem.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image
        Me.TbPreisBindingNavigatorSaveItem.Image = CType(resources.GetObject("TbPreisBindingNavigatorSaveItem.Image"), System.Drawing.Image)
        Me.TbPreisBindingNavigatorSaveItem.Name = "TbPreisBindingNavigatorSaveItem"
        Me.TbPreisBindingNavigatorSaveItem.Size = New System.Drawing.Size(23, 23)
        Me.TbPreisBindingNavigatorSaveItem.Text = "Daten speichern"
        '
        'TbPreisDataGridView
        '
        Me.TbPreisDataGridView.AllowUserToAddRows = False
        Me.TbPreisDataGridView.AllowUserToDeleteRows = False
        Me.TbPreisDataGridView.AllowUserToResizeColumns = False
        Me.TbPreisDataGridView.AllowUserToResizeRows = False
        Me.TbPreisDataGridView.AutoGenerateColumns = False
        Me.TbPreisDataGridView.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.DisplayedCells
        Me.TbPreisDataGridView.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.TbPreisDataGridView.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.DataGridViewTextBoxColumn1, Me.DataGridViewTextBoxColumn2, Me.DataGridViewTextBoxColumn3, Me.DataGridViewTextBoxColumn4, Me.DataGridViewTextBoxColumn5, Me.DataGridViewTextBoxColumn6})
        Me.TbPreisDataGridView.DataSource = Me.TbPreisBindingSource
        Me.TbPreisDataGridView.Location = New System.Drawing.Point(40, 50)
        Me.TbPreisDataGridView.Name = "TbPreisDataGridView"
        Me.TbPreisDataGridView.Size = New System.Drawing.Size(812, 450)
        Me.TbPreisDataGridView.TabIndex = 1
        '
        'DataGridViewTextBoxColumn1
        '
        Me.DataGridViewTextBoxColumn1.DataPropertyName = "PreisId"
        Me.DataGridViewTextBoxColumn1.HeaderText = "PreisId"
        Me.DataGridViewTextBoxColumn1.Name = "DataGridViewTextBoxColumn1"
        Me.DataGridViewTextBoxColumn1.Width = 64
        '
        'DataGridViewTextBoxColumn2
        '
        Me.DataGridViewTextBoxColumn2.DataPropertyName = "BezeichnungD"
        Me.DataGridViewTextBoxColumn2.HeaderText = "BezeichnungD"
        Me.DataGridViewTextBoxColumn2.Name = "DataGridViewTextBoxColumn2"
        Me.DataGridViewTextBoxColumn2.Width = 102
        '
        'DataGridViewTextBoxColumn3
        '
        Me.DataGridViewTextBoxColumn3.DataPropertyName = "BezeichnungE"
        Me.DataGridViewTextBoxColumn3.HeaderText = "BezeichnungE"
        Me.DataGridViewTextBoxColumn3.Name = "DataGridViewTextBoxColumn3"
        Me.DataGridViewTextBoxColumn3.Width = 101
        '
        'DataGridViewTextBoxColumn4
        '
        Me.DataGridViewTextBoxColumn4.DataPropertyName = "BezeichnungF"
        Me.DataGridViewTextBoxColumn4.HeaderText = "BezeichnungF"
        Me.DataGridViewTextBoxColumn4.Name = "DataGridViewTextBoxColumn4"
        '
        'DataGridViewTextBoxColumn5
        '
        Me.DataGridViewTextBoxColumn5.DataPropertyName = "BezeichnungI"
        Me.DataGridViewTextBoxColumn5.HeaderText = "BezeichnungI"
        Me.DataGridViewTextBoxColumn5.Name = "DataGridViewTextBoxColumn5"
        Me.DataGridViewTextBoxColumn5.Width = 97
        '
        'DataGridViewTextBoxColumn6
        '
        Me.DataGridViewTextBoxColumn6.DataPropertyName = "Preis"
        Me.DataGridViewTextBoxColumn6.HeaderText = "Preis"
        Me.DataGridViewTextBoxColumn6.Name = "DataGridViewTextBoxColumn6"
        Me.DataGridViewTextBoxColumn6.Width = 55
        '
        'lblSuchTitel1
        '
        Me.lblSuchTitel1.AutoSize = True
        Me.lblSuchTitel1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblSuchTitel1.Location = New System.Drawing.Point(36, 9)
        Me.lblSuchTitel1.Name = "lblSuchTitel1"
        Me.lblSuchTitel1.Size = New System.Drawing.Size(157, 20)
        Me.lblSuchTitel1.TabIndex = 141
        Me.lblSuchTitel1.Text = "Preisliste mutieren"
        '
        'btnLoeschen
        '
        Me.btnLoeschen.Location = New System.Drawing.Point(777, 12)
        Me.btnLoeschen.Name = "btnLoeschen"
        Me.btnLoeschen.Size = New System.Drawing.Size(75, 23)
        Me.btnLoeschen.TabIndex = 142
        Me.btnLoeschen.Text = "Schliessen"
        Me.btnLoeschen.UseVisualStyleBackColor = True
        '
        'PreiseMutieren
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(891, 694)
        Me.Controls.Add(Me.btnLoeschen)
        Me.Controls.Add(Me.lblSuchTitel1)
        Me.Controls.Add(Me.TbPreisDataGridView)
        Me.Controls.Add(Me.TbPreisBindingNavigator)
        Me.Name = "PreiseMutieren"
        Me.Text = "PreiseMutieren"
        CType(Me.CampingDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TbPreisBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.TbPreisBindingNavigator, System.ComponentModel.ISupportInitialize).EndInit()
        Me.TbPreisBindingNavigator.ResumeLayout(False)
        Me.TbPreisBindingNavigator.PerformLayout()
        CType(Me.TbPreisDataGridView, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CampingDataSet As CampingTriesen.campingDataSet
    Friend WithEvents TbPreisBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents TbPreisTableAdapter As CampingTriesen.campingDataSetTableAdapters.tbPreisTableAdapter
    Friend WithEvents TableAdapterManager As CampingTriesen.campingDataSetTableAdapters.TableAdapterManager
    Friend WithEvents TbPreisBindingNavigator As System.Windows.Forms.BindingNavigator
    Friend WithEvents BindingNavigatorAddNewItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorCountItem As System.Windows.Forms.ToolStripLabel
    Friend WithEvents BindingNavigatorDeleteItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveFirstItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMovePreviousItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorPositionItem As System.Windows.Forms.ToolStripTextBox
    Friend WithEvents BindingNavigatorSeparator1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents BindingNavigatorMoveNextItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorMoveLastItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents BindingNavigatorSeparator2 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents TbPreisBindingNavigatorSaveItem As System.Windows.Forms.ToolStripButton
    Friend WithEvents TbPreisDataGridView As System.Windows.Forms.DataGridView
    Friend WithEvents DataGridViewTextBoxColumn1 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn2 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn3 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn4 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn5 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents DataGridViewTextBoxColumn6 As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents lblSuchTitel1 As System.Windows.Forms.Label
    Friend WithEvents btnLoeschen As System.Windows.Forms.Button
End Class
