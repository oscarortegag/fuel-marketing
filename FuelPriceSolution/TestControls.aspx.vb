Imports System.Data
Imports Telerik.Web

Public Class TestControls
    Inherits System.Web.UI.Page
    Dim algo = 0
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            LlenaCombo()
        End If
    End Sub

    Private Sub LlenaCombo()
        Dim _tbl As New DataTable()
        _tbl.Columns.Add("ID")
        _tbl.Columns.Add("Description")

        Dim _fech = New DateTime(2020, 12, 1)

        For x = 1 To 12
            _fech = _fech.AddMonths(1)
            Dim _row As DataRow = _tbl.NewRow
            _row("ID") = x
            _row("Description") = _fech.ToString("MMMM")
            _tbl.Rows.Add(_row)
        Next

        'Me.cmbMulti.DataSource = _tbl
        'Me.cmbMulti.DataTextField = "Description"
        'Me.cmbMulti.DataValueField = "ID"
        'Me.cmbMulti.DataBind()
        'Me.cmbMulti.Items.Insert(0, New ListItem("-- Seleccione Mes --", "0"))

        Me.radCombo1.DataSource = _tbl
        Me.radCombo1.DataTextField = "Description"
        Me.radCombo1.DataValueField = "ID"
        Me.radCombo1.DataBind()
        Me.radCombo1.Items.Insert(0, "-- Seleccione Mes --")



        'ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), "none", "setMulti();", True)
        'ScriptManager.RegisterStartupScript(Me, Me.GetType(), "Pop", "openModal();", True)
    End Sub

    Protected Sub btnSave_Click(sender As Object, e As EventArgs) Handles btnSave.Click
        algo = Val(Me.lblUno.Text)
        algo += 1
        Me.lblUno.Text = algo
    End Sub
End Class