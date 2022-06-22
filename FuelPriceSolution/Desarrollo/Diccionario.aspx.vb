Imports FuelPrice.Business
Public Class Diccionario
    Inherits System.Web.UI.Page
    Private Datos As New FMBTraductor("Diccionario")
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            CargaFichas()
            CargaDiccionario()

        End If
    End Sub

    Private Sub CargaFichas()
        Dim dt As DataTable = Datos.getFichas()
        cmbDocumentos.DataSource = dt
        cmbDocumentos.DataTextField = dt.Columns("FCH_NOMBRE_FICHA").ToString()
        cmbDocumentos.DataValueField = dt.Columns("FCH_CODIGO").ToString()
        cmbDocumentos.DataBind()
    End Sub
    Private Sub CargaDiccionario()
        Dim dvDiccionario As New DataView(
        Datos.getDiccionarioParaFicha(cmbDocumentos.SelectedItem.Value))

        rptFraces.DataSource = dvDiccionario
        rptFraces.DataBind()
        CargaDiccionarioEnFicha()
    End Sub
    Private Sub CargaDiccionarioEnFicha()
        Dim dvDiccionario As New DataView(
        Datos.getDiccionarioDeFicha(cmbDocumentos.SelectedItem.Value))
        rptEnFicha.DataSource = dvDiccionario
        rptEnFicha.DataBind()
        txtFichaSeleccionada.Text = cmbDocumentos.SelectedItem.Text + ".aspx"
    End Sub

    Protected Sub btnBusca_Click(sender As Object, e As EventArgs)
        Dim dvDiccionario As New DataView(
        Datos.getDiccionarioParaFicha(cmbDocumentos.SelectedItem.Value, txtLlave.Text, txtBase.Text))
        rptFraces.DataSource = dvDiccionario
        rptFraces.DataBind()
        CargaDiccionarioEnFicha()
    End Sub

    Protected Sub rptFraces_ItemCommand(source As Object, e As RepeaterCommandEventArgs)


        If e.CommandName = "agrega" Then
            Datos.AgregaLlaveFicha(cmbDocumentos.SelectedItem.Value, e.CommandArgument.ToString())
            CargaDiccionario()
        End If
    End Sub

    Protected Sub rptEnFicha_ItemCommand(source As Object, e As RepeaterCommandEventArgs)
        If e.CommandName = "quita" Then
            Datos.QuitaLlaveFicha(cmbDocumentos.SelectedItem.Value, e.CommandArgument.ToString())
            CargaDiccionario()
        End If
    End Sub

    Protected Sub btnAgregaFicha_Click(sender As Object, e As EventArgs)
        If txtNuevaFicha.Text.Trim <> "" Then
            Datos.putAgregaFicha(txtNuevaFicha.Text)
            CargaFichas()
            CargaDiccionario()
        End If
    End Sub
End Class