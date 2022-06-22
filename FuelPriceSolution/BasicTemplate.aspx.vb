Imports System.Data
Imports System.Drawing
Imports System.IO
Imports FuelPrice.Business
Imports SelectPdf
Public Class BasicTemplate
    Inherits System.Web.UI.Page
    Dim bl As FMBussinesLayer = New FMBussinesLayer()
    Dim ul As FMBussinesUtils = New FMBussinesUtils()

#Region "Variables para Tabs"
    Public chkRegistro As String
    Public chkHistorico As String
    Public chkDefinicion As String
    Public chkInforme As String

    Public stlRegistro As String
    Public stlDefinicion As String
    Public stlHistorico As String
    Public stlInforme As String

    Public ariaRegistro As String
    Public ariaDefinicion As String
    Public ariaHistorico As String
    Public ariaInforme As String
#End Region

    Private mCte As Integer
    Public lblResumen As String
    Public lblResumenReporte As String

#Region "Admin Tabs"
    Private Sub ChecaTab(ByVal tabName)
        chkRegistro = ""
        chkDefinicion = ""
        chkHistorico = ""
        chkInforme = ""

        stlRegistro = ""
        stlDefinicion = ""
        stlHistorico = ""
        stlInforme = ""

        ariaRegistro = "false"
        ariaDefinicion = "false"
        ariaHistorico = "false"
        ariaInforme = "false"

        Select Case tabName
            Case "Registro"
                chkRegistro = "active"
                stlRegistro = "active show"
                ariaRegistro = "true"
            Case "Definicion"
                chkHistorico = "active"
                stlDefinicion = "active show"
                ariaDefinicion = "true"
            Case "Historico"
                chkDefinicion = "active"
                stlHistorico = "active show"
                ariaHistorico = "true"
            Case "Informe"
                chkInforme = "active"
                stlInforme = "active show"
                ariaInforme = "true"
        End Select
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Master.Page.Title = "Fuel Market Solutions - Tabla Precios Referencia Terminal"
        mCte = Convert.ToInt32(Session("MaestroCliente"))
        If Not IsPostBack Then
            ChecaTab("Registro")
        End If
    End Sub

End Class