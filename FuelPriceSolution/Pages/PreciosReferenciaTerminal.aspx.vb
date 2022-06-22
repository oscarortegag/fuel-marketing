Imports System.Data
Imports System.Drawing
Imports System.IO
Imports FuelPrice.Business
Imports SelectPdf
Imports FMObjects

Public Class PreciosReferenciaTerminal
    Inherits System.Web.UI.Page
    Dim bl As FMBussinesLayer = New FMBussinesLayer()
    Dim ul As FMBussinesUtils = New FMBussinesUtils()

#Region "Variables para Tabs"
    Public chkRegistro As String
    Public chkHistorico As String
    Public chkDefinicion As String
    Public chkInforme As String
    Public chkImportacion As String

    Public stlRegistro As String
    Public stlDefinicion As String
    Public stlHistorico As String
    Public stlInforme As String
    Public stlImportacion As String

    Public ariaRegistro As String
    Public ariaDefinicion As String
    Public ariaHistorico As String
    Public ariaInforme As String
    Public ariaImportacion As String
#End Region

    Private mCte As Integer
    Private idUser As String
    Public lblResumen As String
    Public lblResumenReporte As String

    Private Cultura As String = "es-MX"
    Private tr As New FMBTraductor

#Region "Admin Tabs"
    Private Sub ChecaTab(ByVal tabName)
        chkRegistro = ""
        chkDefinicion = ""
        chkHistorico = ""
        chkInforme = ""
        chkImportacion = ""

        stlRegistro = ""
        stlDefinicion = ""
        stlHistorico = ""
        stlInforme = ""
        stlImportacion = ""

        ariaRegistro = "false"
        ariaDefinicion = "false"
        ariaHistorico = "false"
        ariaInforme = "false"
        ariaImportacion = "false"

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

            Case "Import"
                chkImportacion = "active"
                stlImportacion = "active show"
                ariaImportacion = "true"
        End Select
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        mCte = Convert.ToInt32(Session("MaestroCliente"))
        Cultura = Session("Cultura")
        idUser = Session("IdUsr")

        Traducciones()
        TraduceDocumento()

        If Not IsPostBack Then
            CargaTerminal()
            CargaProductos()
            ChecaTab("Registro")
        End If
    End Sub
    Protected Sub Traducciones()
        If IsNothing(ViewState("Traducciones")) Then
            Dim lstTranslate = tr.GetTranslations("PreciosReferenciaTerminal", Cultura)
            ViewState("Traducciones") = lstTranslate
        End If
    End Sub
    Protected Function Traduce(ByVal KeyName As String) As String
        Dim result As String = KeyName
        Dim trResult As List(Of Translated) = ViewState("Traducciones")
        For Each _trans In trResult
            If _trans.Key = KeyName Then
                If _trans.TranslatedString <> "" Then
                    result = _trans.TranslatedString
                Else
                    result = _trans.DefaultString
                End If
                Exit For
            End If
        Next
        Return result
    End Function

    Public strTitulo, strGuardar, strRegistroPrecios, strInformeHistorico, strTerminalSuministro, strProductos, strImportarArchivo, strPrecio, strFecha,
    strCerrar, strAplicar, StrDesde, StrHasta, strRegistroPreciosPDF As String
    Protected Sub TraduceDocumento()
        strTitulo = Traduce("_PRECREFTERMMIN_")

        strGuardar = Traduce("_GUARDAR_")
        strRegistroPrecios = Traduce("_REGISTROPRECIOS_")
        strInformeHistorico = Traduce("_INFORMEHISTORICO_")
        strTerminalSuministro = Traduce("_TERMINALSUMINISTRO_")
        strProductos = Traduce("_PRODUCTOS_")
        strImportarArchivo = Traduce("_IMPORTAR_")
        strPrecio = Traduce("_PRECIO_")
        strFecha = Traduce("_FECHA_")
        strCerrar = Traduce("_CERRAR_")
        strAplicar = Traduce("_APLICAR_")

        ScriptManager.RegisterStartupScript(panelRegistro, panelRegistro.GetType(), "Javascript",
                                            "setTranslation('spanUpload', 'title', '" + strImportarArchivo + "');" &
                                            "setTranslation('MainContent_btnCerrar', 'text', '" + strCerrar + "');" &
                                            "setTranslation('MainContent_btnAplica', 'text', '" + strAplicar + "');" &
                                            "setTranslation('MainContent_btnSave', 'title', '" + strGuardar + "');", True)

        grdResumen.Columns(2).HeaderText = Traduce("_PRODUCTO_")
        grdResumen.Columns(3).HeaderText = Traduce("_PRECIO_")
        grdResumen.Columns(4).HeaderText = Traduce("_FECHA_")

        Master.Page.Title = "Fuel Market Solutions - " & strTitulo
    End Sub
    Protected Sub CargaTerminal()
        Dim tblEmpresas As DataTable
        tblEmpresas = bl.GetTars()
        cmbTerminal.DataSource = tblEmpresas
        cmbTerminal.DataTextField = "TSU_DESCRIPCION"
        cmbTerminal.DataValueField = "TSU_CODIGO"
        cmbTerminal.DataBind()
        cmbTerminal.Items.Insert(0, New ListItem("-- " & Traduce("_SELECCIONETERMINA_") & " --", "0"))

        cmbTerminalInforme.DataSource = tblEmpresas.Copy
        cmbTerminalInforme.DataTextField = "TSU_DESCRIPCION"
        cmbTerminalInforme.DataValueField = "TSU_CODIGO"
        cmbTerminalInforme.DataBind()
        cmbTerminalInforme.EmptyMessage = Traduce("_SELECCIONETERMINA_")
        cmbTerminalInforme.Localization.AllItemsCheckedString = Traduce("_TODOSELECCIONADO_")
        cmbTerminalInforme.Localization.CheckAllString = Traduce("_SELECCIONARTODO_")
        cmbTerminalInforme.Localization.ItemsCheckedString = Traduce("_ELEMENTOSELECCIONADO_")
        cmbTerminalInforme.LoadingMessage = Traduce("_CARGANDO_")
    End Sub
    Protected Sub CargaProductos()
        Dim tblProductos As DataTable
        'tblProductos = bl.GetProductosVtas()

        tblProductos = bl.GetProductosCteBio(mCte)

        cmbProductos.DataSource = tblProductos
        cmbProductos.DataTextField = "CON_DESCRIPCION"
        cmbProductos.DataValueField = "CON_CODIGO"
        cmbProductos.DataBind()
        cmbProductos.EmptyMessage = Traduce("_SELECCIONEPRODUCTO_")
        cmbProductos.Localization.AllItemsCheckedString = Traduce("_TODOSELECCIONADO_")
        cmbProductos.Localization.CheckAllString = Traduce("_SELECCIONARTODO_")
        cmbProductos.Localization.ItemsCheckedString = Traduce("_ELEMENTOSELECCIONADO_")
        cmbProductos.LoadingMessage = Traduce("_CARGANDO_")

        cmbProductoInforme.DataSource = tblProductos.Copy
        cmbProductoInforme.DataTextField = "CON_DESCRIPCION"
        cmbProductoInforme.DataValueField = "CON_CODIGO"
        cmbProductoInforme.DataBind()
        cmbProductoInforme.EmptyMessage = Traduce("_SELECCIONEPRODUCTO_")
        cmbProductoInforme.Localization.AllItemsCheckedString = Traduce("_TODOSELECCIONADO_")
        cmbProductoInforme.Localization.CheckAllString = Traduce("_SELECCIONARTODO_")
        cmbProductoInforme.Localization.ItemsCheckedString = Traduce("_ELEMENTOSELECCIONADO_")
        cmbProductoInforme.LoadingMessage = Traduce("_CARGANDO_")
    End Sub
    Private Function TblTemp() As DataTable
        Dim _tbl As New DataTable
        _tbl.Columns.Add("IdProd")
        _tbl.Columns.Add("Producto")
        _tbl.Columns.Add("Precio")
        _tbl.Columns.Add("Fecha", GetType(DateTime))
        Return _tbl
    End Function
    Protected Sub cmbTerminal_SelectedIndexChanged(sender As Object, e As EventArgs)
        grdResumen.DataSource = New DataTable()
        grdResumen.DataBind()
        ChecaTab("Registro")
    End Sub

    Protected Sub Btn_Import_PDF_Click(sender As Object, e As EventArgs) ' AQUI ES DONDE EMPEZARAS A DESARROLLAR TU PROCESO


        ChecaTab("Import")
    End Sub

    Protected Sub btnCargaReg_Click(sender As Object, e As EventArgs)
        If cmbTerminal.SelectedIndex > 0 And cmbProductos.CheckedItems.Count > 0 Then
            Dim tblPrecios = TblTemp()
            For i = 0 To cmbProductos.CheckedItems.Count - 1
                tblPrecios.Rows.Add(cmbProductos.CheckedItems(i).Value, cmbProductos.CheckedItems(i).Text, 0, DateTime.Now.ToString("yyyy-MM-dd"))
            Next
            ViewState("tblPRecios") = tblPrecios
            grdResumen.DataSource = tblPrecios
            grdResumen.DataBind()
        End If
        ChecaTab("Registro")
    End Sub
    Protected Sub grdResumen_SelectedIndexChanging(sender As Object, e As GridViewSelectEventArgs)
        Dim idProd As String = grdResumen.Rows(e.NewSelectedIndex).Cells(1).Text
        Dim nomProd As String = grdResumen.Rows(e.NewSelectedIndex).Cells(2).Text

        lblProducto.Text = idProd
        Dim tblPRecios As DataTable = ViewState("tblPRecios")

        For index = 0 To tblPRecios.Rows.Count - 1

            If tblPRecios.Rows(index)("IdProd").ToString = idProd Then
                txtPrecio.Text = tblPRecios.Rows(index)("Precio")
                txtFechaPrecio.Text = Convert.ToDateTime(tblPRecios.Rows(index)("Fecha")).ToString("yyyy-MM-dd")
                tblPRecios.AcceptChanges()

                grdResumen.DataSource = tblPRecios
                grdResumen.DataBind()

                Exit For
            End If

        Next

        nomProd = "Precio para <strong>" & nomProd & "</strong>"

        ChecaTab("Registro")
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "Pop", "openModalMasiva('" + nomProd + "');", True)
    End Sub
    Protected Sub btnAplica_Click(sender As Object, e As EventArgs)
        If txtPrecio.Text <> "" And txtFechaPrecio.Text <> "" Then
            Dim tblPRecios As DataTable = ViewState("tblPRecios")

            For index = 0 To tblPRecios.Rows.Count - 1

                If tblPRecios.Rows(index)("IdProd").ToString = lblProducto.Text Then
                    tblPRecios.Rows(index)("Precio") = txtPrecio.Text
                    tblPRecios.Rows(index)("Fecha") = txtFechaPrecio.Text
                    tblPRecios.AcceptChanges()

                    grdResumen.DataSource = tblPRecios
                    grdResumen.DataBind()

                    Exit For
                End If

            Next
        End If

        ChecaTab("Registro")
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "Pop", "closeModalMasiva();", True)
    End Sub
    Protected Sub btnCerrar_Click(sender As Object, e As EventArgs)
        txtPrecio.Text = ""
        txtFechaPrecio.Text = ""
        lblProducto.Text = ""
        ChecaTab("Registro")
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "Pop", "closeModalMasiva();", True)
    End Sub
    Protected Sub btnSave_Click(sender As Object, e As EventArgs)
        Dim tblPRecios As DataTable = ViewState("tblPRecios")

        If tblPRecios IsNot Nothing Then
            If tblPRecios.Rows.Count > 0 Then
                Dim _inserta As Boolean = bl.SetPreciosReferencia(tblPRecios, cmbTerminal.SelectedValue)

                If _inserta Then
                    lblResumen = ""
                    grdResumen.DataSource = New DataTable
                    grdResumen.DataBind()
                Else
                    lblResumen = "<div class=""text-danger"">" & Traduce("_ERROR_") & "</div>"
                End If
            End If
        End If

        ChecaTab("Registro")
    End Sub
    Protected Sub btnLoadReport_Click(sender As Object, e As EventArgs)
        If cmbTerminalInforme.CheckedItems.Count > 0 And cmbProductoInforme.CheckedItems.Count > 0 And txtDesde.Text <> "" And txtHasta.Text <> "" Then
            Dim lstTerminal As String = ""
            Dim lstProductos As String = ""

            For index = 0 To cmbTerminalInforme.CheckedItems.Count - 1
                lstTerminal = lstTerminal & "'" & cmbTerminalInforme.CheckedItems(index).Value & "',"
            Next
            lstTerminal = Mid(lstTerminal, 1, Len(lstTerminal) - 1)

            For index = 0 To cmbProductoInforme.CheckedItems.Count - 1
                lstProductos = lstProductos & "'" & cmbProductoInforme.CheckedItems(index).Value & "',"
            Next
            lstProductos = Mid(lstProductos, 1, Len(lstProductos) - 1)

            ViewState("lstProductos") = lstProductos

            Dim tblReporte = bl.GetReportePrecios(lstTerminal, lstProductos, txtDesde.Text, txtHasta.Text, mCte, 0)
            If tblReporte.Rows.Count > 0 Then
                ViewState("tblReporte") = tblReporte
                lblResumenReporte = HtmlInforme(tblReporte, lstProductos)
                btnExcel.Visible = True
            Else
                lblResumenReporte = "<div class=""text-danger"">No existen datos para mostrar</div>"
                btnExcel.Visible = False
            End If

        End If
        ChecaTab("Definicion")
    End Sub
    Protected Sub btnExcel_Click(sender As Object, e As EventArgs)
        Dim tblReporte As DataTable = ViewState("tblReporte")
        Dim lstProductos As String = ViewState("lstProductos")

        Dim strReporte As String = "<table><tr><td colspan='5'><strong>Reporte de Precios de Referencia por Terminal</strong></td></tr><tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr></table>"

        strReporte &= "<br /><br />"

        strReporte &= "<div class='row' style='overflow-x: auto;'>" & vbNewLine &
                         "    <div Class='divGrid'>" & vbNewLine &
                         "        <div Class='grid'>" & vbNewLine
        strReporte &= HtmlInforme(tblReporte, lstProductos) & vbNewLine
        strReporte &= "      </div>" & vbNewLine &
                     "    </div>" & vbNewLine &
                     "</div>" & vbNewLine

        ViewState("lstProductos") = lstProductos
        ViewState("tblReporte") = tblReporte

        Response.ClearContent()
        Response.Buffer = True
        Response.AddHeader("content-disposition", String.Format("attachment; filename={0}", "Precios_Referencia_Terminal.xls"))
        Response.ContentEncoding = Encoding.UTF8
        Response.ContentType = "application/ms-excel"
        Response.Write(strReporte)
        Response.End()

        ChecaTab("Definicion")
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "none", "closeLoader();", False)
    End Sub
    Protected Sub btnPdf_Click(sender As Object, e As EventArgs)
        Dim tblReporte As DataTable = ViewState("tblReporte")
        Dim lstProductos As String = ViewState("lstProductos")

        Dim strReporte As String = "<table><tr><td colspan='5'><strong>Reporte de Precios de Referencia por Terminal</strong></td></tr><tr><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td><td>&nbsp;</td></tr></table>"

        strReporte &= "<br /><br />"

        strReporte &= "<div class='row' style='overflow-x: auto;'>" & vbNewLine &
                         "    <div Class='divGrid'>" & vbNewLine &
                         "        <div Class='grid'>" & vbNewLine
        strReporte &= HtmlInforme(tblReporte, lstProductos) & vbNewLine
        strReporte &= "      </div>" & vbNewLine &
                     "    </div>" & vbNewLine &
                     "</div>" & vbNewLine

        Dim baseUrl As String = ""

        Dim pdf_page_size As String = "Legal"
        Dim pageSize As PdfPageSize = DirectCast([Enum].Parse(GetType(PdfPageSize), pdf_page_size, True), PdfPageSize)

        Dim pdf_orientation As String = "Landscape"
        Dim pdfOrientation As PdfPageOrientation = DirectCast(
                    [Enum].Parse(GetType(PdfPageOrientation),
                    pdf_orientation, True), PdfPageOrientation)

        Dim webPageWidth As Integer = 1024

        Dim webPageHeight As Integer = 0

        ' instantiate a html to pdf converter object
        Dim converter As New HtmlToPdf()

        ' set converter options
        converter.Options.PdfPageSize = pageSize
        converter.Options.PdfPageOrientation = pdfOrientation
        converter.Options.WebPageWidth = webPageWidth
        converter.Options.WebPageHeight = webPageHeight

        ' create a new pdf document converting an url
        Dim doc As PdfDocument = converter.ConvertHtmlString(strReporte, baseUrl)

        ' save pdf document
        doc.Save(Response, False, "Precios_Referencia_Terminal.pdf")

        ' close pdf document
        doc.Close()

        ViewState("lstProductos") = lstProductos
        ViewState("tblReporte") = tblReporte
        ChecaTab("Definicion")
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "none", "closeLoader();", False)
    End Sub
    Protected Function HtmlInforme(ByVal tblReporte As DataTable, ByVal lstProductos As String) As String
        Dim strH As New StringBuilder

        Dim prod() As String = lstProductos.Split(",")

        Dim terList As New List(Of String)
        terList = (From row In tblReporte.AsEnumerable()
                   Select row.Field(Of String)("TSU_DESCRIPCION")).Distinct().ToList()

        terList.Sort()

        Dim dtDesde As DateTime = DateTime.Parse(txtDesde.Text)
        Dim dtHasta As DateTime = DateTime.Parse(txtHasta.Text)
        Dim totDays As Integer = DateDiff(DateInterval.Day, dtDesde, dtHasta)

        Dim tblRpt As New DataTable
        tblRpt.Columns.Add("Fecha")

        For index = totDays To 0 Step -1
            tblRpt.Rows.Add(dtHasta.ToString("dd/MM/yyyy"))
            dtHasta = dtHasta.AddDays(-1)
        Next

        strH.AppendLine("<table Class='datatable' cellspacing='0' cellpadding='0' style='border-collapse:collapse;'>")

        Dim colorTerminal As String = "84a5b5"

        strH.AppendLine("<tr class='header' style='color:White;background-color:#5F97B3;font-weight:bold;'>")
        strH.AppendLine("<td rowspan='2'><strong>" & Traduce("_FECHA_") & "</strong></td>")
        For Each ter As String In terList
            strH.AppendLine("<td colspan='" & prod.Count.ToString & "' style='background-color: #" & colorTerminal & " !important;'><strong>" & ter & "</strong></td>")
            If colorTerminal = "84a5b5" Then
                colorTerminal = "5F97B3"
            Else
                colorTerminal = "84a5b5"
            End If
        Next
        strH.AppendLine("</tr>")

        colorTerminal = "84a5b5"

        strH.AppendLine("<tr class='header' style='color:White;background-color:#5F97B3;font-weight:bold;'>")
        For Each ter As String In terList
            For index = 0 To prod.Count - 1
                For ind = 0 To cmbProductoInforme.Items.Count - 1
                    If ("'" & cmbProductoInforme.Items(ind).Value & "'") = prod(index) Then
                        strH.AppendLine("<td style='background-color: #" & colorTerminal & " !important;'><strong>" & cmbProductoInforme.Items(ind).Text & "</strong></td>")
                    End If
                Next
            Next
            If colorTerminal = "84a5b5" Then
                colorTerminal = "5F97B3"
            Else
                colorTerminal = "84a5b5"
            End If
        Next
        strH.AppendLine("</tr>")

        Dim trColor = "odd"

        For index = 0 To tblRpt.Rows.Count - 1
            Dim strR As New StringBuilder
            Dim dblRow As New Double

            If trColor = "odd" Then
                trColor = "even"
            Else
                trColor = "odd"
            End If
            strR.AppendLine("<tr class='" & trColor & "'>")
            strR.AppendLine("<td>" & tblRpt.Rows(index)(0).ToString & "</td>")
            For Each ter As String In terList
                For ind = 0 To prod.Count - 1
                    Dim rows() As DataRow = tblReporte.Select("CON_CODIGO=" & prod(ind) & " And TSU_DESCRIPCION='" & ter & "' And PRF_FECHA='" & tblRpt.Rows(index)(0).ToString & "'")
                    If rows.Count > 0 Then
                        strR.AppendLine("<td>" & rows(0)("PRF_PRECIO") & "</td>")
                        dblRow += Convert.ToDouble(rows(0)("PRF_PRECIO"))
                    Else
                        strR.AppendLine("<td>&nbsp;</td>")
                    End If
                Next
            Next
            strR.AppendLine("</tr>")

            If dblRow > 0 Then
                strH.Append(strR.ToString)
            End If
        Next

        strH.AppendLine("</table>")

        ViewState("tblReporte") = tblReporte


        Return strH.ToString
    End Function
    Protected Sub btnImport_Click(sender As Object, e As EventArgs)
        If fileXls.HasFile Then
            Dim fs As Stream = fileXls.PostedFile.InputStream
            Dim br As New BinaryReader(fs)
            Dim bytes As Byte() = br.ReadBytes(Convert.ToInt32(fs.Length))

            Dim tblResult As DataTable = ul.ImportPreciosReferencia(bytes)

            grdMasiva.DataSource = tblResult
            grdMasiva.DataBind()
        Else
            Dim result As DataTable = New DataTable
            result.Columns.Add("Informe")
            result.Rows.Add("Error al leer el archivo")

            grdMasiva.DataSource = result
            grdMasiva.DataBind()
        End If
        ScriptManager.RegisterStartupScript(Me, Me.GetType(), "Pop", "openModalMasiva2('Resultado de Carga masiva');", True)
        ChecaTab("Registro")
    End Sub
End Class