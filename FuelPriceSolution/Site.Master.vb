Imports System.Web
Imports FuelPrice.Business
Imports Newtonsoft.Json

Public Class SiteMaster
    Inherits MasterPage
    Public ServerPath As String
    Public CteMaestro As String, IdUsr As String, UserName As String
    Private LoginDts As New FMBussinesLogin()
    Public Cultura As String
    Dim Traductor As New FMBTraductor("Menu")
    Dim elTv As New TreeView()
    Dim ParametrosApi As New FMObjects.ApiParams
    Dim entidadNegocios As FMBBussinesLayerNegocios = New FMBBussinesLayerNegocios()
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load

        Try
            '''''''''''''''''FAVOR DE NO MOVER ESTA SECCION DE AQUI, no agregue más código arriba
            IdUsr = Session("IdUsr")
            If IdUsr = Nothing Then
                Response.Redirect("/Login.aspx", True)
                Return
            End If
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            Dim absoluteurl As String = HttpContext.Current.Request.Url.AbsoluteUri
            ServerPath = absoluteurl.Replace(HttpContext.Current.Request.Url.LocalPath, "")
            Me.Page.Title = "Fuel Market Solutions "
            CteMaestro = Session("MaestroCliente")

            UserName = Session("UserName")
            Cultura = Session("Cultura")

            '*****************************************************************************************************
            ' Nov/2021 - Agregó AhRS, Logo de cliente maestro
            Dim dtBanderaCL As DataTable
            dtBanderaCL = entidadNegocios.S_Clientes(Opcion:=3, IdUser:=IdUsr, MaeCliente:=CteMaestro)
            If dtBanderaCL IsNot Nothing And dtBanderaCL.Rows.Count > 0 Then
                imgLogoMCL.Visible = If(dtBanderaCL.Rows(0)("MCL_RUTALOGO").ToString() = "", False, True)
                imgLogoMCL.ImageUrl = dtBanderaCL.Rows(0)("MCL_RUTALOGO").ToString()
            Else
                imgLogoMCL.Visible = False
                imgLogoMCL.ImageUrl = ""
            End If
            '*****************************************************************************************************

            'System.Web.HttpContext.Current.Session("MaestroCliente") = 1
            ' cargarOpcionesUsuario(IdUsr, treeMenu)
            Dim dtMenu As New DataSet()
            dtMenu.Tables.Add(LoginDts.ArbolMenuPrfl(IdUsr))
            'con cada nodo escribimos un li o ul

            CrearNodosDelPadre(New TreeView(), 0, Nothing, dtMenu)
            ArbolBootStrap(elTv)
        Catch ex As Exception
            'Response.Redirect("YourPage.aspx")
            Server.Transfer("Login.aspx", False)
        End Try
    End Sub

    Private Sub ArbolBootStrap(tv As TreeView)

        Dim stControl As New StringBuilder()
        stControl.Append(" <div id=""sidebar-menu"" class=""main_menu_side hidden-print main_menu"">")
        stControl.Append(vbNewLine)
        stControl.Append("<div class='menu_section'>")
        stControl.Append(vbNewLine)
        stControl.Append("<ul class='nav side-menu'>")
        stControl.Append(vbNewLine)
        Dim NodosconPermiso = 0
        For Each nodo As TreeNode In tv.Nodes
            If nodo.ToolTip <> "NOV" Then

                NodosconPermiso = NodosconPermiso + 1
                If nodo.ChildNodes.Count > 0 Then
                    'TieneHijos

                    stControl.Append("<li><a><i class='" + nodo.ToolTip + "'></i><strong>" + nodo.Text + "</strong><span class='fa fa-chevron-down'></span></a>")
                    stControl.Append(vbNewLine)

                    stControl.Append("<ul class='nav child_menu'>")
                    stControl.Append(NodosHijosBootStrap(nodo.ChildNodes))
                    stControl.Append("</ul>")

                    stControl.Append(vbNewLine)
                    stControl.Append("</li>")
                    stControl.Append(vbNewLine)
                Else

                    ParametrosApi = JsonConvert.DeserializeObject(Of FMObjects.ApiParams)(nodo.ImageUrl)
                    If ParametrosApi IsNot Nothing Then
                        'hay parametros api y la api es true
                        If ParametrosApi.EsApi Then
                            ' Es Api se manda onClik
                            stControl.Append("<li><a href='#' onClick=""" + OnClicDeMenu(ParametrosApi, nodo.NavigateUrl, nodo.Value) + """> " + nodo.Text + "</a></li>")
                        Else
                            'si no es api es solo un link comun
                            stControl.Append("<li><a href='" + nodo.NavigateUrl + "'> " + nodo.Text + "</a></li>")
                        End If
                    Else
                        'si no hay parametros api es solo un link comun
                        stControl.Append("<li><a href='" + nodo.NavigateUrl + "'> " + nodo.Text + "</a></li>")
                    End If
                    stControl.Append(vbNewLine)
                End If
                stControl.Append(vbNewLine)
            End If
        Next
        stControl.Append("</ul>")
        stControl.Append(vbNewLine)
        stControl.Append("</div>")
        stControl.Append(vbNewLine)
        stControl.Append("</div>")
        stControl.Append(vbNewLine)
        If NodosconPermiso > 0 Then
            pnlMenu.Controls.Add(New LiteralControl(stControl.ToString()))
        End If
    End Sub
    Function OnClicDeMenu(ByVal LosParApi As FMObjects.ApiParams, ByVal laURL As String, ByVal IdMenu As String) As String

        If LosParApi IsNot Nothing Then
            If LosParApi.EsApi Then
                If LosParApi.Metodo.ToUpper() = "POST" Then
                    'Construit funcion OnClic par paasar parametros post
                    Dim Datas As String = ""
                    If LosParApi.Parametros.Count > 0 Then

                        For i = 0 To LosParApi.Parametros.Count - 1
                            Datas = Datas + LosParApi.Parametros.Keys(i).ToString() + ":""" + LosParApi.Parametros.Values(i).ToString() + """"
                            If i <> LosParApi.Parametros.Count - 1 Then
                                Datas = Datas + ","
                            End If
                        Next
                    End If
                    'Dim DestinoPost As String = "<a href='#' onClick='window.location.replace('/Pages/linker.aspx?MenuClicked=" + IdMenu + "');'> " + NodoText + "</a>"
                    Dim DestinoPost As String = "window.location.replace('/Pages/linker.aspx?MenuClicked=" + IdMenu + "');"

                    Return DestinoPost
                Else
                    ' si es GET
                    LosParApi.Parametros.Add("usn_id_usuario", Session("IdUsr"))
                    If LosParApi.Parametros.Count > 0 Then
                        laURL = laURL + "?"
                        For i = 0 To LosParApi.Parametros.Count - 1
                            laURL = laURL + LosParApi.Parametros.Keys(i).ToString() + "=" + LosParApi.Parametros.Values(i).ToString()
                            If i <> LosParApi.Parametros.Count - 1 Then
                                laURL = laURL + "&"
                            End If
                        Next
                    End If
                    ' Dim DestinoGet As String = "<a href='#' onClick='window.location.replace('/Pages/linker.aspx?MenuClicked=" + IdMenu + "');'> " + NodoText + "</a>"
                    Dim DestinoGet As String = "window.location.replace('" + laURL + "');"
                    Return DestinoGet
                End If
            Else
                ' Simplemente hacemos redirect al navigation.url
                Return "window.location.replace('" + laURL + "');"
            End If
        Else
            Return "window.location.replace('" + laURL + "');"
        End If

    End Function
    Private Function NodosHijosBootStrap(Nodos As TreeNodeCollection) As String

        Dim stControl As New StringBuilder()
        Dim NodosconPermiso = 0
        For Each nodoHijo As TreeNode In Nodos
            If nodoHijo.ToolTip <> "NOV" Then
                NodosconPermiso = NodosconPermiso + 1
                If nodoHijo.ChildNodes.Count > 0 Then

                    'TieneHijos
                    stControl.Append("<li><a>" + nodoHijo.Text + "<span class='fa fa-chevron-down'></span></a>")

                    stControl.Append(vbNewLine)

                    stControl.Append("<ul class='nav child_menu'>")
                    stControl.Append(NodosHijosBootStrap(nodoHijo.ChildNodes))
                    stControl.Append("</ul>")

                    stControl.Append(vbNewLine)
                    stControl.Append("</li>")
                    stControl.Append(vbNewLine)
                Else
                    ParametrosApi = JsonConvert.DeserializeObject(Of FMObjects.ApiParams)(nodoHijo.ImageUrl)
                    If ParametrosApi IsNot Nothing Then
                        'hay parametros api y la api es true
                        If ParametrosApi.EsApi Then
                            ' Es Api se manda onClik
                            stControl.Append("<li><a href='#' onClick=""" + OnClicDeMenu(ParametrosApi, nodoHijo.NavigateUrl, nodoHijo.Value) + """> " + nodoHijo.Text + "</a></li>")
                        Else
                            'si no es api es solo un link comun
                            stControl.Append("<li><a href='" + nodoHijo.NavigateUrl + "'> " + nodoHijo.Text + "</a></li>")
                        End If
                    Else
                        'si no hay parametros api es solo un link comun
                        stControl.Append("<li><a href='" + nodoHijo.NavigateUrl + "'> " + nodoHijo.Text + "</a></li>")
                    End If

                    ' stControl.Append("<li><a href='" + nodoHijo.NavigateUrl + "'>" + nodoHijo.Text + "</a></li>")
                    stControl.Append(vbNewLine)
                End If
            End If
        Next
        'stControl.Append("</ul>")
        stControl.Append(vbNewLine)
        Return stControl.ToString()
    End Function

    Private Sub CrearNodosDelPadre(ByVal elArbol As TreeView, ByVal indicePadre As Integer, ByVal nodePadre As TreeNode, ByVal dataSetArbol As DataSet)

        Dim dataViewHijos As DataView
        dataViewHijos = New DataView(dataSetArbol.Tables(0))
        dataViewHijos.RowFilter = dataSetArbol.Tables(0).Columns("IdMenuPadre").ColumnName & " = " & indicePadre.ToString()

        For Each dataRowCurrent As DataRowView In dataViewHijos

            Dim nuevoNodo As TreeNode = New TreeNode()
            nuevoNodo.Text = Traductor.Traduce(dataRowCurrent("Nombre").ToString(), Cultura)
            nuevoNodo.Value = dataRowCurrent("IdMenu").ToString().Trim()
            If dataRowCurrent("Estatus") Then

                nuevoNodo.NavigateUrl = dataRowCurrent("Url").ToString().Trim()
                nuevoNodo.ToolTip = dataRowCurrent("Desc_Class").ToString().Trim()
                nuevoNodo.ImageUrl = dataRowCurrent("Prms_API").ToString().Trim()
            Else
                If dataRowCurrent("IdMenuPadre") = 0 Then
                    nuevoNodo.ToolTip = dataRowCurrent("Desc_Class").ToString().Trim()
                Else
                    nuevoNodo.ToolTip = "NOV" 'NOV = No Visible
                End If
                nuevoNodo.NavigateUrl = "/#"

            End If



            If nodePadre Is Nothing Then
                elArbol.Nodes.Add(nuevoNodo)

            Else

                nodePadre.ChildNodes.Add(nuevoNodo)


            End If

            CrearNodosDelPadre(elArbol, Int32.Parse(dataRowCurrent("IdMenu").ToString()), nuevoNodo, dataSetArbol)
        Next


        If Not IsPostBack Then
            elArbol.CollapseAll()
        End If
        elTv = elArbol
    End Sub


    Private Shared _parametroDatatable As String
    Public Shared Property ParametroDatatable() As String
        Get
            _parametroDatatable = """lengthChange"": false, ""language"": { ""url"" : ""/Scripts/DataTableSpanish.json"" }"
            Return _parametroDatatable
        End Get
        Set(ByVal value As String)
            _parametroDatatable = value
        End Set
    End Property
    Public Shared Sub Notificacion(ByVal pagina As Page, ByVal mensaje As String,
                 Optional ByVal tipo As String = "info",
                 Optional ByVal retraso As Integer = 3000)
        Dim sb = New StringBuilder()
        sb.Append("<script type='text/javascript'>")
        sb.Append($"notificacionGenerica('{mensaje}', '{tipo}', {retraso})")
        sb.Append("</script>")
        ScriptManager.RegisterStartupScript(pagina, pagina.GetType(), Guid.NewGuid().ToString(), sb.ToString(), False)
    End Sub
    Shared Sub CarregarDataTable(ByVal pagina As Page, ByVal identificador As String)
        Dim sb = New StringBuilder()
        sb.Append("<script type='text/javascript'>")
        sb.Append($"$('{identificador}').DataTable({{{ParametroDatatable}}});")
        sb.Append("</script>")
        ScriptManager.RegisterStartupScript(pagina, pagina.GetType(), Guid.NewGuid().ToString(), sb.ToString(), False)
    End Sub
    ''' <summary>
    '''     Abrir modal
    ''' </summary>
    ''' <param name="pagina">Pagina fuentel</param>
    ''' <param name="identificador">identificador modal</param>
    Shared Sub ModalAbierto(ByVal pagina As Page, ByVal identificador As String)
        Dim sb = New StringBuilder()
        sb.Append("<script type='text/javascript'>")
        'sb.Append(" document.getElementById('" + identificador + "').modal('show');")
        sb.Append("$('" + identificador + "').modal('show');")
        sb.Append("</script>")
        ScriptManager.RegisterStartupScript(pagina, pagina.GetType(), Guid.NewGuid().ToString(), sb.ToString(), False)
    End Sub
    ''' <summary>
    '''     fechaSoltera
    ''' </summary>
    ''' 
    Shared Sub CarregarFechaSoltera(ByVal pagina As Page, ByVal identificador As String)
        Dim sb = New StringBuilder()
        sb.Append("<script type='text/javascript'>")
        sb.Append("  $('" & identificador & " ').daterangepicker({ singleDatePicker: true, singleClasses: 'picker_1', locale: { format: 'YYYY/MM/DD' }, startDate: moment(moment().toDate()).format('YYYY/MM/DD') });")
        sb.Append("</script>")
        ScriptManager.RegisterStartupScript(pagina, pagina.GetType(), Guid.NewGuid().ToString(), sb.ToString(), False)
    End Sub
    Shared Sub CarregarFechaSoltera(ByVal pagina As Page, ByVal identificador As String, data As String)
        Dim sb = New StringBuilder()
        sb.Append("<script type='text/javascript'>")
        sb.Append("  $('" & identificador & " ').daterangepicker({ singleDatePicker: true, singleClasses: 'picker_1', locale: { format: 'YYYY/MM/DD' }, startDate: '" & data & "' });")
        sb.Append("</script>")
        ScriptManager.RegisterStartupScript(pagina, pagina.GetType(), Guid.NewGuid().ToString(), sb.ToString(), False)
    End Sub
    Shared Sub CarregarFechaEditar(ByVal pagina As Page, ByVal identificador As String, data As String)
        Dim sb = New StringBuilder()
        sb.Append("<script type='text/javascript'>")
        sb.Append("document.getElementById('" & identificador & "').value = '" & data & "'")
        sb.Append("</script>")
        ScriptManager.RegisterStartupScript(pagina, pagina.GetType(), Guid.NewGuid().ToString(), sb.ToString(), False)
    End Sub
    ''' <summary>
    '''     Carrega fecha con solo exibición del año
    ''' </summary>
    ''' 
    Shared Sub CarregarFechaAno(ByVal pagina As Page, ByVal identificador As String)
        Dim sb = New StringBuilder()
        sb.Append("<script type='text/javascript'>")
        sb.Append("  $('" & identificador & " ').datetimepicker({ format: 'YYYY', viewMode: 'years' });")
        sb.Append("</script>")
        ScriptManager.RegisterClientScriptBlock(pagina, pagina.GetType(), Guid.NewGuid().ToString(), sb.ToString(), False)
    End Sub

    ''' <summary>
    '''     Carrega fecha con solo exibición del año
    ''' </summary>
    ''' 
    Shared Sub CarregarFechaMes(ByVal pagina As Page, ByVal identificador As String)
        Dim sb = New StringBuilder()
        sb.Append("<script type='text/javascript'>")
        sb.Append("  $('" & identificador & " ').datetimepicker({ format: 'MM', viewMode: 'months' });")
        sb.Append("</script>")
        ScriptManager.RegisterClientScriptBlock(pagina, pagina.GetType(), Guid.NewGuid().ToString(), sb.ToString(), False)
    End Sub
    ''' <summary>
    '''     Carregar Fecha ranger
    ''' </summary>
    ''' 
    Public Shared Sub CarregarFechaRanger(pagina As Page, identificador As String,
                                          Optional startDate As String = "`${moment().year()}/${moment().month() + 1}/01`",
                                          Optional endDate As String = "moment(moment().toDate()).format('YYYY/MM/DD')")
        Dim sb = New StringBuilder()
        sb.Append("<script type='text/javascript'>")
        sb.Append($"$('" & identificador & " ').daterangepicker({ locale: { format: 'YYYY/MM/DD' }, startDate: " + startDate + ", endDate: " + endDate + " });")
        sb.Append("</script>")
        ScriptManager.RegisterClientScriptBlock(pagina, pagina.GetType(), Guid.NewGuid().ToString(), sb.ToString(), False)
    End Sub

    ''' <summary>
    '''     Inhabilita la tab
    ''' </summary>
    Shared Sub InhabilitarTab(ByVal pagina As Page, ByVal identificador As String)
        Dim sb = New StringBuilder()
        sb.Append("<script type='text/javascript'>")
        sb.Append("$('" + identificador + "').addClass('disabled');")
        sb.Append("</script>")
        ScriptManager.RegisterStartupScript(pagina, pagina.GetType(), Guid.NewGuid().ToString(), sb.ToString(), False)
    End Sub
    Shared Sub CarregarSelect2(ByVal pagina As Page, ByVal identificador As String, ByVal placeholder As String)
        Dim sb = New StringBuilder()
        sb.Append("<script type='text/javascript'>")
        sb.Append("$('" + identificador + "').select2({ language: 'es', width: '100%', placeholder: '" + placeholder + "', minimumInputLength: 0, allowClear: false});")
        sb.Append("</script>")
        ScriptManager.RegisterClientScriptBlock(pagina, pagina.GetType(), Guid.NewGuid().ToString(), sb.ToString(), False)
    End Sub
    ''' <summary>
    '''     Cerrar modal
    ''' </summary>
    ''' <param name="pagina">Pagina fuente</param>
    ''' <param name="identificador">Identificador de lo Select2</param>
    Shared Sub LlenarDatosSelect2(pagina As Page, identificador As String, datos As String)
        Dim sb = New StringBuilder()
        sb.Append("<script type='text/javascript'>")
        sb.Append($"$('{identificador}').val([{datos}]).trigger('change');")
        sb.Append("</script>")
        ScriptManager.RegisterStartupScript(pagina, pagina.GetType(), Guid.NewGuid().ToString(), sb.ToString(), False)
    End Sub
    Shared Sub CerrarModal(ByVal pagina As Page, ByVal identificador As String)
        Dim sb = New StringBuilder()
        sb.Append("<script type='text/javascript'>")
        sb.Append($"$('{identificador}').modal('hide');")
        sb.Append("</script>")
        ScriptManager.RegisterStartupScript(pagina, pagina.GetType(), Guid.NewGuid().ToString(), sb.ToString(), False)
    End Sub

    ''' <summary>
    '''     Cerrar modal
    ''' </summary>
    ''' <param name="mesNumero">mes en número</param>
    Public Shared Function RetornaMesEscrito(ByVal mesNumero As String) As String
        Dim retorno = ""

        If mesNumero = "01" Then
            retorno = "ENERO"
        ElseIf mesNumero = "2" Then
            retorno = "FEBRERO"
        ElseIf mesNumero = "3" Then
            retorno = "MARZO"
        ElseIf mesNumero = "4" Then
            retorno = "ABRIL"
        ElseIf mesNumero = "5" Then
            retorno = "MAYO"
        ElseIf mesNumero = "6" Then
            retorno = "JUNIO"
        ElseIf mesNumero = "7" Then
            retorno = "JULIO"
        ElseIf mesNumero = "8" Then
            retorno = "AGOSTO"
        ElseIf mesNumero = "9" Then
            retorno = "SEPTIEMBRE"
        ElseIf mesNumero = "10" Then
            retorno = "OCTUBRE"
        ElseIf mesNumero = "11" Then
            retorno = "NOVIEMBRE"
        ElseIf mesNumero = "12" Then
            retorno = "DECIEMBRE"
        End If

        Return retorno
    End Function
End Class