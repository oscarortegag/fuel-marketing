<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="PreciosReferenciaTerminal.aspx.vb" Inherits="FuelPriceSolution.PreciosReferenciaTerminal" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<asp:Content ID="Content2" ContentPlaceHolderID="ScriptPlaceHolder" runat="server">
    <script>
        <%-- Funcion para mantener las Tabs Activas desde Front --%>
        function openCity(evt, cityName) {
            var i, tabcontent, tablinks;
            tablinks = document.getElementsByClassName("head-tab");
            document.getElementById(cityName).classList.remove("active")
            for (i = 0; i < tablinks.length; i++) {
                tablinks[i].className = tablinks[i].className.replace(" active", "");
                tablinks[i].className = tablinks[i].className.replace(" show", "");
            }
            document.getElementById(cityName + '-head-tab').classList.add("active");
            document.getElementById(cityName + '-tab').classList.add("active");
            document.getElementById(cityName).classList.add("active");
            document.getElementById(cityName).classList.add("show");
        }
        function openModalMasiva(title) {
            $('#md-tit').html(title);
            $('#masiveModal').modal('show');
        }
        function closeModalMasiva() {
            $('#masiveModal').modal('hide');
        }
        function openModalMasiva2(title) {
            $('#md-tit2').html(title);
            $('#masiveModal2').modal('show');
        }
        function closeModalMasiva2() {
            $('#masiveModal2').modal('hide');
        }
        function closeLoader() {
            $("#preloader").hide();
            $('#savingDiv').hide();
        }
        $(window).bind('beforeunload', function () {
            closeLoader();
        });
        function UploadFile(fileUpload) {
            if (fileUpload.value != '') {
                document.getElementById("<%=btnImport.ClientID %>").click();
            }
        }
        function setTranslation(ControlId, PropertyToChange, TranslationContent) {
            $("#" + ControlId).prop(PropertyToChange, TranslationContent);
        }
    </script>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="MainContent" runat="server">
    <div class="page-title">
		<div class="title_left">
			<h3><% =strTitulo %></h3>
		</div>
	</div>
    <div class="clearfix"></div>
	<div class="row">
		<div class="col-md-12 col-sm-12 ">
			<div class="x_panel">
				<div class="x_title">
					<div class="clearfix"></div>
				</div>
				<div class="x_content">
                    <ul class="nav nav-tabs bar_tabs" id="myTab" role="tablist">
                        <li class="nav-item head-tab <% =chkRegistro %>" id="registro-head-tab">
                        <a class="nav-link head-tab <% =chkRegistro %>" id="registro-tab" onclick="openCity(event, 'registro')" data-toggle="tab" href="#registro" role="tab" aria-controls="registro" aria-selected="<% =ariaRegistro %>"><% =strRegistroPrecios %></a>
                        </li>
                        <li class="nav-item head-tab <% =chkImportacion %>" id="import-head-tab" >
                      <a class="nav-link head-tab  <% =chkImportacion %>"  id="import-tab"    onclick="openCity(event, 'import')" data-toogle="tab" href="#import" role="tab" aria-controls="import" aria-selected="<% =ariaImportacion %>">Registro de precios PDF</a>
                        </li>
                            
                        <li class="nav-item head-tab <% =chkHistorico %>" id="definicion-head-tab">
                        <a class="nav-link head-tab <% =chkHistorico %>" id="definicion-tab" onclick="openCity(event, 'definicion')" data-toggle="tab" href="#definicion" role="tab" aria-controls="definicion" aria-selected="<% =ariaDefinicion %>"><% =strInformeHistorico %></a>
                        </li>
                    </ul>
                    <div class="tab-content" id="myTabContent">
                        <%-- Tab Registro --%>
                        <div class="tab-pane fade head-tab <% =stlRegistro %>" id="registro" role="tabpanel" aria-labelledby="registrotab-tab">
                            <asp:UpdatePanel runat="server" ID="panelRegistro">
                                <ContentTemplate>
                                    <asp:Panel runat="server" ID="pnlReg">
                                        <div class="row">&nbsp</div>

                                        <div class="row">
                                            <div class="col-4">
                                                <label><% =strTerminalSuministro %></label>
                                                <asp:DropDownList runat="server" ID="cmbTerminal" CssClass="form-control" AutoPostBack="true" OnSelectedIndexChanged="cmbTerminal_SelectedIndexChanged"></asp:DropDownList>
                                            </div>
                                            <div class="col-4">
                                                <label><% =strProductos %></label>
                                                <telerik:radcombobox runat="server" ID="cmbProductos" CheckBoxes="true" EnableCheckAllItemsCheckBox="true" 
                                                    style="width:100%;"
                                                    AllowCustomText="true" ForeColor="black"
                                                    EnableAriaSupport="True" Skin="MetroTouch" ></telerik:radcombobox>
                                            </div>
                                            <div class="col-4 text-right">
                                                <br />
                                                <asp:LinkButton runat="server" ID="btnImport" CssClass="btn btn-success" OnClick="btnImport_Click" style="display: none;" title="Importar archivo"><i class="fa fa-plus"></i></asp:LinkButton>
<%--                                            </div>
                                            <div class="col-2 text-right">--%>
                                                <label>
                                                    <span id="spanUpload" class="btn btn-primary"><i class="fa fa-upload"></i></span>
                                                    <asp:FileUpload runat="server" ID="fileXls" style="display: none" onchange="UploadFile(this);"/>
                                                </label>
                                                &nbsp;&nbsp;
                                                <asp:LinkButton runat="server" ID="btnCargaReg" CssClass="btn btn-primary" title="Cargar" OnClick="btnCargaReg_Click"><i class="fa fa-plus"></i></asp:LinkButton>
                                                &nbsp;&nbsp;
                                                <asp:LinkButton runat="server" ID="btnSave" CssClass="btn btn-primary" OnClick="btnSave_Click"><i class="fa fa-save"></i></asp:LinkButton>
                                            </div>
                                        </div>

                                        <div class="row">&nbsp;</div>

                                        <div class="row" style="overflow-x: auto;">
							                <%--<div id="divResumen" runat="server">--%>
								                <% =lblResumen %>   
								                <div class="divGrid">
									                <div class="grid">
										                <asp:GridView ID="grdResumen" 
												            runat="server"
													        CssClass="datatable"
													        CellPadding="0" 
													        CellSpacing="0"
													        GridLines="None" 
                                                            OnSelectedIndexChanging="grdResumen_SelectedIndexChanging"
													        AutoGenerateColumns="False" >

											                <RowStyle CssClass="even"/>
											                <HeaderStyle CssClass="header" />
											                <AlternatingRowStyle CssClass="odd"/>

											                <Columns>
                                                                <asp:CommandField ShowSelectButton="true" ButtonType="Link" SelectText="<span class='fa fa-pencil' style='font-size: 24px;' title='Modificar'></span>" ControlStyle-ForeColor="#5f97b3" />
												                <asp:BoundField HeaderText="ID Prod" DataField="IdProd" HeaderStyle-CssClass="hideGridColumn" ItemStyle-CssClass="hideGridColumn" />
                                                                <asp:BoundField HeaderText="Producto" DataField="Producto"/>
                                                                <asp:BoundField HeaderText="Precio" DataField="Precio"/>
                                                                <asp:BoundField HeaderText="Fecha" DataField="Fecha" dataformatstring="{0:dd/MM/yyyy}"/>
											                </Columns>

											                <HeaderStyle BackColor="#5f97b3" Font-Bold="true" ForeColor="White" />
										                </asp:GridView>

									                </div>
								                </div>
							                <%--</div> --%>
						                </div>     

                                        <div class="row">&nbsp;</div>

<%--                                        <div class="row">
                                            <div class="col-10">&nbsp;</div>
                                            <div class="col-2">
                                                <asp:LinkButton runat="server" ID="btnSave" CssClass="btn btn-primary form-control" OnClick="btnSave_Click">Guardar</asp:LinkButton>
                                            </div>
                                        </div>--%>

                                        <%--Modal--%>
                                        <div class="modal fade bs-example-modal-md" tabindex="-1" role="dialog" aria-hidden="true" id="masiveModal">
                                            <div class="modal-dialog modal-md">
                                                <div class="modal-content">
                                                    <div class="modal-header">
                                                        <h4 class="modal-title" id="md-tit"></h4>
                                                        <button type="button" class="close" data-dismiss="modal"><span aria-hidden="true">×</span></button>
                                                    </div>
                                                    <div class="modal-body" style="max-height: 350px; overflow-y:auto;">
                                                        <div class="row">
                                                            <asp:Label runat="server" ID="lblProducto" Visible="false"></asp:Label>
                                                            <div class="col-6"><% =strPrecio %></div>
                                                            <div class="col-6"><asp:TextBox runat="server" ID="txtPrecio" CssClass="form-control"></asp:TextBox></div>
                                                        </div>
                                                        <div class="row">&nbsp;</div>
                                                        <div class="row">
                                                            <div class="col-6"><% =strFecha %></div>
                                                            <div class="col-6"><asp:TextBox runat="server" ID="txtFechaPrecio" CssClass="form-control" type="date"></asp:TextBox></div>
                                                        </div>
                                                    </div>
                                                    <div class="modal-footer">
                                                        <%--<button type="button" class="btn btn-secondary" data-dismiss="modal">Cerrar</button>--%>
                                                        <asp:LinkButton CssClass="btn btn-secondary" runat="server" ID="btnCerrar" OnClick="btnCerrar_Click"></asp:LinkButton>
                                                        &nbsp;&nbsp;
                                                        <asp:LinkButton CssClass="btn btn-primary" runat="server" ID="btnAplica" OnClick="btnAplica_Click"></asp:LinkButton>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>    
                                        <%--Modal--%>

                                        <%--Modal--%>
                                        <div class="modal fade bs-example-modal-lg" tabindex="-1" role="dialog" aria-hidden="true" id="masiveModal2">
                                            <div class="modal-dialog modal-lg">
                                                <div class="modal-content">
                                                    <div class="modal-header">
                                                        <h4 class="modal-title" id="md-tit2"></h4>
                                                        <button type="button" class="close" data-dismiss="modal"><span aria-hidden="true">×</span></button>
                                                    </div>
                                                    <div class="modal-body" style="max-height: 350px; overflow-y:auto;">
                                                        <div class="row">
                                                            <div class="col-md-12 col-sm-12">
                                                                <div class="divGrid">
                                                                    <div class="grid">
                                                                        <asp:GridView ID="grdMasiva" 
                                                                            runat="server"
                                                                            CssClass="datatable"
                                                                            CellPadding="0" 
                                                                            CellSpacing="0"
                                                                            GridLines="None"
                                                                            AutoGenerateColumns="False" >

                                                                            <RowStyle CssClass="even"/>
                                                                            <HeaderStyle CssClass="header" />
                                                                            <AlternatingRowStyle CssClass="odd"/>

                                                                            <Columns>
                                                                                <asp:BoundField HeaderText="Resumen" DataField="Informe"/>   
                                                                            </Columns>

                                                                            <HeaderStyle BackColor="#5f97b3" Font-Bold="true" ForeColor="White" />
                                                                        </asp:GridView>

                                                                    </div>
                                                                </div>
                                                            </div>
                                                        </div>
                                                    </div>
                                                    <div class="modal-footer">
                                                        <button type="button" class="btn btn-secondary" data-dismiss="modal">Cerrar</button>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                        <%--Modal--%>
                                    </asp:Panel>
                                </ContentTemplate>
                                <Triggers>
                                    <asp:PostBackTrigger ControlID="cmbTerminal" />
                                    <asp:PostBackTrigger ControlID="cmbProductos" />
                                    <asp:PostBackTrigger ControlID="btnCargaReg" />
                                    <asp:PostBackTrigger ControlID="grdResumen" />
                                    <asp:PostBackTrigger ControlID="btnCerrar" />
                                    <asp:PostBackTrigger ControlID="btnAplica" />
                                    <asp:PostBackTrigger ControlID="btnSave" />
                                    <asp:PostBackTrigger ControlID="btnImport" />
                                </Triggers>
                            </asp:UpdatePanel>
                        </div>
                        <%-- Tab Registro --%>
                        

                        <%-- Tab Importacion  --%>
                             <div class="tab-pane fade head-tab <% =stlImportacion %>" id="import" role="tabpanel" aria-labelledby="import-tab">
                                 <asp:UpdatePanel ID="panelImportacion" runat ="server" >
                                     <ContentTemplate>
                                       <asp:Panel runat="server" ID="pnlImportacion">
                                           <div class="row">
                                               <div class="col-3">
                                                     <label>Fecha de aplicación</label>
                                                     <asp:TextBox runat="server" ID="Txt_Fecha_Aplicacion" type="date" CssClass="form-control Calenderr"></asp:TextBox>
                                                
                                                        </div>
                                               &nbsp;&nbsp;
                                               <br />
                                                <asp:LinkButton runat="server" ID="Btn_Import_PDF" CssClass="btn btn-success" OnClick="Btn_Import_PDF_Click" style="display: none;" title="Importar archivo"><i class="fa fa-upload"></i></asp:LinkButton>
<%--                                            </div>
                                            <div class="col-2 text-right">--%>
                                                <label>
                                                    <span id="spanUploadPDF" class="btn btn-primary"><i class="fa fa-upload"></i></span>
                                                    <asp:FileUpload runat="server" ID="FileUpload1" style="display: none" onchange="UploadFile(this);"/>
                                                </label>
                                           </div>
                                       </asp:Panel>    
                                     </ContentTemplate>
                                     <Triggers>
                                         <asp:PostBackTrigger ControlID="Btn_Import_PDF" />
                                     </Triggers>
                                 </asp:UpdatePanel>

                                 
                                 </div>
                        <%-- Tab Definición --%>
                        <div class="tab-pane fade head-tab <% =stlDefinicion %>" id="definicion" role="tabpanel" aria-labelledby="definicion-tab">
                            <asp:UpdatePanel ID="panelInforme" runat="server">
                                <ContentTemplate>
                                    <asp:Panel runat="server" ID="pnlInforme">
                                        <div class="row">                                            
                                            <div class="col-3">
                                                <telerik:radcombobox runat="server" ID="cmbTerminalInforme" CheckBoxes="true" EnableCheckAllItemsCheckBox="true" 
                                                    style="width:100%;"
                                                    AllowCustomText="true" ForeColor="black"
                                                    EnableAriaSupport="True" Skin="MetroTouch" ></telerik:radcombobox>
                                            </div>
                                            <div class="col-3">
                                                <telerik:radcombobox runat="server" ID="cmbProductoInforme" CheckBoxes="true" EnableCheckAllItemsCheckBox="true" 
                                                    style="width:100%;"
                                                    AllowCustomText="true" ForeColor="black"
                                                    EnableAriaSupport="True" Skin="MetroTouch" ></telerik:radcombobox>
                                            </div>
                                            <div class="col-3">
                                                <asp:TextBox runat="server" ID="txtDesde" type="date" CssClass="form-control"></asp:TextBox>
                                            </div>
                                            <div class="col-3">
                                                <asp:TextBox runat="server" ID="txtHasta" type="date" CssClass="form-control"></asp:TextBox>
                                            </div>
                                        </div>
                                        <div class="row">&nbsp;</div>
                                        <div class="row">
                                            <div class="col-3">
                                                <asp:LinkButton runat="server" ID="btnExcel" CssClass="btn btn-primary" OnClick="btnExcel_Click" OnClientClick="return closeLoader()" Visible="false"><i class="fa fa-file-excel-o text-success"></i></asp:LinkButton>
                                                &nbsp;&nbsp;
                                                <%--<asp:LinkButton runat="server" ID="btnPdf" CssClass="btn btn-white" OnClick="btnPdf_Click" OnClientClick="return closeLoader()" Visible="false"><img src="../Imagenes/pdf32.png" /></asp:LinkButton>--%>
                                            </div>
                                            <div class="col-6">&nbsp;</div>
                                            <div class="col-3 text-right">
                                                <asp:LinkButton runat="server" ID="btnLoadReport" CssClass="btn btn-primary" OnClick="btnLoadReport_Click"><i class="fa fa-search"></i></asp:LinkButton>
                                            </div>
                                        </div>

                                        <div class="row">&nbsp;</div>

                                        <div class="row" style="overflow-x: auto;">
								            <div class="divGrid">
									            <div class="grid">
                                                    <% =lblResumenReporte %>   
									            </div>
								            </div>
						                </div>     

                                        <div class="row">&nbsp;</div>

                                    </asp:Panel>
                                </ContentTemplate>
                                <Triggers>
                                    <asp:PostBackTrigger ControlID="btnExcel" />
                                    <%--<asp:PostBackTrigger ControlID="btnPdf" />--%>
                                    <asp:PostBackTrigger ControlID="btnLoadReport" />
                                </Triggers>
                            </asp:UpdatePanel>
                        </div>
                        <%-- Tab Definición --%>
                    </div>
                </div>
            </div>
		</div>				
    </div>   
    <div class="clearfix"></div>
</asp:Content>
