<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="BasicTemplate.aspx.vb" Inherits="FuelPriceSolution.BasicTemplate" %>
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
    </script>

</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="MainContent" runat="server">

    <div class="page-title">
		<div class="title_left">
			<h3>Basic Template Title</h3>
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
                        <a class="nav-link head-tab <% =chkRegistro %>" id="registro-tab" onclick="openCity(event, 'registro')" data-toggle="tab" href="#registro" role="tab" aria-controls="registro" aria-selected="<% =ariaRegistro %>">Registro de Indirectos</a>
                        </li>
                        <li class="nav-item head-tab <% =chkHistorico %>" id="definicion-head-tab">
                        <a class="nav-link head-tab <% =chkHistorico %>" id="definicion-tab" onclick="openCity(event, 'definicion')" data-toggle="tab" href="#definicion" role="tab" aria-controls="definicion" aria-selected="<% =ariaDefinicion %>">Definicion de valores y vigencias</a>
                        </li>
                        <li class="nav-item head-tab <% =chkDefinicion %>" id="historico-head-tab">
                        <a class="nav-link head-tab <% =chkDefinicion %>" id="historico-tab" onclick="openCity(event, 'historico')" data-toggle="tab" href="#historico" role="tab" aria-controls="historico" aria-selected="<% =ariaHistorico %>">Histórico de costos indirectos</a>
                        </li>
                        <li class="nav-item head-tab <% =chkInforme %>" id="informe-head-tab">
                        <a class="nav-link head-tab <% =chkInforme %>" id="informe-tab" onclick="openCity(event, 'informe')" data-toggle="tab" href="#informe" role="tab" aria-controls="informe" aria-selected="<% =ariaInforme %>">Informe de costos indirectos</a>
                        </li>
                    </ul>
                    <div class="tab-content" id="myTabContent">
                        <%-- Tab Registro --%>
                        <div class="tab-pane fade head-tab <% =stlRegistro %>" id="registro" role="tabpanel" aria-labelledby="registrotab-tab">
                           
                        </div>
                        <%-- Tab Registro --%>
                            
                        <%-- Tab Definición --%>
                        <div class="tab-pane fade head-tab <% =stlDefinicion %>" id="definicion" role="tabpanel" aria-labelledby="definicion-tab">
                            
                        </div>
                        <%-- Tab Definición --%>
                            
                        <%-- Tab Histórico --%>
                        <div class="tab-pane fade head-tab <% =stlHistorico %>" id="historico" role="tabpanel" aria-labelledby="historico-tab">
                            
                        </div>
                        <%-- Tab Histórico --%>

                        <%-- Tab Informe --%>
                        <div class="tab-pane fade head-tab <% =stlInforme %>" id="informe" role="tabpanel" aria-labelledby="informe-tab">
                            
                        </div>
                        <%-- Tab Informe --%>
                    </div>
                </div>
            </div>
		</div>				
    </div>   
    <div class="clearfix"></div>

</asp:Content>
