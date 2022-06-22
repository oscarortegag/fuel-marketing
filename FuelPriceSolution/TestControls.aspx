<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="TestControls.aspx.vb" Inherits="FuelPriceSolution.TestControls" %>
<%@ Register Assembly="Telerik.Web.UI" Namespace="Telerik.Web.UI" TagPrefix="telerik" %>
<asp:Content ID="Content2" ContentPlaceHolderID="ScriptPlaceHolder" runat="server">
    <style>
        div.RadComboBox .rcbInputCell .rcbInput { padding-left: 30px;    }
    </style>
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="MainContent" runat="server">    
    <div class="right_col" role="main">
        <div>
            <div class="page-title">
			    <div class="title_left">
				    <h3>Prueba de controles personalizados</h3>
			    </div>
		    </div>
            <div class="clearfix"></div>
			<div class="row">
				<div class="col-md-12 col-sm-12 ">
					<div class="x_panel">
						<div class="x_title">
							<h2>Controles</h2>
							<div class="clearfix"></div>
						</div>
						<div class="x_content">
							<br />
							<div class="form-label-left input_mask">

                                <div class="row">
                                    <div class="col-md-4 col-sm-4 form-group has-feedback">									    
                                        <asp:TextBox runat="server" ID="txtDescrip" CssClass="form-control dropdown-span" />
                                        <span class="fa fa-line-chart form-control-feedback left" aria-hidden="true"></span>                        
                                    </div>
                                    
                                    <div class="col-md-4 col-sm-4 form-group has-feedback">
                                        <telerik:radcombobox runat="server" ID="radCombo1" CheckBoxes="true" EnableCheckAllItemsCheckBox="true" 
                                            style="width:100%;"
                                            AllowCustomText="true" EmptyMessage="-- Seleccione --"
                                            EnableAriaSupport="True" Skin="MetroTouch" Localization-AllItemsCheckedString="Todo Seleccionado" 
                                            Localization-CheckAllString="Seleccionar todo" Localization-ItemsCheckedString="elementos seleccionados" LoadingMessage="Cargando..." ></telerik:radcombobox>
                                        <span class="fa fa-fire form-control-feedback left" aria-hidden="true"></span>  
                                    </div>

                                    <div class="col-md-2 col-sm-2 form-group has-feedback">&nbsp;</div>

                                    <asp:UpdatePanel ID="UpdatePanel1" runat="server" class="col-md-2 col-sm-2" updatemode="Conditional" >
                                        <ContentTemplate>
                                            <asp:Panel ID="panel1" runat="server">
                                                <div class="col-md-12 col-sm-12 form-group has-feedback">                                         
                                                    <asp:Button runat="server" Text="Agregar" CssClass="btn btn-info form-control" ID="btnSave" style="bottom: 0px;" OnClick="btnSave_Click" visible="false"/>
                                                    <asp:LinkButton ID="ibtncountry" runat="server" Width="25%" Height="25%" OnClick="btnSave_Click" CausesValidation="false" CssClass="btn btn-info form-control">Test</asp:LinkButton>
                                                </div>
                                                </div>


                                            </div>
                                            <div class="row">
                                                <div class="col-md-4 col-sm-4 form-group has-feedback">	                                                                                                                           
                                                    <asp:Label runat="server" Text="" CssClass="form-control" ID="lblUno"></asp:Label>                                        
                                                    <span class="fa fa-line-chart form-control-feedback left" aria-hidden="true"></span>                        
                                                </div>
                                                <div class="col-md-8 form-group has-feedback">&nbsp;</div>
                                            </div>
                                            </asp:Panel>
                                            
                                    </ContentTemplate>
                                    <Triggers>
                                        <%--<asp:AsyncPostBackTrigger ControlID="btnSave" EventName="Click" />
                                        <asp:AsyncPostBackTrigger ControlID="ibtncountry" EventName="Click" />--%>
                                        <asp:PostBackTrigger ControlID="btnSave" />
                                        <asp:PostBackTrigger ControlID="ibtncountry" />
                                    </Triggers>
                                </asp:UpdatePanel>

                            
                        </div>
                    </div>
				</div>				
            </div>
        </div>
    </div>  
</asp:Content>
