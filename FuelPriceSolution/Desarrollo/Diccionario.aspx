<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Site.Master" CodeBehind="Diccionario.aspx.vb" Inherits="FuelPriceSolution.Diccionario" %>
<%@ Register TagPrefix="telerik" Namespace="Telerik.Web.UI" Assembly="Telerik.Web.UI" %>
<asp:Content ID="Content2" ContentPlaceHolderID="ScriptPlaceHolder" runat="server">
</asp:Content>
<asp:Content ID="Content3" ContentPlaceHolderID="MainContent" runat="server">
    <asp:TextBox runat="server" ID="txtLlave" placeholder="_LLAVE_"></asp:TextBox>
    <asp:TextBox runat="server" ID="txtBase" placeholder="TextoBase"></asp:TextBox>
    
    <div class="col-12">&nbsp;</div>
    <div class="row">
        <label>Ficha</label>
        <div class="col-8">
            
            <asp:DropDownList CssClass="form-control" runat="server" ID="cmbDocumentos"></asp:DropDownList>
        </div>
        <div class="col-2">
            <button type="button" class="btn btn-success" data-toggle="modal" data-target="#exampleModal"><i class="fa fa-plus"></i></button>
            <!-- Button trigger modal -->


<!-- Modal -->
<div class="modal fade" id="exampleModal" tabindex="-1" aria-labelledby="exampleModalLabel" aria-hidden="true">
  <div class="modal-dialog">
    <div class="modal-content">
      <div class="modal-header">
        <h5 class="modal-title" id="exampleModalLabel">Agregar Ficha</h5>
        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
          <span aria-hidden="true">&times;</span>
        </button>
      </div>
      <div class="modal-body">
        <label>Nombre de Nueva Ficha (mismo nombre del archivo sin la extension '.aspx')</label>
        <asp:TextBox runat="server" ID="txtNuevaFicha" />
      </div>
      <div class="modal-footer">
        <button type="button" class="btn btn-secondary" data-dismiss="modal">Cancelar</button>
        <asp:LinkButton runat="server" ID="btnAgregaFicha" OnClick="btnAgregaFicha_Click" CssClass="btn btn-success">Agregar</asp:LinkButton>
      </div>
    </div>
  </div>
</div>
        </div>
    </div>
    <div class="col-12">&nbsp;</div>
    <asp:LinkButton runat="server" ID="btnBusca" CssClass="btn btn-success" Text="Busca" OnClick="btnBusca_Click"></asp:LinkButton>
    <div class="row">
        <div class="col-lg-6">
            <div class="card shadow mb-4">
            <div class="card-header py-3">
                <h6 class="m-0 font-weight-bold text-primary">Diccionario Disponible</h6>
            </div>
            <div class="card-body">
                <table>
        <tr>
            <td>Llave</td>
            <td>Frase</td>
            <td>Agrega</td>
        </tr>
         
        <asp:Repeater runat="server" ID="rptFraces" OnItemCommand ="rptFraces_ItemCommand">
            <ItemTemplate>
                    <tr>
                <td style="width:45%">
                    <%#Eval("IDM_LLAVE") %>
                </td>
                <td style="width:45%"><%#Eval("IDM_BASE") %></td>
                <td style="width:10%"><asp:LinkButton runat="server" CommandName="agrega" CommandArgument='<%#Eval("IDM_LLAVE") %>'><i class="fa fa-plus-circle" aria-hidden="true"></i></asp:LinkButton></td>
                </tr>
            </ItemTemplate>
        </asp:Repeater>
        
    </table>
            </div>
        </div>
        </div>
        
        <div class="col-lg-6">
            <div class="card shadow mb-4">
            <div class="card-header py-3">
                <h6 class="m-0 font-weight-bold text-primary">Diccionario en <asp:Label runat="server" ID="txtFichaSeleccionada" /> </h6>
            </div>
            <div class="card-body">
                <table>
        <tr>
            <td>Quita</td>
            <td>Llave</td>
            <td>Frase</td>
        </tr>
         
        <asp:Repeater runat="server" ID="rptEnFicha" OnItemCommand ="rptEnFicha_ItemCommand">
            <ItemTemplate>
                    <tr>
                <td style="width:10%"><asp:LinkButton runat="server" CommandName="quita" CommandArgument='<%#Eval("IDM_LLAVE") %>'><i class="fa fa-trash" aria-hidden="true"></i></asp:LinkButton></td>
                <td style="width:45%">
                    <%#Eval("IDM_LLAVE") %>
                </td>
                <td style="width:45%"><%#Eval("IDM_BASE") %></td>
                
                </tr>
            </ItemTemplate>
        </asp:Repeater>
        
    </table>
            </div>
        </div>
        </div>
    </div>
    
    
</asp:Content>
