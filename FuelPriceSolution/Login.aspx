<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Login.aspx.vb" Inherits="FuelPriceSolution.Login" %>

<!DOCTYPE html>

<html runat="server" xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <meta http-equiv="Content-Type" content="text/html; charset=UTF-8" />

    <meta charset="utf-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Fuel Marketing | Login</title>

     <link rel="icon" href="/Content/images/favicon.ico" type="image/ico" />
        
        <!-- jQuery -->
        <script src="/Content/vendors/jquery/dist/jquery.min.js"></script>
        <script src="/Scripts/jquery-3.5.1.min.js"></script>

        <!-- Bootstrap -->
        <script src="/Content/vendors/bootstrap/dist/js/bootstrap.bundle.min.js"></script>        
        <link href="/Content/vendors/bootstrap/dist/css/bootstrap.min.css" rel="stylesheet">

        <!-- moment -->
        <script src="/Content/vendors/moment/min/moment.min.js"></script>

        <!-- Combobox MultiSelect -->
        <link href="/Content/multiselect3/css/bootstrap-multiselect.min.css" rel="stylesheet" />
        <script src="/Content/multiselect3/js/bootstrap-multiselect.min.js"></script>

        <!-- Font Awesome -->
        <link href="/Content/vendors/font-awesome/css/font-awesome.min.css" rel="stylesheet">
        <!-- NProgress -->
        <%--<link href="<% =ServerPath %>/Content/vendors/nprogress/nprogress.css" rel="stylesheet">--%>
        <!-- iCheck -->
        <link href="/Content/vendors/iCheck/skins/flat/green.css" rel="stylesheet">
	
        <!-- bootstrap-progressbar -->
        <link href="/Content/vendors/bootstrap-progressbar/css/bootstrap-progressbar-3.3.4.min.css" rel="stylesheet">
        <!-- JQVMap -->
        <link href="/Content/vendors/jqvmap/dist/jqvmap.min.css" rel="stylesheet"/>
        <!-- bootstrap-daterangepicker -->
        <link href="/Content/vendors/bootstrap-daterangepicker/daterangepicker.css" rel="stylesheet">

        <!-- Custom Theme Style -->
        <link href="/Content/build/css/custom.min.css" rel="stylesheet">

        <!-- Developer Tema Extra Styles -->
        <link href="/Content/ExtraStyle.css" rel="stylesheet" />  

        <!-- Select2 -->
        <link href="/Content/vendors/select2/dist/css/select2.css" rel="stylesheet" />

        <%--Ladda--%>
        <link href="/Content/css/ladda/ladda-themeless.min.css" rel="stylesheet" />

        <!-- DatePicker -->
        <link href="/Content/vendors/bootstrap-datetimepicker/build/css/bootstrap-datetimepicker.min.css" rel="stylesheet" />
        <script src="/Content/vendors/bootstrap-datetimepicker/build/js/bootstrap-datetimepicker.min.js"></script>    
</head>
<body class="login" runat="server">
<video autoplay muted loop class="myVideo">
<source src="videoh.mp4" type="video/mp4">
Your browser does not support HTML5 video.
</video>

<div class="VideoContent" >

<div class="login_wrapper" >
    <div class="animate form login_form">

        <section class="login_content">
          <form runat="server">
            <div class="row">
                <div class="col-12">
                    <h1>Login</h1>
                </div>
            </div>
            <div class="row">&nbsp;</div>
            <div class="row">
                <div class="col-12">
                    <asp:TextBox runat="server" ID="txtUserName" type="text" CssClass="form-control" placeholder="Username" required="" />
                </div>
            </div>
            <div class="row">
                <div class="col-12">
                    <asp:TextBox runat="server" ID="txtUserPass" type="password" CssClass="form-control" placeholder="Password" required="" />
                </div>
            </div>
            <div class="row">
                <div class="col-12">
                    <asp:CheckBox runat="server" id="chkPersistCookie" Text="&nbsp;&nbsp;  Permanecer Conectado" />
                </div>
            </div>
            <div class="row">
                <div class="col-12">
                    <asp:Button runat="server" ID="btnLogin" OnClick="btnLogin_Click" CssClass="myBtn" Text="Log In" />
                </div>
            </div>
            <div class="row">
                <div class="col-12">
                    <a class="reset_pass" href="#">¿Olvidaste tu contrase&ntilde;a?</a>
                </div>
            </div>                          
             
            <div class="clearfix"></div>
            <div class="separator">
            
           
            <div class="clearfix"></div>
            
            <div>
                <h1><img src="Content/images/LogoFuelBlanco.png" style="height:50px" /> </h1>
                <%--<p>©2016 All Rights Reserved. Gentelella Alela! is a Bootstrap 3 template. Privacy and Terms</p>--%>
            </div>
            </div>
        </form>
        </section>
    </div>
  
</div>
</div>


</body>
</html>
