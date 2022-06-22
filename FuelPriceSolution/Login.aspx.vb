Public Class Login
    Inherits System.Web.UI.Page
    Public ServerPath As String
    Private Datos As New FuelPrice.Business.FMBussinesLogin()
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim absoluteurl As String = HttpContext.Current.Request.Url.AbsoluteUri
        ServerPath = absoluteurl.Replace(HttpContext.Current.Request.Url.LocalPath, "")
    End Sub

    Protected Sub btnLogin_Click(sender As Object, e As EventArgs)
        If ValidateUser(txtUserName.Text, txtUserPass.Text) Then
            Dim tkt As FormsAuthenticationTicket
            Dim cookiestr As String
            Dim ck As HttpCookie
            tkt = New FormsAuthenticationTicket(1, txtUserName.Text, DateTime.Now, DateTime.Now.AddMinutes(30), chkPersistCookie.Checked, "custom data")
            ck = New HttpCookie(FormsAuthentication.FormsCookieName, cookiestr)
            If chkPersistCookie.Checked Then
                ck.Expires = tkt.Expiration
            End If
            ck.Path = FormsAuthentication.FormsCookiePath
            Response.Cookies.Add(ck)
            Dim strRedirect As String
            strRedirect = Request("ReturnUrl")
            If strRedirect = Nothing Then
                strRedirect = "Default.aspx"
            End If
            Response.Redirect(strRedirect, True)
        Else
            Response.Redirect("Login.aspx", True)
        End If
    End Sub

    Private Function ValidateUser(usuario As String, password As String) As Boolean
        Dim dtUsuario As DataTable = Datos.ChecaLogin(usuario, password)
        If dtUsuario IsNot (Nothing) Then
            If dtUsuario.Rows.Count > 0 Then
                HttpContext.Current.Session("MaestroCliente") = dtUsuario.Rows(0)("MCL_CODIGO")
                HttpContext.Current.Session("IdUsr") = dtUsuario.Rows(0)("id")
                HttpContext.Current.Session("Cultura") = dtUsuario.Rows(0)("IDM_CULTURA")
                HttpContext.Current.Session("UserName") = dtUsuario.Rows(0)("UserName")
                Dim dtNegocios As DataTable = Datos.getNegociosdelUsuario(dtUsuario.Rows(0)("id"))
                HttpContext.Current.Session("Negocios") = dtNegocios
            End If
            Return True
        Else
            Return False
        End If
    End Function
End Class