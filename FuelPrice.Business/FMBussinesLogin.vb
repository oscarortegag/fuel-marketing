Imports FuelPrice.DataAccess
Public Class FMBussinesLogin

    Private da As FMDataAccess = New FMDataAccess()
    Private Comun As String = "FuelPriceComun"
    Private Cliente As String = "FuelPrice"
    Private Cultura As String = "es-MX"

    Public Sub New(ByVal _Cultrua As String)
        Cultura = _Cultrua
    End Sub

    Public Sub New()
    End Sub
    Public Function getNegociosdelUsuario(ByVal idUsuario As String) As DataTable
        Dim qNegociosUsuario As String = "select UNE_CLI_CODIGO as CLI_CODIGO, '' as CLI_DESCRIPCION from FP_Usuario_Negocio where UNE_ID_USUARIO='" + idUsuario + "'"
        Dim dt As DataTable
        dt = da.Consulta(qNegociosUsuario, Comun)

        If dt.Rows.Count > 0 Then
            Dim qNombres As String
            Dim dtNombres As DataTable
            For i = 0 To dt.Rows.Count - 1
                qNombres = "select CLI_DESCRIPCION from FP_Cliente where CLI_CODIGO='" + dt.Rows(i)("CLI_CODIGO").ToString() + "'"
                dtNombres = da.Consulta(qNombres, Cliente)
                Try
                    dt.Rows(i)("CLI_DESCRIPCION") = dtNombres.Rows(0)("CLI_DESCRIPCION").ToString()
                Catch ex As Exception
                    dt.Rows(i)("CLI_DESCRIPCION") = "Err-" + dt.Rows(i)(0)
                End Try


            Next
        End If

        Return dt
    End Function
    Public Function ChecaLogin(usuario As String, password As String) As DataTable
        Try

            Dim sLogin As String = "Select * from AspNetUsers where (email='" + usuario + "' or UserName='" + usuario + "') and PasswordHAsh=HASHBYTES('SHA2_512','" + password + "')"
            Dim dtLogin As New DataTable()
            dtLogin = da.Consulta(sLogin, Cliente)
            Return dtLogin
        Catch ex As Exception
            Throw New Exception("ErrLogin - " + ex.Message)
        End Try
    End Function
    Public Function LeeUsr(usrID As String) As DataTable
        Try

            Dim sLogin As String = "Select * from AspNetUsers where id='" + usrID + "'"
            Dim dtLogin As New DataTable()
            dtLogin = da.Consulta(sLogin, Cliente)
            Return dtLogin
        Catch ex As Exception
            Throw New Exception("ErrLogin - " + ex.Message)
        End Try
    End Function
    Public Function ArbolMenuPrfl(idUsuario As String) As DataTable
        Try
            Dim query = "select 
                    Cat_Menu.IdMenu
                    ,Cat_Menu.Nombre
                    ,Cat_Menu.IdMenuPadre
                    ,Cat_Menu.Url
                    ,Cat_Menu.Orden
                    ,Cat_Menu.Desc_Class
                    ,PerfilesMenus.Estatus
                    ,isnull(Cat_Menu.Prms_API,'') as Prms_API
                    from Cat_Menu
                    inner join PerfilesMenus on Cat_Menu.IdMenu=PerfilesMenus.IdMenu
                    inner join AspNetPerfiles on PerfilesMenus.IdPerfil=AspNetPerfiles.Id
                    inner join AspNetUserPerfiles on AspNetPerfiles.Id=AspNetUserPerfiles.PerfilId
                    where AspNetUserPerfiles.UserId='" + idUsuario + "' and Cat_Menu.Estatus=1 order by Cat_Menu.Orden asc"

            Dim dt As New DataTable()
            dt = da.Consulta(query, Comun)
            Return dt
        Catch ex As Exception
            Throw New Exception("ErrCatMenu")
        End Try


    End Function

    Public Sub UpdateUser(idUsr As String, UserName As String, Nombre As String, Apat As String, Amat As String, Telefono As String, Email As String, Idioma As String)
        Dim query As String = "update AspNetUsers set 
                               [Email]='" + Email + "' 
                              ,[PhoneNumber] ='" + Telefono + "'
                              ,[UserName] ='" + UserName + "'
                              ,[Nombre] ='" + Nombre + "'
                              ,[ApellidoPaterno] ='" + Apat + "'
                              ,[ApellidoMaterno] ='" + Amat + "'
                              ,[IDM_CULTURA] ='" + Idioma + "' where id='" + idUsr + "'"
        Dim dt As New DataTable()
        dt = da.Consulta(query, Cliente)
    End Sub

    Public Sub UpdatePasswd(idUsr As String, Password As String)
        Dim query As String = "update AspNetUsers set PasswordHash=HASHBYTES('SHA2_512','" + Password + "') where id='" + idUsr + "'"
        Dim dt As New DataTable()
        dt = da.Consulta(query, Cliente)
    End Sub
End Class
