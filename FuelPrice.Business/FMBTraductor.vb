Imports FuelPrice.DataAccess
Imports FMObjects
Public Class FMBTraductor
    Shared da As FMDataAccess = New FMDataAccess()
    Shared Comun As String = "FuelPriceComun"
    Private TablaTraducida As New DataTable()

    Public Sub New()
    End Sub
    Public Sub New(ByVal NomDocumento As String)
        Try
            Dim query As String = "SELECT FP_Idioma.*
            FROM FP_Idioma_Ficha
            inner join FP_Ficha on FP_Idioma_Ficha.IDF_FCH_CODIGO=FP_Ficha.FCH_CODIGO
            inner join FP_Idioma on FP_Idioma_Ficha.IDF_IDM_LLAVE=FP_Idioma.IDM_LLAVE 
            where FP_Ficha.FCH_NOMBRE_FICHA ='" + NomDocumento + "'"
            Dim dt As DataTable = da.Consulta(query, Comun)
            Dim foundRows As DataRow() = dt.Select("IDM_LLAVE = ''")
            TablaTraducida = dt
        Catch ex As Exception
            Dim query As String = "SELECT '' as IDM_LLAVE,'' as IDM_DESCRIPCION,'' as IDM_BASE"
            Dim dt As DataTable = da.Consulta(query, Comun)
        End Try


    End Sub

    Public Function getFichas() As DataTable
        Dim query As String = "select FCH_CODIGO,FCH_NOMBRE_FICHA from FP_Ficha"
        Dim dt As DataTable = da.Consulta(query, Comun)
        Return dt
    End Function

    Public Function getDiccionarioCompleto() As DataTable
        Dim query As String = "select [IDM_LLAVE]
                                ,[IDM_DESCRIPCION]
                                ,[IDM_BASE] from FP_Idioma order by IDM_LLAVE, IDM_BASE"
        Dim dt As DataTable = da.Consulta(query, Comun)
        Return dt
    End Function

    Public Function getDiccionarioParaFicha(ByVal laFicha As String) As DataTable
        Dim query As String = "select idm_llave,idm_base from fp_idioma where IDM_LLAVE not in(
                            select IDM_LLAVE from FP_Idioma
                            inner join FP_Idioma_Ficha on FP_Idioma.IDM_LLAVE = FP_Idioma_Ficha.IDF_IDM_LLAVE
                            where FP_Idioma_Ficha.IDF_FCH_CODIGO=" + laFicha + ")
"
        Dim dt As DataTable = da.Consulta(query, Comun)
        Return dt
    End Function

    Public Sub AgregaLlaveFicha(Ficha As String, Llave As String)
        Dim query As String = "insert into [FP_Idioma_Ficha] (IDF_FCH_CODIGO,IDF_IDM_LLAVE) values (" + Ficha + ",'" + Llave + "')"
        da.Consulta(query, Comun)
    End Sub

    Public Sub putAgregaFicha(Ficha As String)
        Dim queryHay As String = "select FCH_CODIGO from FP_Ficha where FCH_NOMBRE_FICHA='" + Ficha + "'"
        If Not da.IfExist(queryHay, Comun) Then
            Dim query As String = "Innsert into FP_Ficha (FCH_NOMBRE_FICHA) values ('" + Ficha + "')"
            da.Consulta(query, Comun)
        End If
    End Sub

    Public Sub QuitaLlaveFicha(Ficha As String, Llave As String)
        Dim query As String = "delete from FP_Idioma_Ficha where IDF_FCH_CODIGO=" + Ficha + " And IDF_IDM_LLAVE='" + Llave + "'"
        da.Consulta(query, Comun)

    End Sub

    Public Function getDiccionarioDeFicha(ByVal laFicha As String) As DataTable
        Dim query As String = "select IDM_LLAVE,IDM_BASE from FP_Idioma
                               inner join FP_Idioma_Ficha on FP_Idioma.IDM_LLAVE = FP_Idioma_Ficha.IDF_IDM_LLAVE
                               where FP_Idioma_Ficha.IDF_FCH_CODIGO=" + laFicha
        Dim dt As DataTable = da.Consulta(query, Comun)
        Return dt
    End Function

    Public Function getDiccionarioParaFicha(ByVal laFicha As String, ByVal LLave As String, ByVal Base As String) As DataTable
        Dim query As String = "select idm_llave,idm_base from fp_idioma where IDM_LLAVE not in(
                            select IDM_LLAVE from FP_Idioma
                            inner join FP_Idioma_Ficha on FP_Idioma.IDM_LLAVE = FP_Idioma_Ficha.IDF_IDM_LLAVE
                            where FP_Idioma_Ficha.IDF_FCH_CODIGO=" + laFicha + ") and IDM_LLAVE like '%" + LLave + "%' 
                                and IDM_BASE like '%" + Base + "%' 
"
        Dim dt As DataTable = da.Consulta(query, Comun)
        Return dt
    End Function
    Public Function getDiccionarioCompleto(ByVal LLave As String, ByVal Base As String) As DataTable
        Dim query As String = "select [IDM_LLAVE]
                                ,[IDM_DESCRIPCION]
                                ,[IDM_BASE]from FP_Idioma 
                                where IDM_LLAVE like '%" + LLave + "%' 
                                and IDM_BASE like '%" + Base + "%' 
                                order by IDM_LLAVE, IDM_BASE"
        Dim dt As DataTable = da.Consulta(query, Comun)
        Return dt
    End Function

    Public Function CulturasDisponibles() As String()
        Dim query As String = "select * from FP_Idioma where IDM_LLAVE=''"
        Dim dt As DataTable = da.Consulta(query, Comun)
        Dim Culturas(dt.Columns.Count - 4) As String

        For i = 0 To Culturas.Length - 1
            Culturas(i) = Mid(dt.Columns(i + 3).ColumnName.Replace("_", "-"), 5)
        Next
        Return Culturas
    End Function

    Public Function GetTranslations(ByVal NomDocumento As String, ByVal Culture As String) As List(Of Translated)
        Dim _listado As New List(Of Translated)

        Try
            Dim query As String = "SELECT FP_Idioma.IDM_LLAVE, FP_IDIOMA.IDM_BASE, FP_Idioma.IDM_" & Culture.Replace("-", "_") & " " & vbNewLine &
                      "FROM FP_Idioma_Ficha " & vbNewLine &
                      "inner join FP_Ficha on FP_Idioma_Ficha.IDF_FCH_CODIGO=FP_Ficha.FCH_CODIGO " & vbNewLine &
                      "inner join FP_Idioma on FP_Idioma_Ficha.IDF_IDM_LLAVE=FP_Idioma.IDM_LLAVE " & vbNewLine &
                      "where FP_Ficha.FCH_NOMBRE_FICHA ='" + NomDocumento + "'"
            Dim dt As DataTable = da.Consulta(query, Comun)

            If Not dt Is Nothing Then
                If dt.Rows.Count > 0 Then
                    For index = 0 To dt.Rows.Count - 1
                        Dim _tr As New Translated
                        _tr.Key = dt.Rows(index)(0).ToString
                        _tr.DefaultString = dt.Rows(index)(1).ToString
                        _tr.TranslatedString = dt.Rows(index)(2).ToString
                        _listado.Add(_tr)
                    Next
                End If
            End If
        Catch ex As Exception
            _listado = New List(Of Translated)
        End Try

        Return _listado
    End Function

    Public Function Traduce(ByVal Llave As String, ByVal Cultura As String) As String
        If Cultura Is Nothing Then
            Cultura = "es-MX"
        End If
        Dim foundRows As DataRow() = TablaTraducida.Select("IDM_LLAVE = '" + Llave + "'")
        If foundRows.Length > 0 Then
            Try
                Dim tradCult As String = foundRows(0)("IDM_" + Cultura.Replace("-", "_"))
                If tradCult = "" Then
                    Throw New Exception("No Existe Traduccion")
                Else
                    Return tradCult
                End If
            Catch ex As Exception
                Dim TradBase As String = foundRows(0)("IDM_BASE")
                Return TradBase
            End Try

        Else
            'no Existe la llave
            Return Llave
        End If
    End Function

End Class
