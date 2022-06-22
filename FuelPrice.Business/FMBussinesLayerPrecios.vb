Imports System.IO
Imports System.Data
Imports FuelPrice.DataAccess
Imports ExcelDataReader

Partial Public Class FMBussinesLayer
    Public Function SetPreciosReferencia(ByVal TblPrecios, ByVal Terminal) As Boolean
        Dim result As Boolean

        Try
            Dim inserta As Integer

            For index = 0 To TblPrecios.Rows.Count - 1

                consulta = "Select PRF_CODIGO From FP_PreciosReferenciaTerminal " & vbNewLine &
                            "Where PRF_TSU_CODIGO = '" & Terminal & "' " & vbNewLine &
                            " And PRF_CON_CODIGO = '" & TblPrecios.Rows(index)("IdProd").ToString & "' " & vbNewLine &
                            "And PRF_FECHA = '" & Convert.ToDateTime(TblPrecios.Rows(index)("Fecha")).ToString("yyyy-MM-dd") & "'"
                Dim _tbl = da.Consulta(consulta, _fuelPrice)

                If IsNothing(_tbl) Then
                    consulta = "INSERT INTO FP_PreciosReferenciaTerminal
                                    (PRF_TSU_CODIGO
                                    ,PRF_CON_CODIGO
                                    ,PRF_PRECIO
                                    ,PRF_FECHA
                                    ,PRF_HORAAPLICACION)
                                VALUES
                                    ('" & Terminal & "','" & TblPrecios.Rows(index)("IdProd").ToString & "'
                                    ,'" & TblPrecios.Rows(index)("Precio").ToString & "','" & Convert.ToDateTime(TblPrecios.Rows(index)("Fecha")).ToString("yyyy-MM-dd") & "','00:00:00.000')"

                    inserta = da.NoQuery(consulta, _fuelPrice)
                Else
                    If _tbl.Rows.Count <= 0 Then
                        consulta = "INSERT INTO FP_PreciosReferenciaTerminal
                                       (PRF_TSU_CODIGO
                                       ,PRF_CON_CODIGO
                                       ,PRF_PRECIO
                                       ,PRF_FECHA
                                       ,PRF_HORAAPLICACION)
                                 VALUES
                                       ('" & Terminal & "','" & TblPrecios.Rows(index)("IdProd").ToString & "'
                                       ,'" & TblPrecios.Rows(index)("Precio").ToString & "','" & Convert.ToDateTime(TblPrecios.Rows(index)("Fecha")).ToString("yyyy-MM-dd") & "','00:00:00.000')"

                        inserta = da.NoQuery(consulta, _fuelPrice)
                    Else
                        consulta = "UPDATE FP_PreciosReferenciaTerminal " & vbNewLine &
                                "SET PRF_PRECIO = " & TblPrecios.Rows(index)("Precio").ToString & vbNewLine &
                                "WHERE PRF_CODIGO = " & _tbl.Rows(0)(0).ToString

                        inserta = da.NoQuery(consulta, _fuelPrice)
                    End If
                End If
            Next

            result = True
        Catch ex As Exception
            result = False
        End Try

        Return result
    End Function
    Public Function GetReportePrecios(ByVal lstTerminal, ByVal lstProductos, ByVal Desde, ByVal Hasta, ByVal IdMaestro, Optional ByVal EsBio = 0) As DataTable
        Dim result As New DataTable
        result.Columns.Add("TSU_CODIGO")
        result.Columns.Add("TSU_DESCRIPCION")
        result.Columns.Add("CON_CODIGO")
        result.Columns.Add("CON_DESCRIPCION")
        result.Columns.Add("PRF_PRECIO")
        result.Columns.Add("PRF_FECHA")

        Try
            Dim tblTer As DataTable = GetTerminalesList(lstTerminal)
            Dim tblPro As DataTable = GetProductosList(lstProductos, IdMaestro, EsBio)

            consulta = "Select PRF_TSU_CODIGO, PRF_CON_CODIGO, PRF_PRECIO, PRF_FECHA " & vbNewLine &
                       "From FP_PreciosReferenciaTerminal " & vbNewLine &
                       "Where PRF_TSU_CODIGO In (" & lstTerminal & ") And " & vbNewLine &
                       "PRF_CON_CODIGO In (" & lstProductos & ") And (PRF_FECHA Between '" & Desde & "' And '" & Hasta & "')"

            Dim tblPrecios = da.Consulta(consulta, _fuelPrice)

            Dim vLINQ = (From DT1 In tblPrecios.AsEnumerable
                         Join DT2 In tblTer.AsEnumerable
                            On DT1.Field(Of String)("PRF_TSU_CODIGO") Equals DT2.Field(Of Guid)("TSU_CODIGO").ToString
                         Join DT3 In tblPro.AsEnumerable
                            On DT1.Field(Of String)("PRF_CON_CODIGO") Equals DT3.Field(Of Guid)("CON_CODIGO").ToString
                         Select New With
                        {
                            .TSU_CODIGO = DT1.Field(Of String)("PRF_TSU_CODIGO"),
                            .TSU_DESCRIPCION = DT2.Field(Of String)("TSU_DESCRIPCION"),
                            .CON_CODIGO = DT1.Field(Of String)("PRF_CON_CODIGO"),
                            .CON_DESCRIPCION = DT3.Field(Of String)("CON_DESCRIPCION"),
                            .PRF_PRECIO = DT1.Field(Of Decimal)("PRF_PRECIO").ToString("N6"),
                            .PRF_FECHA = DT1.Field(Of DateTime)("PRF_FECHA").ToString("dd/MM/yyyy")
                        }).Distinct().ToList()

            For Each item In vLINQ
                result.Rows.Add(item.TSU_CODIGO, item.TSU_DESCRIPCION, item.CON_CODIGO, item.CON_DESCRIPCION, item.PRF_PRECIO, item.PRF_FECHA)
            Next

        Catch ex As Exception
            result = New DataTable
        End Try

        Return result
    End Function
    Public Function GetTerminalesList(ByVal lstTerminal) As DataTable
        Dim result As DataTable

        Try
            consulta = "Select TSU_CODIGO, TSU_DESCRIPCION From FP_MAE_TerminalSuministro " & vbNewLine &
                       "Where TSU_CODIGO In(" & lstTerminal & ")"

            result = da.Consulta(consulta, _fuelPriceComun)
        Catch ex As Exception
            result = New DataTable()
        End Try

        Return result
    End Function
    Public Function GetProductosList(ByVal lstProductos) As DataTable
        Dim result As DataTable

        Try
            consulta = "Select CON_CODIGO, CON_DESCRIPCION From FP_MAE_Combustible " & vbNewLine &
                       "Where CON_CODIGO In(" & lstProductos & ")"

            result = da.Consulta(consulta, _fuelPriceComun)
        Catch ex As Exception
            result = New DataTable()
        End Try

        Return result
    End Function
    Public Function GetProductosList(ByVal lstProductos, ByVal IdMaestro, ByVal EsBio) As DataTable
        Dim result As DataTable

        Try
            consulta = "Select MC.CON_CODIGO, MC.CON_DESCRIPCION, MC.CON_DESCRIPCION_CORTA " & vbNewLine &
                       "From FP_MAE_Combustible MC Inner Join FP_ClientesProductos CP On Convert(varchar(50), MC.CON_CODIGO) = Convert(varchar(50), CP.CPO_CON_CODIGO) And CP.CPO_MCL_CODIGO = " & IdMaestro & vbNewLine &
                       "Where MC.CON_ESBIO = " & EsBio & " And Convert(varchar(50), MC.CON_CODIGO) In(" & lstProductos & ")"

            result = da.Consulta(consulta, _fuelPriceComun)
        Catch ex As Exception
            result = New DataTable()
        End Try

        Return result
    End Function
    Public Function SetVolRealObjetivo(ByVal TblPrecios) As Boolean
        Dim result As Boolean

        Try
            Dim inserta As Integer

            For index = 0 To TblPrecios.Rows.Count - 1

                consulta = "SELECT VRO_CODIGO FROM FP_VolVta_Real_Objetivo " & vbNewLine &
                           "Where VRO_CLI_CODIGO = '" & TblPrecios.Rows(index)("IdProd").ToString & "' " & vbNewLine &
                           "And VRO_CON_CODIGO = '" & TblPrecios.Rows(index)("IdNegocio").ToString & "' " & vbNewLine &
                           "And VRO_FECHA = '" & Convert.ToDateTime(TblPrecios.Rows(index)("Fecha")).ToString("yyyy-MM-dd") & "'"

                Dim _tbl = da.Consulta(consulta, _fuelPrice)

                If IsNothing(_tbl) Then
                    consulta = "INSERT INTO FP_VolVta_Real_Objetivo
                                       (VRO_CLI_CODIGO
                                       ,VRO_CON_CODIGO
                                       ,VRO_VOL_REAL
                                       ,VRO_VOL_OBJETIVO
                                       ,VRO_FECHA)
                                 VALUES
                                       ('" & TblPrecios.Rows(index)("IdNegocio").ToString & "','" & TblPrecios.Rows(index)("IdProd").ToString & "'
                                       ,'" & IIf(TblPrecios.Rows(index)("Real").ToString = "-", "0", TblPrecios.Rows(index)("Real").ToString) & "',
                                       '" & IIf(TblPrecios.Rows(index)("Objetivo").ToString = "-", "0", TblPrecios.Rows(index)("Objetivo").ToString) & "',
                                       '" & Convert.ToDateTime(TblPrecios.Rows(index)("Fecha")).ToString("yyyy-MM-dd") & "')"

                ElseIf _tbl.Rows.Count <= 0 Then
                    consulta = "INSERT INTO FP_VolVta_Real_Objetivo
                                       (VRO_CLI_CODIGO
                                       ,VRO_CON_CODIGO
                                       ,VRO_VOL_REAL
                                       ,VRO_VOL_OBJETIVO
                                       ,VRO_FECHA)
                                 VALUES
                                       ('" & TblPrecios.Rows(index)("IdNegocio").ToString & "','" & TblPrecios.Rows(index)("IdProd").ToString & "'
                                       ,'" & IIf(TblPrecios.Rows(index)("Real").ToString = "-", "0", TblPrecios.Rows(index)("Real").ToString) & "',
                                       '" & IIf(TblPrecios.Rows(index)("Objetivo").ToString = "-", "0", TblPrecios.Rows(index)("Objetivo").ToString) & "',
                                       '" & Convert.ToDateTime(TblPrecios.Rows(index)("Fecha")).ToString("yyyy-MM-dd") & "')"
                Else
                    consulta = "Update FP_VolVta_Real_Objetivo Set" & vbNewLine

                    If TblPrecios.Rows(index)("Real").ToString <> "-" Then
                        consulta = consulta & "VRO_VOL_REAL = '" & TblPrecios.Rows(index)("Real").ToString & "'," & vbNewLine
                    End If

                    If TblPrecios.Rows(index)("Objetivo").ToString <> "-" Then
                        consulta = consulta & "VRO_VOL_OBJETIVO = '" & TblPrecios.Rows(index)("Objetivo").ToString & "'" & vbNewLine
                    End If

                    If consulta.Trim().EndsWith(",") Then
                        consulta = consulta.Trim()
                        consulta = Mid(consulta, 1, Len(consulta) - 1) & vbNewLine
                    End If

                    consulta = consulta & "Where VRO_CODIGO = '" & _tbl.Rows(0)(0).ToString & "'"
                End If

                inserta = da.NoQuery(consulta, _fuelPrice)
            Next

            result = True
        Catch ex As Exception
            result = False
        End Try

        Return result
    End Function
    Public Function GetVolRealObjetivo(ByVal TipoVol, ByVal LstNegocio, ByVal LstProducto, ByVal Desde, ByVal Hasta, ByVal IdMaestro) As DataTable
        Dim result As New DataTable
        result.Columns.Add("CLI_CODIGO")
        result.Columns.Add("CLIENTE")
        result.Columns.Add("CON_CODIGO")
        result.Columns.Add("PRODUCTO")
        result.Columns.Add("REAL")
        result.Columns.Add("OBJETIVO")
        result.Columns.Add("FECHA")

        Try
            consulta = "Select Convert(Varchar(50), CLI_CODIGO) As CLI_CODIGO, CLI_DESCRIPCION From FP_Cliente Where Convert(Varchar(50), CLI_CODIGO) In (" & LstNegocio & ")"
            Dim tblNeg = da.Consulta(consulta, _fuelPrice)

            Dim tblProd = GetProductosSelectionFormat(IdMaestro)

            consulta = "Select MCL_REGPVP From FP_MAE_Clientes Where MCL_CODIGO = " & IdMaestro
            Dim tblName As String = da.Consulta(consulta, _fuelPriceComun).Rows(0)(0).ToString

            If tblName.ToUpper() <> "FP_MAE_COMBUSTIBLE" Then
                For index = 0 To tblProd.Rows.Count - 1
                    LstProducto = LstProducto.Replace(tblProd.Rows(index)("CON_CODIGO"), tblProd.Rows(index)("VTP_CODIGO"))
                Next
            End If

            consulta = "Select VRO_FECHA, " & vbNewLine &
                       "VRO_VOL_REAL, " & vbNewLine &
                       "VRO_VOL_OBJETIVO, " & vbNewLine &
                       "Convert(varchar(50), VRO_CLI_CODIGO) as VRO_CLI_CODIGO, Convert(varchar(50), VRO_CON_CODIGO) As VRO_CON_CODIGO " & vbNewLine &
                       "From FP_VolVta_Real_Objetivo " & vbNewLine &
                       "Where Convert(varchar(50), VRO_CLI_CODIGO) In(" & LstNegocio & ") " & vbNewLine &
                       "And Convert(varchar(50), VRO_CON_CODIGO) In(" & LstProducto & ") " & vbNewLine &
                       "And VRO_FECHA Between '" & Desde & "' And '" & Hasta & "'"
            Dim tblVol = da.Consulta(consulta, _fuelPrice)

            Dim nomCampo As String = IIf(tblName.ToUpper() = "FP_MAE_COMBUSTIBLE", "CON_CODIGO", "VTP_CODIGO")

            Dim vLINQ = (From DT1 In tblVol.AsEnumerable
                         Join DT2 In tblNeg.AsEnumerable
                            On DT1.Field(Of String)("VRO_CLI_CODIGO").ToUpper() Equals DT2.Field(Of String)("CLI_CODIGO").ToString().ToUpper()
                         Join DT3 In tblProd.AsEnumerable
                            On DT1.Field(Of String)("VRO_CON_CODIGO").ToUpper() Equals DT3.Field(Of String)(nomCampo).ToString().ToUpper()
                         Select New With
                        {
                            .CLI_CODIGO = DT2.Field(Of String)("CLI_CODIGO").ToString,
                            .CLIENTE = DT2.Field(Of String)("CLI_DESCRIPCION"),
                            .CON_CODIGO = DT3.Field(Of String)("CON_CODIGO").ToString,
                            .PRODUCTO = DT3.Field(Of String)("CON_DESCRIPCION").ToString,
                            .REAL = DT1.Field(Of Decimal)("VRO_VOL_REAL").ToString,
                            .OBJETIVO = DT1.Field(Of Decimal)("VRO_VOL_OBJETIVO").ToString,
                            .FECHA = DT1.Field(Of DateTime)("VRO_FECHA").ToString("dd/MM/yyyy")
                        }).Distinct().ToList()

            For Each item In vLINQ
                result.Rows.Add(item.CLI_CODIGO, item.CLIENTE, item.CON_CODIGO, item.PRODUCTO, item.REAL, item.OBJETIVO, item.FECHA)
            Next

        Catch ex As Exception
            result = New DataTable
        End Try

        Return result
    End Function
    Public Function LastPVP(ByVal IdProducto, ByVal IdEstacion, ByVal Fecha) As DataTable
        Dim result As DataTable
        Try
            consulta = "Select Top 1 PVP_EST_CODIGO, PVP_CON_CODIGO, PVP_PRECIO, PVP_HORA, PVP_FECHA " & vbNewLine &
                       "From FP_RegistroPVP " & vbNewLine &
                       "Where PVP_EST_CODIGO = '" & IdEstacion & "' " & vbNewLine &
                       "And PVP_CON_CODIGO = '" & IdProducto & "' " & vbNewLine &
                       "And PVP_FECHA < '" & Fecha & "' " & vbNewLine &
                       "Order by PVP_FECHA Desc"

            result = da.Consulta(consulta, _fuelPriceComun)
        Catch ex As Exception
            result = New DataTable()
        End Try

        Return result
    End Function
    Public Function InformePVP(ByVal lstProductos, ByVal IdNegocio, ByVal Desde, ByVal Hasta, ByVal IdMaestro) As DataTable
        Dim result As New DataTable
        result.Columns.Add("IdNegocio")
        result.Columns.Add("Negocio")
        result.Columns.Add("IdProducto")
        result.Columns.Add("Producto")
        result.Columns.Add("Precio")
        result.Columns.Add("Fecha")
        result.Columns.Add("Hora")

        Try
            consulta = "Select Convert(VarChar(50), PVP_EST_CODIGO) As PVP_EST_CODIGO, Convert(VarChar(50), PVP_CON_CODIGO) As PVP_CON_CODIGO, PVP_PRECIO, PVP_HORA, PVP_FECHA " & vbNewLine &
                       "From FP_RegistroPVP " & vbNewLine &
                       "Where PVP_EST_CODIGO = '" & IdNegocio & "' " & vbNewLine &
                       "And PVP_CON_CODIGO In (" & lstProductos & ") " & vbNewLine &
                       "And PVP_FECHA Between '" & Desde & "' And '" & Hasta & "' " & vbNewLine &
                       "Order by PVP_FECHA Desc"

            Dim tblPvp = da.Consulta(consulta, _fuelPriceComun)

            Dim tblClientes = da.Consulta("Select Distinct Convert(VarChar(50), EST_CODIGO) As EST_CODIGO, EST_RAZON_SOCIAL From FP_Estacion Where EST_CODIGO = '" & IdNegocio & "'", _fuelPriceComun)

            If tblClientes.Rows.Count = 0 Then
                tblClientes = da.Consulta("Select Distinct Convert(VarChar(50), ES.NCO_CODIGO) As EST_CODIGO, ES.NCO_RAZONSOCIAL As EST_RAZON_SOCIAL From FP_Mercado_NegociosCompetencia ES Where NCO_CODIGO = '" & IdNegocio & "'", _fuelPrice)
            End If

            'Dim tblProductos = GetProductos()
            Dim tblProductos = GetProductosSelectionFormat(IdMaestro)

            'On DT1.Field(Of Guid)("PVP_CON_CODIGO").ToString Equals DT3.Field(Of Guid)("CON_CODIGO").ToString

            Dim vLINQ = (From DT1 In tblPvp.AsEnumerable
                         Join DT2 In tblClientes.AsEnumerable
                            On DT1.Field(Of String)("PVP_EST_CODIGO") Equals DT2.Field(Of String)("EST_CODIGO")
                         Join DT3 In tblProductos.AsEnumerable
                            On DT1.Field(Of String)("PVP_CON_CODIGO").ToString Equals DT3.Field(Of String)("CON_CODIGO").ToString
                         Select New With
                        {
                            .CLI_CODIGO = DT1.Field(Of String)("PVP_EST_CODIGO").ToString,
                            .CLI_DESCRIPCION = DT2.Field(Of String)("EST_RAZON_SOCIAL"),
                            .CON_CODIGO = DT1.Field(Of String)("PVP_CON_CODIGO").ToString,
                            .CON_DESCRIPCION = DT3.Field(Of String)("CON_DESCRIPCION"),
                            .PVP_PRECIO = DT1.Field(Of Double)("PVP_PRECIO").ToString("N2"),
                            .PVP_FECHA = DT1.Field(Of String)("PVP_FECHA"),
                            .PVP_HORA = DT1.Field(Of String)("PVP_HORA")
                        }).Distinct().ToList()

            For Each item In vLINQ
                result.Rows.Add(item.CLI_CODIGO, item.CLI_DESCRIPCION, item.CON_CODIGO, item.CON_DESCRIPCION, item.PVP_PRECIO, item.PVP_FECHA, item.PVP_HORA)
            Next

        Catch ex As Exception
            result = New DataTable
        End Try

        Return result
    End Function
    Public Function GetEstacionesByPermiso(ByVal NoPermiso) As DataTable
        Dim result As DataTable

        Try
            consulta = "Select EST_PERMISO, EST_RAZON_SOCIAL, EST_CODIGO From FP_Estacion Where EST_PERMISO LIKE '%" & NoPermiso & "%' Or EST_NUMERO_ESTACION LIKE '%" & NoPermiso & "%' And COALESCE(EST_RFC, '') <> '' Order by EST_RAZON_SOCIAL"

            result = da.Consulta(consulta, _fuelPriceComun)
        Catch ex As Exception
            result = New DataTable()
        End Try

        Return result
    End Function
    Public Function GetEstacionesByPermiso(ByVal NoPermiso, ByVal LstCli) As DataTable
        Dim result As DataTable

        Try
            consulta = "Select Distinct ES.EST_PERMISO, ES.EST_RAZON_SOCIAL, ES.EST_CODIGO "
            If NoPermiso <> "" Then
                consulta = consulta & "From " & dbfuelPriceComun.InitialCatalog.ToString & ".dbo.FP_Estacion ES " & vbNewLine
                consulta = consulta & "Inner Join " & dbfuelPrice.InitialCatalog.ToString & ".dbo.FP_Cliente CL On CL.CLI_EST_CODIGO = ES.EST_CODIGO " & vbNewLine
                consulta = consulta & "Where(ES.EST_PERMISO Like '%" & NoPermiso & "%' " & vbNewLine
                consulta = consulta & "    Or CL.CLI_CODNEGOCIO Like '%" & NoPermiso & "%' " & vbNewLine
                consulta = consulta & "    Or CL.CLI_NOPERMISO Like '%" & NoPermiso & "%' " & vbNewLine
                consulta = consulta & "    Or ES.EST_RAZON_SOCIAL Like '%" & NoPermiso & "%') " & vbNewLine
                consulta = consulta & "And " & vbNewLine
            Else
                consulta = consulta & "From " & dbfuelPriceComun.InitialCatalog.ToString & ".dbo.FP_Estacion ES " & vbNewLine
                consulta = consulta & "Where " & vbNewLine
            End If

            consulta = consulta & "COALESCE(ES.EST_RFC, '') <> '' " & vbNewLine &
                       "And Convert(VarChar(50), ES.EST_CODIGO) In(" & LstCli & ") " & vbNewLine &
                       "Order by ES.EST_RAZON_SOCIAL"

            result = da.Consulta(consulta, _fuelPriceComun)
        Catch ex As Exception
            result = New DataTable()
        End Try

        Return result
    End Function
    Public Function GetEstacionesByPermisoConCompetencia(ByVal NoPermiso, ByVal LstCli) As DataTable
        Dim result As DataTable

        Try
            consulta = "Select Distinct Convert(VarChar(50), ES.EST_PERMISO) As EST_PERMISO, ES.EST_RAZON_SOCIAL, Convert(VarChar(50), ES.EST_CODIGO) As EST_CODIGO " & vbNewLine
            If NoPermiso <> "" Then
                consulta = consulta & "From " & dbfuelPriceComun.InitialCatalog.ToString & ".dbo.FP_Estacion ES " & vbNewLine
                consulta = consulta & "Inner Join " & dbfuelPrice.InitialCatalog.ToString & ".dbo.FP_Cliente CL On Convert(VarChar(50), CL.CLI_EST_CODIGO) = Convert(VarChar(50), ES.EST_CODIGO) " & vbNewLine
                consulta = consulta & "Where(ES.EST_PERMISO Like '%" & NoPermiso & "%' " & vbNewLine
                consulta = consulta & "    Or CL.CLI_CODNEGOCIO Like '%" & NoPermiso & "%' " & vbNewLine
                consulta = consulta & "    Or CL.CLI_NOPERMISO Like '%" & NoPermiso & "%' " & vbNewLine
                consulta = consulta & "    Or ES.EST_RAZON_SOCIAL Like '%" & NoPermiso & "%') " & vbNewLine
                consulta = consulta & "And " & vbNewLine
            Else
                consulta = consulta & "From " & dbfuelPriceComun.InitialCatalog.ToString & ".dbo.FP_Estacion ES " & vbNewLine
                consulta = consulta & "Where " & vbNewLine
            End If

            consulta = consulta & "COALESCE(ES.EST_RFC, '') <> '' And Convert(VarChar(50), ES.EST_CODIGO) In(" & LstCli & ") " & vbNewLine

            consulta = consulta & "Union " & vbNewLine
            consulta = consulta & "Select Convert(VarChar(50), ES.NCO_PERMISO) As EST_PERMISO, ES.NCO_RAZONSOCIAL As EST_RAZON_SOCIAL, Convert(VarChar(50), ES.NCO_CODIGO) As EST_CODIGO " & vbNewLine
            consulta = consulta & "From " & dbfuelPrice.InitialCatalog.ToString & ".dbo.FP_Mercado_NegociosCompetencia ES " & vbNewLine
            consulta = consulta & "Where ES.NCO_CLI_CODIGO In (Select Distinct CL.CLI_CODIGO From " & dbfuelPrice.InitialCatalog.ToString & ".dbo.FP_Cliente CL Where Convert(VarChar(50), CL.CLI_EST_CODIGO) " & vbNewLine
            consulta = consulta & "In(" & LstCli & ")) " & vbNewLine
            consulta = consulta & "And (ES.NCO_PERMISO Like '%" & NoPermiso & "%' Or ES.NCO_CODIGO Like '%" & NoPermiso & "%' Or ES.NCO_RAZONSOCIAL Like '%" & NoPermiso & "%')" & vbNewLine

            consulta = consulta & "Order by EST_RAZON_SOCIAL"


            result = da.Consulta(consulta, _fuelPriceComun)
        Catch ex As Exception
            result = New DataTable()
        End Try

        Return result
    End Function
    Public Function GetEstacionesByCode(ByVal IdEst) As DataTable
        Dim result As New DataTable

        Try
            consulta = "Select EST_PERMISO, EST_RAZON_SOCIAL, EST_CODIGO From FP_Estacion Where EST_CODIGO = '" & IdEst & "' Order by EST_RAZON_SOCIAL"

            result = da.Consulta(consulta, _fuelPriceComun)
        Catch ex As Exception
            result = Nothing
        End Try

        Return result
    End Function
    Public Function GetProductosCre() As DataTable
        Dim result As New DataTable

        Try
            consulta = "Select CON_CODIGO_WEB, CON_DESCRIPCION, CON_CODIGO, CON_CODIGOTRANSFERENCIA " & vbNewLine &
                       "From FP_MAE_Combustible " & vbNewLine &
                       "Where CON_CODIGO In (Select Distinct CON_CODIGOTRANSFERENCIA From FP_MAE_Combustible)  And CON_ESBIO = 0 " & vbNewLine &
                       "Order By CON_CODIGO_WEB"

            result = da.Consulta(consulta, _fuelPriceComun)
        Catch ex As Exception
            result = Nothing
        End Try

        Return result
    End Function

    Public Function GetProductosSelection(ByVal IdMaestro, Optional ByVal EsBio = 0) As DataTable
        Dim result As New DataTable

        Try
            consulta = "Select MCL_REGPVP From FP_MAE_Clientes Where MCL_CODIGO = " & IdMaestro
            Dim tblName As String = da.Consulta(consulta, _fuelPriceComun).Rows(0)(0).ToString

            If tblName.ToUpper() = "FP_MAE_COMBUSTIBLE" Then
                consulta = "Select Convert(VarChar(50), MC.CON_CODIGO) AS CON_CODIGO, MC.CON_CODIGO_WEB, MC.CON_DESCRIPCION, MC.CON_DESCRIPCION_CORTA " & vbNewLine &
                        "From " & dbfuelPriceComun.InitialCatalog.ToString & ".dbo.FP_MAE_Combustible MC Inner Join " & dbfuelPriceComun.InitialCatalog.ToString & ".dbo.FP_ClientesProductos CP on CP.CPO_CON_CODIGO = MC.CON_CODIGO And CP.CPO_MCL_CODIGO = " & IdMaestro & " " & vbNewLine &
                        "Where MC.CON_ESBIO = " & EsBio & " " & vbNewLine &
                        "Order By CON_CODIGO_WEB"
            Else

                consulta = "Select Convert(VarChar(50), VTP_CODIGO) AS CON_CODIGO, VTP_CON_CODIGO, VTP_DESCRIPCION AS CON_DESCRIPCION, VTP_DESCRIPCION_CORTA As CON_DESCRIPCION_CORTA " & vbNewLine &
                           "From " & dbfuelPriceComun.InitialCatalog.ToString & ".dbo.FP_Vtas_Productos Where VTP_MCL_CODIGO = " & IdMaestro

            End If

            result = da.Consulta(consulta, _fuelPriceComun)
        Catch ex As Exception
            result = New DataTable
        End Try

        Return result
    End Function
    Public Function GetProductosSelectionFormat(ByVal IdMaestro, Optional ByVal EsBio = 0) As DataTable
        Dim result As New DataTable

        Try
            consulta = "Select MCL_REGPVP From FP_MAE_Clientes Where MCL_CODIGO = " & IdMaestro
            Dim tblName As String = da.Consulta(consulta, _fuelPriceComun).Rows(0)(0).ToString

            If tblName.ToUpper() = "FP_MAE_COMBUSTIBLE" Then
                consulta = "Select Convert(VarChar(50), MC.CON_CODIGO) As CON_CODIGO, Cast(MC.CON_CODIGO_WEB as VarChar) As VTP_CON_CODIGO, MC.CON_DESCRIPCION, MC.CON_DESCRIPCION_CORTA, Cast(MC.CON_CODIGO_WEB As VarChar) As VTP_CODIGO " & vbNewLine &
                        "From " & dbfuelPriceComun.InitialCatalog.ToString & ".dbo.FP_MAE_Combustible MC Inner Join " & dbfuelPriceComun.InitialCatalog.ToString & ".dbo.FP_ClientesProductos CP on CP.CPO_CON_CODIGO = MC.CON_CODIGO And CP.CPO_MCL_CODIGO = " & IdMaestro & " " & vbNewLine &
                        "Where MC.CON_ESBIO = " & EsBio & " " & vbNewLine &
                        "Order By CON_CODIGO_WEB"
            Else
                'consulta = "Select FORMAT(Convert(int, VTP_CON_CODIGO), '00000000-0000-0000-0000-000000000000') AS CON_CODIGO, VTP_CON_CODIGO, VTP_DESCRIPCION AS CON_DESCRIPCION, VTP_DESCRIPCION_CORTA As CON_DESCRIPCION_CORTA " & vbNewLine &
                '           "From " & dbfuelPriceComun.InitialCatalog.ToString & ".dbo.FP_Vtas_Productos Where VTP_MCL_CODIGO = " & IdMaestro

                'consulta = "Select FORMAT(Convert(int, VTP_CON_CODIGO), '00000000-0000-0000-0000-000000000000') AS CON_CODIGO, VTP_CON_CODIGO, VTP_DESCRIPCION AS CON_DESCRIPCION, VTP_DESCRIPCION_CORTA As CON_DESCRIPCION_CORTA, Cast(VTP_CODIGO As VarChar) As VTP_CODIGO " & vbNewLine &
                '           "From " & dbfuelPriceComun.InitialCatalog.ToString & ".dbo.FP_Vtas_Productos Where VTP_MCL_CODIGO = " & IdMaestro

                consulta = "Select " & vbNewLine &
                            "Case When TRY_CAST(VTP_CON_CODIGO As uniqueidentifier) Is Null Then " & vbNewLine &
                            "	Format(Convert(Int, Replace(VTP_CON_CODIGO, '-', '')), '00000000-0000-0000-0000-000000000000') " & vbNewLine &
                            "Else " & vbNewLine &
                            "	Cast(VTP_CON_CODIGO As uniqueidentifier) " & vbNewLine &
                            "End As CON_CODIGO, VTP_CON_CODIGO, VTP_DESCRIPCION As CON_DESCRIPCION, " & vbNewLine &
                            "VTP_DESCRIPCION_CORTA As CON_DESCRIPCION_CORTA, Cast(VTP_CODIGO As VarChar) As VTP_CODIGO " & vbNewLine &
                            "From " & dbfuelPriceComun.InitialCatalog.ToString & ".dbo.FP_Vtas_Productos Where VTP_MCL_CODIGO = " & IdMaestro

            End If

            result = da.Consulta(consulta, _fuelPriceComun)
        Catch ex As Exception
            result = New DataTable
        End Try

        Return result
    End Function
End Class
