Imports System.IO
Imports System.Data
Imports FuelPrice.DataAccess
Imports ExcelDataReader

Public Class FMBussinesUtils

    Dim bl As FMBussinesLayer = New FMBussinesLayer()
    Private da As New FMDataAccess
    Private _fuelPrice As String = "FuelPrice"
    Private _fuelPriceComun As String = "FuelPriceComun"
    Public Function CreateHistorico(ByRef tblHist, ByVal Cliente, ByVal Producto, ByVal Titulo) As String
        Dim result As String = ""
        Dim nomProd As String
        Dim oddColor As String = "#FFFFFF"
        Dim evenColor As String = "#F2F2F2"
        Dim rowColor As String = evenColor
        Dim tblC As DataTable = tblHist

        result += "<html>" & vbNewLine
        result += "<body>" & vbNewLine
        result += "<table border=""0""><tr><td colspan=""3""><strong>Hist&oacute;rico de " & Titulo & "</strong></td></tr><tr><td colspan=""3""><strong>Distribuidor: " & Cliente & "</strong></td></tr><tr><td colspan=""3""><strong>Producto: " & Producto & "</strong></td></tr></table>" & vbNewLine
        result += "<br /><br />" & vbNewLine
        result += "<div style=""width: " + (tblC.Columns.Count * 128).ToString + "px;"">" & vbNewLine
        result += "<table style=""width:100%; border: 1 solid #D3D3D3;"">" & vbNewLine
        result += "<tr>" & vbNewLine
        For index = 0 To tblC.Columns.Count - 1 Step 2
            Dim valNof() As String = tblC.Columns(index).ColumnName.Split("_")
            nomProd = bl.GetCostosDirectosById(valNof(1)).Rows(0)(1)
            result += "<td style=""background-color: #ffffff;border: 1 solid gray;"" colspan=""2"">" & nomProd & "</td>" & vbNewLine
        Next
        result += "</tr>" & vbNewLine
        result += "<tr>" & vbNewLine
        For index = 0 To tblC.Columns.Count - 1 Step 2
            result += "<td style=""background-color: #5f97b3;color: white;border: 1 solid gray;"">Valor</td><td style=""background-color: #5f97b3;color: white;border: 1 solid gray;"">Vigencia</td>" & vbNewLine
        Next
        result += "</tr>" & vbNewLine
        result += "<tbody>"
        For index = 0 To tblC.Rows.Count - 1
            result += "<tr>"
            For indexC = 0 To tblC.Columns.Count - 1
                result += "<td style=""background-color: " & rowColor & ";color: black;border: 1 solid gray;"">" & If(IsDBNull(tblC.Rows(index)(indexC)), "&nbsp;", tblC.Rows(index)(indexC).ToString) & "</td>"
            Next
            result += "</tr>" & vbNewLine
            If rowColor = oddColor Then
                rowColor = evenColor
            Else
                rowColor = oddColor
            End If
        Next
        result += "</tbody>"
        result += "</table>" & vbNewLine
        result += "</div>" & vbNewLine
        result += "</body>" & vbNewLine
        result += "</html>" & vbNewLine

        Return result
    End Function
    Public Function ExcelToDataTable(ByVal byteArray) As DataTable
        Dim stream As New MemoryStream(CType(byteArray, Byte()))
        Dim reader = ExcelDataReader.ExcelReaderFactory.CreateReader(stream)
        Dim conf = New ExcelDataSetConfiguration With {
            .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration With {
                .UseHeaderRow = True
            }
        }
        Dim dataSet = reader.AsDataSet(conf)
        Dim dataTable = dataSet.Tables(0)
        Return dataTable
    End Function
    Public Function ImportHistorico(ByVal byteFile, ByVal TypeFile, ByVal IdMaestro) As DataTable 'Costos Directos e Indirectos
        Dim result As New DataTable()
        result.Columns.Add("Informe")

        Dim CliCodigo As String = ""
        Dim cteName As String
        Dim prodName As String
        Dim costoId As String
        Dim valor As String
        Dim vigencia As String

        Dim cteValido As Boolean
        Dim prodValido As Boolean
        Dim costoValido As Boolean
        Dim fechaValida As Boolean
        Dim vigenciaValida As Boolean
        Dim costeExistente As Boolean
        Dim strProdId As String

        Dim decPrecio As Decimal
        Dim validPrecio As Boolean
        Dim tblSource As String

        Dim rowMessage As String
        Dim strConsultaVigencia As String
        Dim strConsultaExistente As String

        Dim sepVal As Tuple(Of String, String)

        Dim _dt As DataTable = ExcelToDataTable(byteFile)
        If _dt.Columns.Count <> 5 Then
            Dim _r As DataRow = result.NewRow
            _r(0) = "Número incorrecto de columnas en el archivo"
            result.Rows.Add(_r)
        Else
            For index = 0 To _dt.Rows.Count - 1
                cteName = _dt.Rows(index)(0).ToString
                prodName = _dt.Rows(index)(1).ToString
                costoId = _dt.Rows(index)(2).ToString
                valor = _dt.Rows(index)(3).ToString
                vigencia = _dt.Rows(index)(4).ToString

                'Valida existencia del cliente
                Dim dtCte As DataTable = bl.GetClientById(cteName)
                If IsNothing(dtCte) Then
                    cteValido = False
                ElseIf bl.GetClientById(cteName).Rows.Count <= 0 Then
                    cteValido = False
                Else
                    CliCodigo = dtCte.Rows(0)("CLI_CODIGO").ToString()
                    cteValido = True
                End If

                If cteValido = False Then
                    dtCte = bl.GetNegociosPorPermiso(cteName)
                    If IsNothing(dtCte) Then
                        cteValido = False
                    ElseIf bl.GetNegociosPorPermiso(cteName).Rows.Count <= 0 Then
                        cteValido = False
                    Else
                        CliCodigo = dtCte.Rows(0)("CLI_CODIGO").ToString()
                        cteValido = True
                    End If
                End If

                If cteValido = False Then
                    dtCte = bl.GetClientByCodNegocio(cteName, IdMaestro)
                    If IsNothing(dtCte) Then
                        cteValido = False
                    ElseIf bl.GetClientByCodNegocio(cteName, IdMaestro).Rows.Count <= 0 Then
                        cteValido = False
                    Else
                        CliCodigo = dtCte.Rows(0)("CLI_CODIGO").ToString()
                        cteValido = True
                    End If
                End If

                'Valida existencia del producto
                'Dim dtProd As DataTable = bl.GetProductoByShortName(prodName)
                Dim dtProd As DataTable = bl.GetProductosMaeVtp(IdMaestro)
                Dim elProd As DataRow() = dtProd.Select("CON_DESCRIPCION_CORTA = '" & prodName & "'")
                If elProd.Length > 0 Then
                    strProdId = elProd(0)("CON_CODIGO").ToString
                    prodValido = True
                Else
                    prodValido = False
                End If

                'Valida existencia del Costo
                Select Case TypeFile
                    Case 1, 2 ' Directos
                        If IsNothing(bl.GetCostosDirectosById(costoId)) Then
                            costoValido = False
                        ElseIf bl.GetCostosDirectosById(costoId).Rows.Count <= 0 Then
                            costoValido = False
                        Else
                            costoValido = True
                        End If
                    Case 3, 4 ' Indirectos
                        If IsNothing(bl.GetNoFacturableById(costoId)) Then
                            costoValido = False
                        ElseIf bl.GetNoFacturableById(costoId).Rows.Count <= 0 Then
                            costoValido = False
                        Else
                            costoValido = True
                        End If
                End Select

                rowMessage = "Error en registro " & (index + 2).ToString & ": "

                'Valida que la fecha que tiene el archivo exista
                fechaValida = ValidateDate(vigencia)

                'Valida Precio
                If valor <> "" Then
                    sepVal = SeparatedDecimal(valor)
                    valor = sepVal.Item1 & "." & sepVal.Item2

                    Dim tupDec As Tuple(Of Decimal, Boolean) = ValidaValorDecimal(valor)
                    decPrecio = tupDec.Item1
                    validPrecio = tupDec.Item2
                    If validPrecio = False Then
                        rowMessage += "Valor incorrecto, "
                    End If
                Else
                    validPrecio = True
                    decPrecio = 0
                End If



                'Verifica la existencia de la combinacion en la tabla padre
                Select Case TypeFile
                    Case 1, 2
                        strConsultaExistente = "Select Count(*) From FP_CostosDirectosNegocio Where CDN_CDI_CODIGO='" & costoId & "' And CDN_CLI_CODIGO='" & CliCodigo & "' And CDN_CON_CODIGO='" & strProdId & "'"
                    Case 3, 4
                        strConsultaExistente = "Select Count(*) From FP_CostosIndirectosNegocio Where COI_CIN_CODIGO='" & costoId & "' And COI_CLI_CODIGO='" & CliCodigo & "' And COI_CON_CODIGO='" & strProdId & "'"
                End Select

                Dim _exist As DataTable = bl.GetCosteExistente(strConsultaExistente)
                If _exist IsNot Nothing Then
                    If _exist.Rows(0)(0) > 0 Then
                        'Busca la última vigencia
                        If cteValido = True And prodValido = True And costoValido = True And fechaValida = True Then
                            Select Case TypeFile
                                Case 1
                                    strConsultaVigencia = "Select Top 1 CDN_VIGENCIA From FP_CostosDirectosNegocio Where CDN_CDI_CODIGO=" & costoId & " And CDN_CLI_CODIGO='" & CliCodigo & "' And CDN_CON_CODIGO='" & strProdId & "' Order By CDN_VIGENCIA Desc"
                                Case 2
                                    strConsultaVigencia = "Select Top 1 CDD_VIGENCIADESDE From FP_CostosDirectosDistribuidor Where CDD_CDI_CODIGO=" & costoId & " And CDD_CLI_CODIGO='" & CliCodigo & "' And CDD_CON_CODIGO='" & strProdId & "' Order By CDD_VIGENCIADESDE Desc"
                                Case 3
                                    strConsultaVigencia = "Select Top 1 COI_Vigencia From FP_CostosIndirectosNegocio Where COI_CIN_CODIGO=" & costoId & " And COI_CLI_CODIGO='" & CliCodigo & "' And COI_CON_CODIGO='" & strProdId & "' Order By COI_VIGENCIA Desc"
                                Case 4
                                    strConsultaVigencia = "Select Top 1 CID_VIGENCIADESDE From FP_CostosIndirectosDistribuidor Where CID_CCI_CODIGO=" & costoId & " And CID_MCL_CODIGO='" & CliCodigo & "' And CID_CON_CODIGO='" & strProdId & "' Order By CID_VIGENCIADESDE Desc"
                            End Select

                            Dim fch As DateTime = DateTime.Parse(vigencia)
                            Dim dtFecha As DateTime = bl.GetLastFechaCostoDynamic(strConsultaVigencia, vigencia)

                            Dim _diff As Integer = DateDiff(DateInterval.Day, dtFecha, fch)

                            'Valida la última vigencia
                            If (_diff < 0) Then
                                vigenciaValida = False
                            Else
                                vigenciaValida = True
                            End If
                            costeExistente = True
                        End If
                    Else
                        costeExistente = False
                    End If
                End If

                If costeExistente = False Then
                    rowMessage += "Configuración de COSTE no existente, "
                End If
                If cteValido = False Then
                    rowMessage += "Negocio no existente, "
                End If
                If prodValido = False Then
                    rowMessage += "Producto no existente, "
                End If
                If costoValido = False Then
                    rowMessage += "Costo no existente, "
                End If
                If vigenciaValida = False Then
                    rowMessage += "La Fecha es inválida o menor a la última registrada"
                End If

                If cteValido = True And prodValido = True And costoValido = True And vigenciaValida = True And costeExistente = True And validPrecio = True Then
                    Dim intId As Integer = 0
                    Integer.TryParse(strProdId, intId)

                    If intId = 0 Then
                        tblSource = "FP_MAE_Combustible"
                    Else
                        tblSource = "FP_Vtas_Productos"
                    End If

                    Dim fchInsert As DateTime = DateTime.Parse(vigencia)
                    Dim _ins As Integer = bl.InserHistoricoMasivo(TypeFile, CliCodigo, costoId, strProdId, decPrecio, fchInsert.ToString("yyyy-MM-dd") & " " & DateTime.Now.ToString("HH:mm:ss"), tblSource)

                    Dim dr As DataRow = result.NewRow
                    dr(0) = "Registro " & (index + 2).ToString & " insertado con éxito"
                    result.Rows.Add(dr)
                Else
                    If Mid(rowMessage, rowMessage.Length - 2, 2) = ", " Then
                        rowMessage = Mid(rowMessage, 0, rowMessage.Length - 2)
                    End If
                    Dim dr = result.NewRow
                    dr(0) = rowMessage
                    result.Rows.Add(dr)
                End If
            Next
        End If
        Return result
    End Function
    Public Function ImportPreciosReferencia(ByVal byteFile, Optional ByVal EsBio = 0) As DataTable
        Dim result As DataTable = New DataTable
        result.Columns.Add("Informe")

        Dim codTerm As String = ""
        Dim codProd As String = ""
        Dim codPrecio As String = ""
        Dim codVig As String = ""
        Dim strProdId As String = ""
        Dim decPrecio As Decimal

        Dim validTerm As Boolean
        Dim validProd As Boolean
        Dim validFlete As Boolean
        Dim validVig As Boolean

        Dim rowMessage As String = ""

        Dim _dt As DataTable = ExcelToDataTable(byteFile)
        If _dt.Columns.Count <> 4 Then
            Dim _r As DataRow = result.NewRow
            _r(0) = "Número incorrecto de columnas en el archivo"
            result.Rows.Add(_r)
        Else
            For index = 0 To _dt.Rows.Count - 1
                'CÓDIGO TERMINAL	CÓDIGO PRODUCTO	DESCUENTO REF	VIGENCIA DESDE
                codTerm = _dt.Rows(index)(0).ToString
                codProd = _dt.Rows(index)(1).ToString
                codPrecio = _dt.Rows(index)(2).ToString
                codVig = _dt.Rows(index)(3).ToString

                rowMessage = "Error en registro " & (index + 2).ToString & ": "

                If codTerm = "" Or codProd = "" Or codPrecio = "" Or codVig = "" Then
                    rowMessage += "Valores vacíos en el registro"
                End If

                If rowMessage.EndsWith("Valores vacíos en el registro") Then
                    Dim dr As DataRow = result.NewRow
                    dr(0) = rowMessage
                    result.Rows.Add(dr)
                Else
                    ' Valida existencia del Tar
                    Dim _tblTar As DataTable = bl.GetTarByTermCode(codTerm)
                    If IsNothing(_tblTar) Then
                        validTerm = False
                        rowMessage += "Terminal no econtrada, "
                    ElseIf _tblTar.Rows.Count <= 0 Then
                        validTerm = False
                        rowMessage += "Terminal no econtrada, "
                    Else
                        codTerm = _tblTar.Rows(0)("TSU_CODIGO").ToString
                        validTerm = True
                    End If

                    'Valida existencia del producto
                    Dim dtProd As DataTable = bl.GetProductoByShortName(codProd)
                    If IsNothing(dtProd) Then
                        validProd = False
                        rowMessage += "Producto no econtrado, "
                    ElseIf dtProd.Rows.Count <= 0 Then
                        validProd = False
                        rowMessage += "Producto no econtrado, "
                    Else
                        If Convert.ToInt32(dtProd(0)(2)) = EsBio Then
                            strProdId = dtProd.Rows(0)(0).ToString
                            validProd = True
                        Else
                            validProd = False
                            rowMessage += "Producto no pertenece a este apartado, "
                        End If
                    End If

                    'Valida la fecha
                    validVig = ValidateDate(codVig)

                    'Valida precio
                    If codPrecio <> "" Then
                        Dim tupDec As Tuple(Of Decimal, Boolean) = ValidaValorDecimal(codPrecio)
                        decPrecio = tupDec.Item1
                        validFlete = tupDec.Item2
                        If validFlete = False Then
                            rowMessage += "Valor de precio incorrecto, "
                        End If
                    Else
                        validFlete = False
                        rowMessage += "Valor de precio incorrecto, "
                    End If

                    If validProd And validFlete And validTerm And validVig Then

                        Dim sepDec As Tuple(Of String, String) = SeparatedDecimal(decPrecio)

                        Dim consulta As String = "Select PRF_CODIGO From FP_PreciosReferenciaTerminal " & vbNewLine &
                            "Where PRF_TSU_CODIGO = '" & codTerm & "' " & vbNewLine &
                            " And PRF_CON_CODIGO = '" & strProdId & "' " & vbNewLine &
                            "And PRF_FECHA = '" & ValidaValorFecha(codVig).Item1.ToString("yyyy-MM-dd") & "'"
                        Dim _tbl = da.Consulta(consulta, _fuelPrice)

                        If IsNothing(_tbl) Then
                            Dim _inserta As String = "INSERT INTO dbo.FP_PreciosReferenciaTerminal " &
                                                 "     (PRF_TSU_CODIGO " &
                                                 "     ,PRF_CON_CODIGO " &
                                                 "     ,PRF_PRECIO " &
                                                 "     ,PRF_FECHA " &
                                                 "     ,PRF_HORAAPLICACION) " &
                                                 " OUTPUT INSERTED.PRF_CODIGO " &
                                                 " VALUES " &
                                                 "     ('" & codTerm & "' " &
                                                 "     ,'" & strProdId & "' " &
                                                 "     ,'" & sepDec.Item1 & "." & sepDec.Item2 & "' " &
                                                 "     ,'" & ValidaValorFecha(codVig).Item1.ToString("yyyy-MM-dd") & "' " &
                                                 "     ,'00:00:00.000')"

                            Dim _dtIns As DataTable = da.Consulta(_inserta, _fuelPrice)

                            If IsNothing(_dtIns) Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar"
                                result.Rows.Add(dr)
                            ElseIf Convert.ToInt32(_dtIns.Rows(0)(0)) <= 0 Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar"
                                result.Rows.Add(dr)
                            Else
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Registro " & (index + 2).ToString & " insertado con éxito"
                                result.Rows.Add(dr)
                            End If
                        ElseIf _tbl.Rows.Count <= 0 Then
                            Dim _inserta As String = "INSERT INTO dbo.FP_PreciosReferenciaTerminal " &
                                                 "     (PRF_TSU_CODIGO " &
                                                 "     ,PRF_CON_CODIGO " &
                                                 "     ,PRF_PRECIO " &
                                                 "     ,PRF_FECHA " &
                                                 "     ,PRF_HORAAPLICACION) " &
                                                 " OUTPUT INSERTED.PRF_CODIGO " &
                                                 " VALUES " &
                                                 "     ('" & codTerm & "' " &
                                                 "     ,'" & strProdId & "' " &
                                                 "     ,'" & sepDec.Item1 & "." & sepDec.Item2 & "' " &
                                                 "     ,'" & ValidaValorFecha(codVig).Item1.ToString("yyyy-MM-dd") & "' " &
                                                 "     ,'00:00:00.000')"

                            Dim _dtIns As DataTable = da.Consulta(_inserta, _fuelPrice)

                            If IsNothing(_dtIns) Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar"
                                result.Rows.Add(dr)
                            ElseIf Convert.ToInt32(_dtIns.Rows(0)(0)) <= 0 Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar"
                                result.Rows.Add(dr)
                            Else
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Registro " & (index + 2).ToString & " insertado con éxito"
                                result.Rows.Add(dr)
                            End If
                        Else
                            Dim actualiza As String = "UPDATE FP_PreciosReferenciaTerminal " & vbNewLine &
                               "SET PRF_PRECIO = '" & sepDec.Item1 & "." & sepDec.Item2 & "'" &
                               "WHERE PRF_CODIGO = " & _tbl.Rows(0)(0).ToString

                            Dim _dtUpd As Integer = da.NoQuery(actualiza, _fuelPrice)

                            If _dtUpd <= 0 Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible actualizar"
                                result.Rows.Add(dr)
                            Else
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Registro " & (index + 2).ToString & " actualizado con éxito"
                                result.Rows.Add(dr)
                            End If
                        End If
                    Else
                        If Mid(rowMessage, rowMessage.Length - 2, 2) = ", " Then
                            rowMessage = Mid(rowMessage, 0, rowMessage.Length - 2)
                        End If
                        Dim dr As DataRow = result.NewRow
                        dr(0) = rowMessage
                        result.Rows.Add(dr)
                    End If

                End If
            Next
        End If

        Return result
    End Function
    Public Function ImportVolVtaRealObj(ByVal byteFile, ByVal IdMaestro) As DataTable
        Dim result As DataTable = New DataTable
        result.Columns.Add("Informe")

        Dim codNeg As String = ""
        Dim codProd As String = ""
        Dim codReal As String = ""
        Dim codObj As String = ""
        Dim codFecha As String = ""
        Dim strProdId As String = ""
        Dim decReal As Decimal
        Dim decObj As Decimal

        Dim validNeg As Boolean
        Dim validProd As Boolean
        Dim validReal As Boolean
        Dim validObj As Boolean
        Dim validVig As Boolean

        Dim rowMessage As String = ""

        Dim _dt As DataTable = ExcelToDataTable(byteFile)
        If _dt.Columns.Count <> 5 Then
            Dim _r As DataRow = result.NewRow
            _r(0) = "Número incorrecto de columnas en el archivo"
            result.Rows.Add(_r)
        Else
            For index = 0 To _dt.Rows.Count - 1
                'CÓDIGO TERMINAL	CÓDIGO PRODUCTO	DESCUENTO REF	VIGENCIA DESDE
                codNeg = _dt.Rows(index)(0).ToString
                codProd = _dt.Rows(index)(1).ToString
                codFecha = _dt.Rows(index)(2).ToString
                codReal = _dt.Rows(index)(3).ToString
                codObj = _dt.Rows(index)(4).ToString

                rowMessage = "Error en registro " & (index + 2).ToString & ": "

                If codNeg = "" Or codProd = "" Or codReal = "" Or codFecha = "" Or codObj = "" Then
                    rowMessage += "Valores vacíos en el registro"
                End If

                If rowMessage.EndsWith("Valores vacíos en el registro") Then
                    Dim dr As DataRow = result.NewRow
                    dr(0) = rowMessage
                    result.Rows.Add(dr)
                Else
                    ' Valida existencia del Negocio
                    Dim dtCte As DataTable = bl.GetClientById(codNeg)
                    If IsNothing(dtCte) Then
                        validNeg = False
                        rowMessage += "Cliente no econtrado, "
                    ElseIf dtCte.Rows.Count <= 0 Then
                        validNeg = False
                        rowMessage += "Cliente no econtrado, "
                    Else
                        codNeg = dtCte.Rows(0)("CLI_CODIGO").ToString
                        validNeg = True
                    End If

                    If validNeg = False Then
                        dtCte = bl.GetNegociosPorPermiso(codNeg)
                        If IsNothing(dtCte) Then
                            validNeg = False
                        ElseIf bl.GetNegociosPorPermiso(codNeg).Rows.Count <= 0 Then
                            validNeg = False
                        Else
                            codNeg = dtCte.Rows(0)("CLI_CODIGO").ToString()
                            validNeg = True
                        End If
                    End If

                    If validNeg = False Then
                        dtCte = bl.GetClientByCodNegocio(codNeg, IdMaestro)
                        If IsNothing(dtCte) Then
                            validNeg = False
                        ElseIf bl.GetClientByCodNegocio(codNeg, IdMaestro).Rows.Count <= 0 Then
                            validNeg = False
                        Else
                            codNeg = dtCte.Rows(0)("CLI_CODIGO").ToString()
                            validNeg = True
                        End If
                    End If

                    ''Valida existencia del producto
                    'Dim dtProd As DataTable = bl.GetProductoByShortName(codProd)
                    'If IsNothing(dtProd) Then
                    '    validProd = False
                    '    rowMessage += "Producto no econtrado, "
                    'ElseIf dtProd.Rows.Count <= 0 Then
                    '    validProd = False
                    '    rowMessage += "Producto no econtrado, "
                    'Else
                    '    codProd = dtProd.Rows(0)(0).ToString
                    '    validProd = True
                    'End If

                    'Valida existencia del producto
                    'Dim dtProd As DataTable = bl.GetProductoByShortName(prodName)
                    Dim dtProd As DataTable = bl.GetProductosMaeVtp(IdMaestro)
                    Dim elProd As DataRow() = dtProd.Select("CON_DESCRIPCION_CORTA = '" & codProd & "'")
                    If elProd.Length > 0 Then
                        codProd = elProd(0)("CON_CODIGO").ToString
                        validProd = True
                    Else
                        validProd = False
                        rowMessage += "Producto no econtrado, "
                    End If

                    'Valida la fecha
                    validVig = ValidateDate(codFecha)

                    'Valida Real
                    If codReal <> "" Then
                        If codReal = "-" Then
                            validReal = True
                        Else
                            Dim tupDec As Tuple(Of Decimal, Boolean) = ValidaValorDecimal(codReal)
                            decReal = tupDec.Item1
                            validReal = tupDec.Item2
                            If validReal = False Then
                                rowMessage += "Valor Real incorrecto, "
                            End If
                        End If
                    Else
                        validReal = False
                        rowMessage += "Valor Real incorrecto, "
                    End If

                    'Valida Objetivo
                    If codObj <> "" Then
                        If codObj = "-" Then
                            validObj = True
                        Else
                            Dim tupDec As Tuple(Of Decimal, Boolean) = ValidaValorDecimal(codObj)
                            decObj = tupDec.Item1
                            validObj = tupDec.Item2
                            If validObj = False Then
                                rowMessage += "Valor Objetivo incorrecto, "
                            End If
                        End If
                    Else
                        validObj = False
                        rowMessage += "Valor Objetivo incorrecto, "
                    End If

                    If validProd And validReal And validNeg And validVig And validObj Then

                        Dim consulta = "SELECT VRO_CODIGO FROM FP_VolVta_Real_Objetivo " & vbNewLine &
                           "Where VRO_CLI_CODIGO = '" & codNeg & "' " & vbNewLine &
                           "And VRO_CON_CODIGO = '" & codProd & "' " & vbNewLine &
                           "And VRO_FECHA = '" & Convert.ToDateTime(codFecha).ToString("yyyy-MM-dd") & "'"

                        Dim _tbl = da.Consulta(consulta, _fuelPrice)

                        Dim sepDec As Tuple(Of String, String) = SeparatedDecimal(decObj)

                        If IsNothing(_tbl) Then
                            Dim _inserta = "INSERT INTO FP_VolVta_Real_Objetivo
                                       (VRO_CLI_CODIGO
                                       ,VRO_CON_CODIGO
                                       ,VRO_VOL_REAL
                                       ,VRO_VOL_OBJETIVO
                                       ,VRO_FECHA)
                                 OUTPUT INSERTED.VRO_CODIGO 
                                 VALUES
                                       ('" & codNeg & "','" & codProd & "'
                                       ,'" & IIf(codReal = "-", "0", decReal) & "',
                                       '" & IIf(codObj = "-", "0", sepDec.Item1 & "." & sepDec.Item2) & "',
                                       '" & DateTime.Parse(codFecha).ToString("yyyy-MM-dd") & "')"

                            Dim _dtIns As DataTable = da.Consulta(_inserta, _fuelPrice)

                            If IsNothing(_dtIns) Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar"
                                result.Rows.Add(dr)
                            ElseIf Convert.ToInt32(_dtIns.Rows(0)(0)) <= 0 Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar"
                                result.Rows.Add(dr)
                            Else
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Registro " & (index + 2).ToString & " insertado con éxito"
                                result.Rows.Add(dr)
                            End If
                        ElseIf _tbl.Rows.Count <= 0 Then
                            Dim _inserta = "INSERT INTO FP_VolVta_Real_Objetivo
                                       (VRO_CLI_CODIGO
                                       ,VRO_CON_CODIGO
                                       ,VRO_VOL_REAL
                                       ,VRO_VOL_OBJETIVO
                                       ,VRO_FECHA)
                                 OUTPUT INSERTED.VRO_CODIGO 
                                 VALUES
                                       ('" & codNeg & "','" & codProd & "'
                                       ,'" & IIf(codReal = "-", "0", decReal.ToString) & "',
                                       '" & IIf(codObj = "-", "0", sepDec.Item1 & "." & sepDec.Item2) & "',
                                       '" & DateTime.Parse(codFecha).ToString("yyyy-MM-dd") & "')"

                            Dim _dtIns As DataTable = da.Consulta(_inserta, _fuelPrice)

                            If IsNothing(_dtIns) Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar"
                                result.Rows.Add(dr)
                            ElseIf Convert.ToInt32(_dtIns.Rows(0)(0)) <= 0 Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar"
                                result.Rows.Add(dr)
                            Else
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Registro " & (index + 2).ToString & " insertado con éxito"
                                result.Rows.Add(dr)
                            End If
                        Else
                            consulta = "Update FP_VolVta_Real_Objetivo Set" & vbNewLine

                            If codReal <> "-" Then
                                consulta = consulta & "VRO_VOL_REAL = '" & decReal.ToString & "'," & vbNewLine
                            End If

                            If codObj <> "-" Then
                                consulta = consulta & "VRO_VOL_OBJETIVO = '" & sepDec.Item1 & "." & sepDec.Item2 & "'" & vbNewLine
                            End If

                            If consulta.Trim().EndsWith(",") Then
                                consulta = consulta.Trim()
                                consulta = Mid(consulta, 1, Len(consulta) - 1) & vbNewLine
                            End If

                            consulta = consulta & "Where VRO_CODIGO = '" & _tbl.Rows(0)(0).ToString & "'"

                            Dim _dtUpd As Integer = da.NoQuery(consulta, _fuelPrice)

                            If _dtUpd <= 0 Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible actualizar"
                                result.Rows.Add(dr)
                            Else
                                Dim dr As DataRow = result.NewRow
                                dr(0) = "Registro " & (index + 2).ToString & " actualizado con éxito"
                                result.Rows.Add(dr)
                            End If
                        End If
                    Else
                        If Mid(rowMessage, rowMessage.Length - 2, 2) = ", " Then
                            rowMessage = Mid(rowMessage, 0, rowMessage.Length - 2)
                        End If
                        Dim dr As DataRow = result.NewRow
                        dr(0) = rowMessage
                        result.Rows.Add(dr)
                    End If

                End If
            Next
        End If

        Return result
    End Function
    Public Function ImportPVP(ByVal byteFile, ByVal IdMaestro, ByVal Cultura) As DataTable
        Dim result As DataTable = New DataTable
        result.Columns.Add("Informe")

        Dim tr As New FMBTraductor("RegistroPVP")

        Dim codNeg As String = ""
        Dim codEst As String = ""
        Dim codPermiso As String = ""
        Dim codProd As String = ""
        Dim codPrecio As String = ""
        Dim codHora As String = ""
        Dim codFecha As String = ""
        Dim strProdId As String = ""
        Dim decPrecio As Decimal

        Dim validNeg As Boolean
        Dim validProd As Boolean
        Dim validPrecio As Boolean
        Dim validFecha As Boolean
        Dim validHora As Boolean

        Dim rowMessage As String = ""
        Dim sepDec As Tuple(Of String, String)

        Dim _dt As DataTable = ExcelToDataTable(byteFile)
        If _dt.Columns.Count <> 5 Then
            Dim _r As DataRow = result.NewRow
            _r(0) = tr.Traduce("_NUMEROINCORRECTOCOLUMNAS_", Cultura) '"Número incorrecto de columnas en el archivo"
            result.Rows.Add(_r)
        Else
            Dim cteNoEnt As String = tr.Traduce("_CLIENTENOENCONTRADO_", Cultura) '"Cliente no econtrado"
            Dim prodNoEnc As String = tr.Traduce("_PRODUCTONOENCONTRADO_", Cultura) '"Producto no econtrado"
            Dim fechIncorrecto As String = tr.Traduce("_FECHAINCORRECTO_", Cultura) '"Valor Fecha incorrecto"
            Dim horaIncorrecto As String = tr.Traduce("_HORAINCORRECTO_", Cultura) '"Valor Hora incorrecto"
            Dim precioIncorrecto As String = tr.Traduce("_PRECIOINCORRECTO_", Cultura) '"Valor Precio incorrecto"
            Dim volumenIncorrecto As String = tr.Traduce("_VOLUMENINCORRECTO_", Cultura) '"Valor Volumen incorrecto"
            Dim errorRegistro As String = tr.Traduce("_ERRORREGISTRO_", Cultura) '"Error en registro "
            Dim imposibleInsertar As String = tr.Traduce("_IMPOSIBLEINSERTAR_", Cultura) '"Imposible insertar"
            Dim registro As String = tr.Traduce("_REGISTRO_", Cultura) '"Registro"
            Dim insertadoExito As String = tr.Traduce("_INSERTADOEXITO_", Cultura) 'insertado con éxito"
            Dim imposibleActualizar As String = tr.Traduce("_IMPOSIBLEACTUALIZAR_", Cultura) 'Imposible actualizar
            Dim actualizadoExito As String = tr.Traduce("_ACTUALIZADOEXITO_", Cultura) 'actualizado con éxito
            Dim horaMayor As String = tr.Traduce("_HORAMAYOR_", Cultura) 'Ya existe una hora mayor

            For index = 0 To _dt.Rows.Count - 1
                codNeg = _dt.Rows(index)(0).ToString
                codProd = _dt.Rows(index)(1).ToString
                codPrecio = _dt.Rows(index)(2).ToString
                codFecha = _dt.Rows(index)(3).ToString
                codHora = _dt.Rows(index)(4).ToString

                rowMessage = errorRegistro & (index + 2).ToString & ": "

                If codNeg = "" Or codProd = "" Or codPrecio = "" Or codFecha = "" Or codHora = "" Then
                    rowMessage += tr.Traduce("_VALORESVACIOS_", Cultura) '"Valores vacíos en el registro"
                End If

                If rowMessage.EndsWith(tr.Traduce("_VALORESVACIOS_", Cultura)) Then
                    Dim dr As DataRow = result.NewRow
                    dr(0) = rowMessage
                    result.Rows.Add(dr)
                Else
                    ' Valida existencia del Negocio
                    Dim dtCte As DataTable = bl.GetEstacionesByPermiso(codNeg)
                    If IsNothing(dtCte) Then
                        validNeg = False
                    ElseIf dtCte.Rows.Count <= 0 Then
                        validNeg = False
                    Else
                        codEst = dtCte.Rows(0)("EST_CODIGO").ToString
                        codPermiso = dtCte.Rows(0)("EST_PERMISO").ToString
                        validNeg = True
                    End If

                    If validNeg = False Then
                        dtCte = bl.GetClientByCodNegocio(codNeg, IdMaestro)
                        If IsNothing(dtCte) Then
                            validNeg = False
                            rowMessage += cteNoEnt & ", "
                        ElseIf bl.GetClientByCodNegocio(codNeg, IdMaestro).Rows.Count <= 0 Then
                            validNeg = False
                            rowMessage += cteNoEnt & ", "
                        Else
                            codEst = dtCte.Rows(0)("EST_CODIGO").ToString
                            codPermiso = dtCte.Rows(0)("EST_PERMISO").ToString
                            validNeg = True
                        End If
                    End If

                    'Valida existencia del producto
                    Dim dtProd As DataTable = bl.GetProductosSelectionFormat(IdMaestro)
                    Dim elProd As DataRow() = dtProd.Select("CON_DESCRIPCION_CORTA = '" & codProd & "' OR VTP_CON_CODIGO = '" & codProd & "'")
                    If elProd.Length > 0 Then
                        strProdId = elProd(0)("CON_CODIGO").ToString
                        validProd = True
                    Else
                        validProd = False
                        rowMessage += prodNoEnc & ", "
                    End If


                    'Valida la fecha
                    If ValidateDate(codFecha) Then
                        validFecha = True
                    Else
                        validFecha = False
                        rowMessage += fechIncorrecto & ", "
                    End If

                    'Valida la hora
                    If ValidateDate(codHora) Then
                        validHora = True
                    Else
                        validHora = False
                        rowMessage = horaIncorrecto & ", "
                    End If

                    'Valida Precio
                    If codPrecio <> "" Then
                        sepDec = SeparatedDecimal(codPrecio)
                        codPrecio = sepDec.Item1 & "." & sepDec.Item2

                        Dim tupDec As Tuple(Of Decimal, Boolean) = ValidaValorDecimal(codPrecio)
                        decPrecio = tupDec.Item1
                        validPrecio = tupDec.Item2
                        If validPrecio = False Then
                            rowMessage += precioIncorrecto & ", "
                        End If
                    Else
                        validPrecio = False
                        rowMessage += precioIncorrecto & ", "
                    End If

                    If validProd And validPrecio And validNeg And validFecha And validHora Then

                        Dim consulta = "Select PVP_CODIGO, PVP_HORA From FP_RegistroPVP " & vbNewLine &
                        "Where PVP_EST_CODIGO = '" & codEst & "' And PVP_CON_CODIGO = '" & strProdId & "' " & vbNewLine &
                        "And PVP_FECHA = '" & DateTime.Parse(codFecha).ToString("yyyy-MM-dd") & "'"

                        Dim _tbl = da.Consulta(consulta, _fuelPriceComun)

                        If IsNothing(_tbl) Then
                            'Dim _inserta = "INSERT INTO dbo.FP_RegistroPVP (PVP_YEAR, PVP_FECHA, PVP_HORA, PVP_EST_CODIGO, PVP_CON_CODIGO, PVP_PRECIO, PVP_NOPERMISO) OUTPUT INSERTED.PVP_CODIGO " & vbNewLine &
                            '"VALUES (" & DateTime.Parse(codFecha).ToString("yyyy") & ", '" & DateTime.Parse(codFecha).ToString("yyyy-MM-dd") & "', '" & DateTime.Parse(codHora).ToString("HH:mm") & "'" &
                            '", '" & codEst & "', '" & codProd & "', '" & decPrecio & "', '" & codPermiso & "')"

                            Dim _inserta = "INSERT INTO dbo.FP_RegistroPVP (PVP_YEAR, PVP_FECHA, PVP_HORA, PVP_EST_CODIGO, PVP_CON_CODIGO, PVP_PRECIO, PVP_NOPERMISO) " & vbNewLine &
                            "VALUES (" & DateTime.Parse(codFecha).ToString("yyyy") & ", '" & DateTime.Parse(codFecha).ToString("yyyy-MM-dd") & "', '" & DateTime.Parse(codHora).ToString("HH:mm") & "'" &
                            ", '" & codEst & "', '" & strProdId & "', '" & decPrecio.ToString & "', '" & codPermiso & "')"

                            'Dim _dtIns As DataTable = da.Consulta(_inserta, _fuelPriceComun)
                            Dim _dtIns As Integer = da.NoQuery(_inserta, _fuelPriceComun)

                            If IsNothing(_dtIns) Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = errorRegistro & (index + 2).ToString & ": " & imposibleInsertar
                                result.Rows.Add(dr)
                            ElseIf _dtIns <= 0 Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = errorRegistro & (index + 2).ToString & ": " & imposibleInsertar
                                result.Rows.Add(dr)
                            Else
                                Dim dr As DataRow = result.NewRow
                                dr(0) = registro & " " & (index + 2).ToString & " " & insertadoExito
                                result.Rows.Add(dr)
                            End If
                        ElseIf _tbl.Rows.Count <= 0 Then
                            'Dim _inserta = "INSERT INTO dbo.FP_RegistroPVP (PVP_YEAR, PVP_FECHA, PVP_HORA, PVP_EST_CODIGO, PVP_CON_CODIGO, PVP_PRECIO, PVP_NOPERMISO) OUTPUT INSERTED.PVP_CODIGO " & vbNewLine &
                            '"VALUES (" & DateTime.Parse(codFecha).ToString("yyyy") & ", '" & DateTime.Parse(codFecha).ToString("yyyy-MM-dd") & "', '" & DateTime.Parse(codHora).ToString("HH:mm") & "'" &
                            '", '" & codEst & "', '" & codProd & "', '" & decPrecio & "', '" & codPermiso & "')"

                            Dim _inserta = "INSERT INTO dbo.FP_RegistroPVP (PVP_YEAR, PVP_FECHA, PVP_HORA, PVP_EST_CODIGO, PVP_CON_CODIGO, PVP_PRECIO, PVP_NOPERMISO) " & vbNewLine &
                            "VALUES (" & DateTime.Parse(codFecha).ToString("yyyy") & ", '" & DateTime.Parse(codFecha).ToString("yyyy-MM-dd") & "', '" & DateTime.Parse(codHora).ToString("HH:mm") & "'" &
                            ", '" & codEst & "', '" & strProdId & "', '" & sepDec.Item1 & "." & sepDec.Item2 & "', '" & codPermiso & "')"

                            'Dim _dtIns As DataTable = da.Consulta(_inserta, _fuelPriceComun)
                            Dim _dtIns As Integer = da.NoQuery(_inserta, _fuelPriceComun)

                            If IsNothing(_dtIns) Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = errorRegistro & (index + 2).ToString & ": " & imposibleInsertar
                                result.Rows.Add(dr)
                            ElseIf _dtIns <= 0 Then
                                Dim dr As DataRow = result.NewRow
                                dr(0) = errorRegistro & (index + 2).ToString & ": " & imposibleInsertar
                                result.Rows.Add(dr)
                            Else
                                Dim dr As DataRow = result.NewRow
                                dr(0) = registro & " " & (index + 2).ToString & " " & insertadoExito
                                result.Rows.Add(dr)
                            End If
                        Else
                            If _tbl.Rows(0)("PVP_HORA") < DateTime.Parse(codHora).ToString("HH:mm") Then
                                consulta = "Update FP_RegistroPVP Set " & vbNewLine &
                                "PVP_HORA = '" & DateTime.Parse(codHora).ToString("HH:mm") & "', " & vbNewLine &
                                "PVP_PRECIO = '" & sepDec.Item1 & "." & sepDec.Item2 & "' " & vbNewLine &
                                "Where PVP_CODIGO = " & _tbl.Rows(0)("PVP_CODIGO")
                            Else
                                consulta = ""
                            End If

                            If consulta <> "" Then
                                Dim _dtUpd As Integer = da.NoQuery(consulta, _fuelPriceComun)

                                If _dtUpd <= 0 Then
                                    Dim dr As DataRow = result.NewRow
                                    dr(0) = errorRegistro & (index + 2).ToString & ": " & imposibleActualizar
                                    result.Rows.Add(dr)
                                Else
                                    Dim dr As DataRow = result.NewRow
                                    dr(0) = registro & " " & (index + 2).ToString & " " & actualizadoExito
                                    result.Rows.Add(dr)
                                End If
                            Else
                                Dim dr As DataRow = result.NewRow
                                dr(0) = errorRegistro & (index + 2).ToString & ": " & horaMayor
                                result.Rows.Add(dr)
                            End If

                        End If
                    Else
                        If Mid(rowMessage, rowMessage.Length - 2, 2) = ", " Then
                            rowMessage = Mid(rowMessage, 0, rowMessage.Length - 2)
                        End If
                        Dim dr As DataRow = result.NewRow
                        dr(0) = rowMessage
                        result.Rows.Add(dr)
                    End If

                End If
            Next
        End If

        Return result
    End Function
    Public Function ImportDescuentos(ByVal byteFile, ByVal codProveedor, ByVal IdMaestro) As DataTable
        Dim result As DataTable = New DataTable
        result.Columns.Add("Informe")

        Dim codTerm As String = ""
        Dim codProd As String = ""
        Dim codDesc As String = ""
        Dim codVig As String = ""
        Dim strProdId As String = ""
        Dim decDesc As Decimal

        Dim validTerm As Boolean
        Dim validProd As Boolean
        Dim validDesc As Boolean
        Dim validVig As Boolean
        Dim validRow As Boolean

        Dim rowMessage As String = ""

        Dim _dt As DataTable = ExcelToDataTable(byteFile)
        If _dt.Columns.Count <> 4 Then
            Dim _r As DataRow = result.NewRow
            _r(0) = "Número incorrecto de columnas en el archivo"
            result.Rows.Add(_r)
        Else
            For index = 0 To _dt.Rows.Count - 1
                'CÓDIGO TERMINAL	CÓDIGO PRODUCTO	DESCUENTO REF	VIGENCIA DESDE
                codTerm = _dt.Rows(index)(0).ToString
                codProd = _dt.Rows(index)(1).ToString
                codDesc = _dt.Rows(index)(2).ToString
                codVig = _dt.Rows(index)(3).ToString

                rowMessage = "Error en registro " & (index + 2).ToString & ": "

                If codTerm = "" Or codProd = "" Or codDesc = "" Or codVig = "" Then
                    rowMessage += "Valores vacíos en el registro"
                End If

                If rowMessage.EndsWith("Valores vacíos en el registro") Then
                    Dim dr As DataRow = result.NewRow
                    dr(0) = rowMessage
                    result.Rows.Add(dr)
                Else
                    ' Valida existencia del Tar
                    Dim _tblTar As DataTable = bl.GetTarByTermCode(codTerm, IdMaestro)
                    If IsNothing(_tblTar) Then
                        validTerm = False
                        rowMessage += "Terminal no econtrada, "
                    ElseIf _tblTar.Rows.Count <= 0 Then
                        validTerm = False
                        rowMessage += "Terminal no econtrada, "
                    Else
                        codTerm = _tblTar.Rows(0)("CTS_TSU_CODIGO")
                        validTerm = True
                    End If

                    'Valida existencia del producto
                    Dim dtProd As DataTable = bl.GetProductoByShortName(codProd)
                    If IsNothing(dtProd) Then
                        validProd = False
                        rowMessage += "Producto no econtrado, "
                    ElseIf dtProd.Rows.Count <= 0 Then
                        validProd = False
                        rowMessage += "Producto no econtrado, "
                    Else
                        strProdId = dtProd.Rows(0)(0).ToString
                        validProd = True
                    End If

                    'Valida la fecha
                    validVig = ValidateDate(codVig)

                    'Valida descuento
                    If codDesc <> "" Then
                        Dim tupDec As Tuple(Of Decimal, Boolean) = ValidaValorDecimal(codDesc)
                        decDesc = tupDec.Item1
                        validDesc = tupDec.Item2
                        If validDesc = False Then
                            rowMessage += "Valor de descuento incorrecto, "
                        End If
                    Else
                        validDesc = False
                        rowMessage += "Valor de descuento incorrecto, "
                    End If

                    If validProd And validDesc And validTerm And validVig Then
                        Dim consulta As String = "Select count(*) from FP_ProveedoresDescuentos
                                                Where PDS_PRO_IDPROVEEDOR = '" & codProveedor & "' And PDS_TSU_CODIGO = '" & codTerm & "' And PDS_CON_CODIGO = '" & strProdId & "' And PDS_VIGENTEDESDE >= '" & DateTime.Parse(codVig).ToString("yyyy-MM-dd") & " 00:00:00.000'"
                        Dim _dtCount As DataTable = da.Consulta(consulta, _fuelPrice)

                        If _dtCount.Rows(0)(0) <= 0 Then
                            validRow = True
                        Else
                            validRow = False
                            rowMessage += "Ya existe una fecha posterior en el sistema, "
                        End If
                    Else
                        rowMessage += "Fecha inválida, "
                    End If

                    If validRow Then
                        Dim sepDec As Tuple(Of String, String) = SeparatedDecimal(decDesc)

                        Dim _inserta As String = "INSERT INTO FP_ProveedoresDescuentos
                                                       (PDS_PRO_IDPROVEEDOR
                                                       ,PDS_TSU_CODIGO
                                                       ,PDS_CON_CODIGO
                                                       ,PDS_DESCUENTO_TSU
                                                       ,PDS_VIGENTEDESDE)
                                                 OUTPUT INSERTED.PDS_CODIGO
                                                 VALUES
                                                       (" & codProveedor & ",
                                                        '" & codTerm & "',
                                                        '" & strProdId & "',
                                                        '" & sepDec.Item1 & "." & sepDec.Item2 & "',
                                                       '" & DateTime.Parse(codVig).ToString("yyyy-MM-dd") & " 00:00:00')"

                        Dim _dtIns As DataTable = da.Consulta(_inserta, _fuelPrice)

                        If IsNothing(_dtIns) Then
                            Dim dr As DataRow = result.NewRow
                            dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar el registro"
                            result.Rows.Add(dr)
                        ElseIf Convert.ToInt32(_dtIns.Rows(0)(0)) <= 0 Then
                            Dim dr As DataRow = result.NewRow
                            dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar el registro"
                            result.Rows.Add(dr)
                        Else
                            Dim dr As DataRow = result.NewRow
                            dr(0) = "Registro " & (index + 2).ToString & " insertado con éxito"
                            result.Rows.Add(dr)
                        End If
                    Else
                        If Mid(rowMessage, rowMessage.Length - 2, 2) = ", " Then
                            rowMessage = Mid(rowMessage, 0, rowMessage.Length - 2)
                        End If
                        Dim dr As DataRow = result.NewRow
                        dr(0) = rowMessage
                        result.Rows.Add(dr)
                    End If

                End If
            Next
        End If

        Return result
    End Function
    Public Function ImportFlete(ByVal byteFile, ByVal codProveedor, ByVal IdMaestro) As DataTable
        Dim result As DataTable = New DataTable
        result.Columns.Add("Informe")

        Dim codNeg As String = ""
        Dim codTerm As String = ""
        Dim codProd As String = ""
        Dim codFlete As String = ""
        Dim codVig As String = ""
        Dim strProdId As String = ""
        Dim CliCodigo As String = ""
        Dim decFlete As Decimal

        Dim validNeg As Boolean
        Dim validTerm As Boolean
        Dim validProd As Boolean
        Dim validFlete As Boolean
        Dim validVig As Boolean
        Dim validRow As Boolean

        Dim rowMessage As String = ""

        Dim _dt As DataTable = ExcelToDataTable(byteFile)
        If _dt.Columns.Count <> 5 Then
            Dim _r As DataRow = result.NewRow
            _r(0) = "Número incorrecto de columnas en el archivo"
            result.Rows.Add(_r)
        Else
            For index = 0 To _dt.Rows.Count - 1
                'CÓDIGO TERMINAL	CÓDIGO PRODUCTO	DESCUENTO REF	VIGENCIA DESDE
                codNeg = _dt.Rows(index)(0).ToString
                codProd = _dt.Rows(index)(1).ToString
                codFlete = _dt.Rows(index)(2).ToString
                codTerm = _dt.Rows(index)(3).ToString
                codVig = _dt.Rows(index)(4).ToString

                rowMessage = "Error en registro " & (index + 2).ToString & ": "

                If codTerm = "" Or codProd = "" Or codFlete = "" Or codVig = "" Then
                    rowMessage += "Valores vacíos en el registro"
                End If

                If rowMessage.EndsWith("Valores vacíos en el registro") Then
                    Dim dr As DataRow = result.NewRow
                    dr(0) = rowMessage
                    result.Rows.Add(dr)
                Else
                    'Valida existencia del Negocio
                    Dim dtCte As DataTable = bl.GetClientById(codNeg)
                    If IsNothing(dtCte) Then
                        validNeg = False
                        rowMessage += "Negocio no econtrado, "
                    ElseIf bl.GetClientById(codNeg).Rows.Count <= 0 Then
                        validNeg = False
                        rowMessage += "Negocio no econtrado, "
                    Else
                        CliCodigo = dtCte.Rows(0)("CLI_CODIGO").ToString()
                        validNeg = True
                    End If

                    If validNeg = False Then
                        dtCte = bl.GetNegociosPorPermiso(codNeg)
                        If IsNothing(dtCte) Then
                            validNeg = False
                            rowMessage += "Negocio no econtrado, "
                        ElseIf bl.GetClientById(codNeg).Rows.Count <= 0 Then
                            validNeg = False
                            rowMessage += "Negocio no econtrado, "
                        Else
                            CliCodigo = dtCte.Rows(0)("CLI_CODIGO").ToString()
                            validNeg = True
                        End If
                    End If

                    ' Valida existencia del Tar
                    Dim _tblTar As DataTable = bl.GetTarByTermCode(codTerm, IdMaestro)
                    If IsNothing(_tblTar) Then
                        validTerm = False
                        rowMessage += "Terminal no econtrada, "
                    ElseIf _tblTar.Rows.Count <= 0 Then
                        validTerm = False
                        rowMessage += "Terminal no econtrada, "
                    Else
                        codTerm = _tblTar.Rows(0)("CTS_TSU_CODIGO")
                        validTerm = True
                    End If

                    'Valida existencia del producto
                    Dim dtProd As DataTable = bl.GetProductoByShortName(codProd)
                    If IsNothing(dtProd) Then
                        validProd = False
                        rowMessage += "Producto no econtrado, "
                    ElseIf dtProd.Rows.Count <= 0 Then
                        validProd = False
                        rowMessage += "Producto no econtrado, "
                    Else
                        strProdId = dtProd.Rows(0)(0).ToString
                        validProd = True
                    End If

                    'Valida la fecha
                    validVig = ValidateDate(codVig)

                    'Valida descuento
                    If codFlete <> "" Then
                        Dim tupDec As Tuple(Of Decimal, Boolean) = ValidaValorDecimal(codFlete)
                        decFlete = tupDec.Item1
                        validFlete = tupDec.Item2
                        If validFlete = False Then
                            rowMessage += "Valor de flete incorrecto, "
                        End If
                    Else
                        validFlete = False
                        rowMessage += "Valor de flete incorrecto, "
                    End If

                    If validProd And validFlete And validTerm And validVig Then
                        Dim consulta As String = "Select count(*) from FP_ProveedoresFletes
                                            Where PFL_PRO_IDPROVEEDOR = '" & codProveedor & "' And PFL_CLI_CODIGO = '" & CliCodigo & "' And PFL_TSU_ENTREGA_CODIGO = '" & codTerm & "' And PFL_CON_CODIGO = '" & strProdId & "' And PFL_VIGENTEDESDE >= '" & DateTime.Parse(codVig).ToString("yyyy-MM-dd") & " 00:00:00.000'"
                        Dim _dtCount As DataTable = da.Consulta(consulta, _fuelPrice)

                        If _dtCount.Rows(0)(0) <= 0 Then
                            validRow = True
                        Else
                            validRow = False
                            rowMessage += "Ya existe una fecha posterior en el sistema, "
                        End If
                    Else
                        rowMessage += "Fecha inválida, "
                    End If

                    If validRow Then
                        Dim sepDec As Tuple(Of String, String) = SeparatedDecimal(decFlete)

                        Dim _inserta As String = "INSERT INTO FP_ProveedoresFletes
                                                   (PFL_PRO_IDPROVEEDOR
                                                   ,PFL_CLI_CODIGO
                                                   ,PFL_CON_CODIGO
                                                   ,PFL_FLETE
                                                   ,PFL_TSU_ENTREGA_CODIGO
                                                   ,PFL_VIGENTEDESDE)
                                             OUTPUT INSERTED.PFL_CODIGO
                                             VALUES
                                                   (" & codProveedor & ",
                                                    '" & CliCodigo & "',
                                                    '" & strProdId & "',
                                                    '" & sepDec.Item1 & "." & sepDec.Item2 & "',
                                                    '" & codTerm & "',
                                                   '" & DateTime.Parse(codVig).ToString("yyyy-MM-dd") & " 00:00:00')"

                        Dim _dtIns As DataTable = da.Consulta(_inserta, _fuelPrice)

                        If IsNothing(_dtIns) Then
                            Dim dr As DataRow = result.NewRow
                            dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar el registro"
                            result.Rows.Add(dr)
                        ElseIf Convert.ToInt32(_dtIns.Rows(0)(0)) <= 0 Then
                            Dim dr As DataRow = result.NewRow
                            dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar el registro"
                            result.Rows.Add(dr)
                        Else
                            Dim dr As DataRow = result.NewRow
                            dr(0) = "Registro " & (index + 2).ToString & " insertado con éxito"
                            result.Rows.Add(dr)
                        End If
                    Else
                        If Mid(rowMessage, rowMessage.Length - 2, 2) = ", " Then
                            rowMessage = Mid(rowMessage, 0, rowMessage.Length - 2)
                        End If
                        Dim dr As DataRow = result.NewRow
                        dr(0) = rowMessage
                        result.Rows.Add(dr)
                    End If

                End If
            Next
        End If

        Return result
    End Function
    Public Function ImportVolumen(ByVal byteFile, ByVal codProveedor, ByVal IdMaestro) As DataTable
        Dim result As DataTable = New DataTable
        result.Columns.Add("Informe")

        Dim codTerm As String = ""
        Dim codProd As String = ""
        Dim codVol As String = ""
        Dim codAnio As String = ""
        Dim codMes As String = ""
        Dim strProdId As String = ""
        Dim decVol As Decimal
        Dim intAnio As Integer
        Dim intMes As Integer

        Dim validTerm As Boolean
        Dim validProd As Boolean
        Dim validvol As Boolean
        Dim validAnio As Boolean
        Dim validMes As Boolean
        Dim validRow As Boolean

        Dim rowMessage As String = ""

        Dim _dt As DataTable = ExcelToDataTable(byteFile)
        If _dt.Columns.Count <> 5 Then
            Dim _r As DataRow = result.NewRow
            _r(0) = "Número incorrecto de columnas en el archivo"
            result.Rows.Add(_r)
        Else
            For index = 0 To _dt.Rows.Count - 1
                'CÓDIGO TERMINAL	CÓDIGO PRODUCTO	DESCUENTO REF	VIGENCIA DESDE
                codTerm = _dt.Rows(index)(0).ToString
                codProd = _dt.Rows(index)(1).ToString
                codVol = _dt.Rows(index)(2).ToString
                codAnio = _dt.Rows(index)(3).ToString
                codMes = _dt.Rows(index)(4).ToString

                rowMessage = "Error en registro " & (index + 2).ToString & ": "

                If codTerm = "" Or codProd = "" Or codVol = "" Or codAnio = "" Or codMes = "" Then
                    rowMessage += "Valores vacíos en el registro"
                End If

                If rowMessage.EndsWith("Valores vacíos en el registro") Then
                    Dim dr As DataRow = result.NewRow
                    dr(0) = rowMessage
                    result.Rows.Add(dr)
                Else
                    ' Valida existencia del Tar
                    Dim _tblTar As DataTable = bl.GetTarByTermCode(codTerm, IdMaestro)
                    If IsNothing(_tblTar) Then
                        validTerm = False
                        rowMessage += "Terminal no econtrada, "
                    ElseIf _tblTar.Rows.Count <= 0 Then
                        validTerm = False
                        rowMessage += "Terminal no econtrada, "
                    Else
                        codTerm = _tblTar.Rows(0)("CTS_TSU_CODIGO")
                        validTerm = True
                    End If

                    'Valida existencia del producto
                    Dim dtProd As DataTable = bl.GetProductoByShortName(codProd)
                    If IsNothing(dtProd) Then
                        validProd = False
                        rowMessage += "Producto no econtrado, "
                    ElseIf dtProd.Rows.Count <= 0 Then
                        validProd = False
                        rowMessage += "Producto no econtrado, "
                    Else
                        strProdId = dtProd.Rows(0)(0).ToString
                        validProd = True
                    End If

                    'Valida Volumen
                    If codVol <> "" Then
                        Dim tupDec As Tuple(Of Decimal, Boolean) = ValidaValorDecimal(codVol)
                        decVol = tupDec.Item1
                        validvol = tupDec.Item2
                        If validvol = False Then
                            rowMessage += "Valor de descuento incorrecto, "
                        End If
                    Else
                        validvol = False
                        rowMessage += "Valor de descuento incorrecto, "
                    End If

                    If codAnio <> "" Then
                        Dim tupInt As Tuple(Of Integer, Boolean) = ValidaValorEntero(codAnio)
                        intAnio = tupInt.Item1
                        validAnio = tupInt.Item2
                        If validAnio = False Then
                            rowMessage += "Valor de año incorrecto, "
                        End If
                    Else
                        validAnio = False
                        rowMessage += "Valor de año incorrecto, "
                    End If

                    If codMes <> "" Then
                        Dim tupInt As Tuple(Of Integer, Boolean) = ValidaValorEntero(codMes)
                        intMes = tupInt.Item1
                        validMes = tupInt.Item2
                        If validMes = False Then
                            rowMessage += "Valor de mes incorrecto, "
                        End If
                    Else
                        validMes = False
                        rowMessage += "Valor de mes incorrecto, "
                    End If

                    If validAnio And validMes And validTerm And validProd And validvol Then
                        Dim consulta As String = "Select count(*) from FP_ProveedoresVolumenes
                                            Where PVO_PRO_IDPROVEEDOR = '" & codProveedor & "' And PVO_TSU_CODIGO = '" & codTerm & "' And PVO_CON_CODIGO = '" & strProdId & "' And PVO_AÑO = '" & intAnio & "' And PVO_MES = '" & intMes & "'"
                        Dim _dtCount As DataTable = da.Consulta(consulta, _fuelPrice)

                        If _dtCount.Rows(0)(0) <= 0 Then
                            validRow = True
                        Else
                            validRow = False
                            rowMessage += "Ya existe este registro en el sistema, "
                        End If
                    Else
                        rowMessage += "Fecha inválida, "
                    End If

                    If validRow Then
                        Dim sepDec As Tuple(Of String, String) = SeparatedDecimal(decVol)

                        Dim _inserta As String = "INSERT INTO FP_ProveedoresVolumenes
                                                   (PVO_PRO_IDPROVEEDOR
                                                   ,PVO_TSU_CODIGO
                                                   ,PVO_CON_CODIGO
                                                   ,PVO_VOLUMENCOMPROMETIDO
                                                   ,PVO_AÑO
                                                   ,PVO_MES)
                                             OUTPUT INSERTED.PVO_CODIGO
                                             VALUES
                                                   (" & codProveedor & ",
                                                   '" & codTerm & "',
                                                   '" & strProdId & "', 
                                                   '" & sepDec.Item1 & "." & sepDec.Item2 & "',
                                                   '" & intAnio & "',
                                                   '" & intMes & "')"

                        Dim _dtIns As DataTable = da.Consulta(_inserta, _fuelPrice)

                        If IsNothing(_dtIns) Then
                            Dim dr As DataRow = result.NewRow
                            dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar el registro"
                            result.Rows.Add(dr)
                        ElseIf Convert.ToInt32(_dtIns.Rows(0)(0)) <= 0 Then
                            Dim dr As DataRow = result.NewRow
                            dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar el registro"
                            result.Rows.Add(dr)
                        Else
                            Dim dr As DataRow = result.NewRow
                            dr(0) = "Registro " & (index + 2).ToString & " insertado con éxito"
                            result.Rows.Add(dr)
                        End If
                    Else
                        If Mid(rowMessage, rowMessage.Length - 2, 2) = ", " Then
                            rowMessage = Mid(rowMessage, 0, rowMessage.Length - 2)
                        End If
                        Dim dr As DataRow = result.NewRow
                        dr(0) = rowMessage
                        result.Rows.Add(dr)
                    End If

                End If
            Next
        End If

        Return result
    End Function
    Public Function ImportTemperatura(ByVal byteFile, ByVal codProveedor, ByVal IdMaestro) As DataTable
        Dim result As DataTable = New DataTable
        result.Columns.Add("Informe")

        Dim codTerm As String = ""
        Dim codProd As String = ""
        Dim codTemp As String = ""
        Dim codVig As String = ""
        Dim strProdId As String = ""
        Dim decDesc As Decimal

        Dim validTerm As Boolean
        Dim validProd As Boolean
        Dim validTemp As Boolean
        Dim validVig As Boolean
        Dim validRow As Boolean

        Dim rowMessage As String = ""

        Dim _dt As DataTable = ExcelToDataTable(byteFile)
        If _dt.Columns.Count <> 4 Then
            Dim _r As DataRow = result.NewRow
            _r(0) = "Número incorrecto de columnas en el archivo"
            result.Rows.Add(_r)
        Else
            For index = 0 To _dt.Rows.Count - 1
                'CÓDIGO TERMINAL	CÓDIGO PRODUCTO	DESCUENTO REF	VIGENCIA DESDE
                codTerm = _dt.Rows(index)(0).ToString
                codProd = _dt.Rows(index)(1).ToString
                codTemp = _dt.Rows(index)(2).ToString
                codVig = _dt.Rows(index)(3).ToString

                rowMessage = "Error en registro " & (index + 2).ToString & ": "

                If codTerm = "" Or codProd = "" Or codTemp = "" Or codVig = "" Then
                    rowMessage += "Valores vacíos en el registro"
                Else

                End If

                If rowMessage.EndsWith("Valores vacíos en el registro") Then
                    Dim dr As DataRow = result.NewRow
                    dr(0) = rowMessage
                    result.Rows.Add(dr)
                Else
                    ' Valida existencia del Tar
                    Dim _tblTar As DataTable = bl.GetTarByTermCode(codTerm, IdMaestro)
                    If IsNothing(_tblTar) Then
                        validTerm = False
                        rowMessage += "Terminal no econtrada, "
                    ElseIf _tblTar.Rows.Count <= 0 Then
                        validTerm = False
                        rowMessage += "Terminal no econtrada, "
                    Else
                        codTerm = _tblTar.Rows(0)("CTS_TSU_CODIGO")
                        validTerm = True
                    End If

                    'Valida existencia del producto
                    Dim dtProd As DataTable = bl.GetProductoByShortName(codProd)
                    If IsNothing(dtProd) Then
                        validProd = False
                        rowMessage += "Producto no econtrado, "
                    ElseIf dtProd.Rows.Count <= 0 Then
                        validProd = False
                        rowMessage += "Producto no econtrado, "
                    Else
                        strProdId = dtProd.Rows(0)(0).ToString
                        validProd = True
                    End If

                    'Valida la fecha
                    validVig = ValidateDate(codVig)

                    'Valida descuento
                    If codTemp <> "" Then
                        Dim tupDec As Tuple(Of Decimal, Boolean) = ValidaValorDecimal(codTemp)
                        decDesc = tupDec.Item1
                        validTemp = tupDec.Item2
                        If validTemp = False Then
                            rowMessage += "Valor de temperatura incorrecto, "
                        End If
                    Else
                        validTemp = False
                        rowMessage += "Valor de temperatura incorrecto, "
                    End If

                    If validVig And validTemp And validProd And validTerm Then
                        Dim consulta As String = "Select count(*) from FP_ProveedoresTemperatura
                                            Where PTM_PRO_IDPROVEEDOR = '" & codProveedor & "' And PTM_TSU_CODIGO = '" & codTerm & "' And PTM_CON_CODIGO = '" & strProdId & "' And PTM_VIGENTEDESDE >= '" & DateTime.Parse(codVig).ToString("yyyy-MM-dd") & " 00:00:00.000'"
                        Dim _dtCount As DataTable = da.Consulta(consulta, _fuelPrice)

                        If _dtCount.Rows(0)(0) <= 0 Then
                            validRow = True
                        Else
                            validRow = False
                            rowMessage += "Ya existe este registro en el sistema, "
                        End If
                    Else
                        rowMessage += "Fecha inválida, "
                    End If

                    If validRow Then
                        Dim sepDec As Tuple(Of String, String) = SeparatedDecimal(decDesc)
                        Dim _inserta As String = "INSERT INTO FP_ProveedoresTemperatura
                                                   (PTM_PRO_IDPROVEEDOR
                                                   ,PTM_TSU_CODIGO
                                                   ,PTM_CON_CODIGO
                                                   ,PTM_TEMPERATURA_REF
                                                   ,PTM_VIGENTEDESDE)
                                             OUTPUT INSERTED.PTM_CODIGO
                                             VALUES
                                                   (" & codProveedor & ",
                                                    '" & codTerm & "',
                                                    '" & strProdId & "',
                                                    '" & sepDec.Item1 & "." & sepDec.Item2 & "',
                                                   '" & DateTime.Parse(codVig).ToString("yyyy-MM-dd") & " 00:00:00')"

                        Dim _dtIns As DataTable = da.Consulta(_inserta, _fuelPrice)

                        If IsNothing(_dtIns) Then
                            Dim dr As DataRow = result.NewRow
                            dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar el registro"
                            result.Rows.Add(dr)
                        ElseIf Convert.ToInt32(_dtIns.Rows(0)(0)) <= 0 Then
                            Dim dr As DataRow = result.NewRow
                            dr(0) = "Error en registro " & (index + 2).ToString & ": Imposible insertar el registro"
                            result.Rows.Add(dr)
                        Else
                            Dim dr As DataRow = result.NewRow
                            dr(0) = "Registro " & (index + 2).ToString & " insertado con éxito"
                            result.Rows.Add(dr)
                        End If
                    Else
                        If Mid(rowMessage, rowMessage.Length - 2, 2) = ", " Then
                            rowMessage = Mid(rowMessage, 0, rowMessage.Length - 2)
                        End If
                        Dim dr As DataRow = result.NewRow
                        dr(0) = rowMessage
                        result.Rows.Add(dr)
                    End If

                End If
            Next
        End If

        Return result
    End Function
    Public Function ValidateDate(ByVal Fecha) As Boolean
        Dim _date As DateTime
        Dim result As Boolean
        Try
            _date = DateTime.Parse(Fecha)
            result = True
        Catch
            result = False
        End Try
        Return result
    End Function
    Public Function ValidaValorDecimal(ByVal Valor) As Tuple(Of Decimal, Boolean)
        Dim result As Decimal
        Dim evaluado As Boolean
        Try
            result = Decimal.Parse(Valor)
            evaluado = True
        Catch ex As Exception
            result = -1
            evaluado = False
        End Try

        Return New Tuple(Of Decimal, Boolean)(result, evaluado)
    End Function
    Public Function ValidaValorEntero(ByVal Valor) As Tuple(Of Integer, Boolean)
        Dim result As Integer
        Dim evaluado As Boolean
        Try
            result = Integer.Parse(Valor)
            evaluado = True
        Catch ex As Exception
            result = -1
            evaluado = False
        End Try

        Return New Tuple(Of Integer, Boolean)(result, evaluado)
    End Function
    Public Function ValidaValorFecha(ByVal Valor) As Tuple(Of DateTime, Boolean)
        Dim result As DateTime
        Dim evaluado As Boolean
        Try
            result = DateTime.Parse(Valor)
            evaluado = True
        Catch ex As Exception
            result = Nothing
            evaluado = False
        End Try

        Return New Tuple(Of DateTime, Boolean)(result, evaluado)
    End Function
    Public Function SeparatedDecimal(ByVal ValorDecimal) As Tuple(Of String, String)
        Dim separatedValue As String()

        If ValorDecimal.ToString.Contains(",") Then
            separatedValue = ValorDecimal.ToString.Split(",")
        ElseIf ValorDecimal.ToString.Contains(".") Then
            separatedValue = ValorDecimal.ToString.Split(".")
        Else
            separatedValue = {ValorDecimal, 0}
        End If

        Return New Tuple(Of String, String)(separatedValue(0), separatedValue(1))
    End Function
End Class
