Imports System.Data
Imports System.IO
Imports ClosedXML.Excel
Imports FuelPrice.Business
Imports Microsoft.Office.Interop.Excel
Imports Newtonsoft.Json

Public Class Excel

    Public Shared Sub CrearExcelHistoricoCostos(linhas As Integer, matriz As Object(,), nombreArchivo As String)
        Dim workbook As New XLWorkbook
        Dim rowTablaIndex As Integer = 1

        Dim worksheet = workbook.Worksheets.Add("HistoricoCostos")
        worksheet.Cell(rowTablaIndex, 1).Value = "Historico Costos"
        Dim Titulo As IXLRange = worksheet.Range(($"A" & rowTablaIndex & ":H" & rowTablaIndex))
        Titulo.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
        Titulo.Merge()
        Titulo.Style.Font.SetBold(True)
        Titulo.Style.Font.SetFontSize(30)
        rowTablaIndex += 1

        CrearCabecalhoColumnas(worksheet)

        rowTablaIndex += 1

        For linha = 0 To linhas
            For coluna = 0 To 7
                If coluna = 0 Then
                    If IsNothing(matriz(linha, coluna)) And IsNothing(matriz(linha, coluna + 1)) And IsNothing(matriz(linha, coluna + 2)) Then
                        Exit For
                    ElseIf IsNothing(matriz(linha, coluna)) And Not IsNothing(matriz(linha, coluna + 1)) Then
                        worksheet.Cell(rowTablaIndex, 1).Value = ""
                    ElseIf IsNothing(matriz(linha, coluna)) And IsNothing(matriz(linha, coluna + 1)) Then
                        worksheet.Cell(rowTablaIndex, 1).Value = ""
                    Else
                        worksheet.Cell(rowTablaIndex, 1).Value = matriz(linha, coluna)
                        Dim intervaloColumnas As IXLCell = worksheet.Cell(rowTablaIndex, 1)
                        intervaloColumnas.Style.Fill.SetBackgroundColor(XLColor.FromArgb(220, 233, 210))
                    End If
                ElseIf coluna = 1 Then
                    If IsNothing(matriz(linha, coluna)) Then
                        worksheet.Cell(rowTablaIndex, 2).Value = ""
                    Else
                        worksheet.Cell(rowTablaIndex, 2).Value = matriz(linha, coluna)
                        Dim intervaloColumnas As IXLCell = worksheet.Cell(rowTablaIndex, 2)
                        intervaloColumnas.Style.Fill.SetBackgroundColor(XLColor.FromArgb(222, 235, 246))
                    End If
                ElseIf coluna = 7 Then
                    worksheet.Cell(rowTablaIndex, 8).Value = matriz(linha, coluna)
                    rowTablaIndex += 1
                Else
                    worksheet.Cell(rowTablaIndex, (coluna + 1)).Value = matriz(linha, coluna)
                    Dim intervaloColumnas As IXLRange = worksheet.Range($"D" & rowTablaIndex, "H" & (coluna + 1))
                    intervaloColumnas.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                End If
            Next
        Next
        worksheet.Columns("A", "H").AdjustToContents()

        workbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory & $"{nombreArchivo}.xlsx")
    End Sub

    Public Shared Sub CrearExcelHistoricoPRT(linhas As Integer, matriz As Object(,), cabecalho As String, nombreArchivo As String, rangofecha As String)
        Dim workbook As New XLWorkbook
        Dim rowTablaIndex As Integer = 1

        Dim nomesCabecalhos = cabecalho.Split("|")

        Dim worksheet = workbook.Worksheets.Add("HistoricoPreciosRefTermina")
        worksheet.Cell(rowTablaIndex, 1).Value = "HISTORICO PRECIOS REF TERMINAL"

        Dim Titulo As IXLRange = worksheet.Range(($"A" & rowTablaIndex & ":K" & rowTablaIndex))
        Titulo.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
        Titulo.Merge()
        Titulo.Style.Font.SetBold(True)
        Titulo.Style.Font.SetFontSize(30)
        rowTablaIndex += 1

        Titulo = worksheet.Range(($"A" & rowTablaIndex & ":K" & rowTablaIndex))
        worksheet.Cell(rowTablaIndex, 1).Value = "Fecha: " & rangofecha
        Titulo.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
        Titulo.Merge()
        Titulo.Style.Font.SetBold(True)
        Titulo.Style.Font.SetFontSize(15)
        rowTablaIndex += 1

        CrearCabecalhoColumnasPRT(worksheet, nomesCabecalhos)
        rowTablaIndex += 1

        Dim cor = 0
        Dim stringCor = XLColor.Alizarin
        Dim stringCorProducto = XLColor.Alizarin
        For linha = 0 To linhas
            If cor = 0 And Not IsNothing(matriz(linha, 0)) Then
                cor = 1
                stringCor = XLColor.FromArgb(252, 228, 214)
                stringCorProducto = XLColor.FromArgb(198, 224, 180)
            ElseIf cor = 1 And Not IsNothing(matriz(linha, 0)) Then
                cor = 0
                stringCor = XLColor.FromArgb(189, 215, 238)
                stringCorProducto = XLColor.FromArgb(255, 230, 153)
            End If
            For coluna = 0 To nomesCabecalhos.Count - 1
                If coluna = 0 Then
                    If IsNothing(matriz(linha, coluna + 1)) Then
                        Exit For
                    ElseIf IsNothing(matriz(linha, coluna)) Then
                        worksheet.Cell(rowTablaIndex, 1).Value = ""
                    Else
                        worksheet.Cell(rowTablaIndex, 1).Value = matriz(linha, coluna)
                        Dim intervaloColumnas As IXLCell = worksheet.Cell(rowTablaIndex, 1)
                        intervaloColumnas.Style.Fill.SetBackgroundColor(stringCorProducto)
                    End If
                ElseIf coluna = 1 Then
                    worksheet.Cell(rowTablaIndex, (coluna + 1)).Value = matriz(linha, coluna)
                    Dim intervaloColumnas As IXLRange = worksheet.Range($"B" & rowTablaIndex, "K" & rowTablaIndex)
                    intervaloColumnas.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    intervaloColumnas.Style.Fill.SetBackgroundColor(stringCor)
                ElseIf coluna = nomesCabecalhos.Count - 1 Then
                    worksheet.Cell(rowTablaIndex, (coluna + 1)).Value = $"' {matriz(linha, coluna)}"
                    Dim intervaloColumnas As IXLRange = worksheet.Range($"B" & rowTablaIndex, "k" & rowTablaIndex)
                    intervaloColumnas.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    intervaloColumnas.Style.Fill.SetBackgroundColor(stringCor)
                    rowTablaIndex += 1
                Else
                    worksheet.Cell(rowTablaIndex, (coluna + 1)).Value = $"' {matriz(linha, coluna)}"
                    Dim intervaloColumnas As IXLRange = worksheet.Range($"B" & rowTablaIndex, "K" & rowTablaIndex)
                    intervaloColumnas.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    intervaloColumnas.Style.Fill.SetBackgroundColor(stringCor)
                End If
            Next
        Next
        worksheet.Columns("A", "k").AdjustToContents()
        workbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory & $"{nombreArchivo}.xlsx")
    End Sub

    Public Shared Sub CrearExcelHistoricoDA(linhas As Integer, matriz As Object(,), cabecalho As String, nombreArchivo As String, rangofecha As String)
        Dim workbook As New XLWorkbook
        Dim rowTablaIndex As Integer = 1

        Dim Worksheet = workbook.Worksheets.Add("HistoricoDescuentosRefTerminal")

        Dim nomesCabecalhos = cabecalho.Split("|")

        Worksheet.Cell(rowTablaIndex, 1).Value = "HISTORICO DESCUENTOS REF TERMINAL"
        Dim Titulo As IXLRange = Worksheet.Range($"A" & rowTablaIndex, "K" & rowTablaIndex)
        Titulo.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
        Titulo.Merge()
        Titulo.Style.Font.SetBold(True)
        Titulo.Style.Font.SetFontSize(30)
        rowTablaIndex += 1

        Titulo = Worksheet.Range($"A" & rowTablaIndex, "K" & rowTablaIndex)
        Worksheet.Cell(rowTablaIndex, 1).Value = "Fecha: " & rangofecha

        Titulo.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
        Titulo.Merge()
        Titulo.Style.Font.SetBold(True)
        Titulo.Style.Font.SetFontSize(15)
        rowTablaIndex += 1

        CrearCabecalhoColumnasPRT(Worksheet, nomesCabecalhos)
        rowTablaIndex += 1

        Dim cor = 0
        Dim stringCor = XLColor.White
        Dim stringCorProducto = XLColor.White
        For linha = 0 To linhas
            If cor = 0 And Not IsNothing(matriz(linha, 0)) Then
                cor = 1
                stringCor = XLColor.FromArgb(252, 228, 214)
                stringCorProducto = XLColor.FromArgb(198, 224, 180)
            ElseIf cor = 1 And Not IsNothing(matriz(linha, 0)) Then
                cor = 0
                stringCor = XLColor.FromArgb(189, 215, 238)
                stringCorProducto = XLColor.FromArgb(255, 230, 153)
            End If
            For coluna = 0 To nomesCabecalhos.Count - 1
                If coluna = 0 Then
                    If IsNothing(matriz(linha, coluna + 1)) Then
                        Exit For
                    ElseIf IsNothing(matriz(linha, coluna)) Then
                        Worksheet.Cell(rowTablaIndex, 1).Value = ""
                    Else
                        Worksheet.Cell(rowTablaIndex, 1).Value = matriz(linha, coluna)
                        Dim intervaloColumnas As IXLCell = Worksheet.Cell(rowTablaIndex, 1)
                        intervaloColumnas.Style.Fill.SetBackgroundColor(stringCorProducto)
                    End If
                ElseIf coluna = 1 Then
                    Worksheet.Cell(rowTablaIndex, (coluna + 1)).Value = matriz(linha, coluna)
                    Dim intervaloColumnas As IXLRange = Worksheet.Range(($"B" & rowTablaIndex & ":K" & rowTablaIndex))
                    intervaloColumnas.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    intervaloColumnas.Style.Fill.SetBackgroundColor(stringCor)
                ElseIf coluna = nomesCabecalhos.Count - 1 Then
                    Worksheet.Cell(rowTablaIndex, (coluna + 1)).Value = $"' {matriz(linha, coluna)}"
                    Dim intervaloColumnas As IXLRange = Worksheet.Range(($"B" & rowTablaIndex & ":k" & rowTablaIndex))
                    intervaloColumnas.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    intervaloColumnas.Style.Fill.SetBackgroundColor(stringCor)
                    rowTablaIndex += 1
                Else
                    Worksheet.Cell(rowTablaIndex, (coluna + 1)).Value = $"' {matriz(linha, coluna)}"
                    Dim intervaloColumnas As IXLRange = Worksheet.Range(($"B" & rowTablaIndex & ":K" & rowTablaIndex))
                    intervaloColumnas.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    intervaloColumnas.Style.Fill.SetBackgroundColor(stringCor)
                End If
            Next
        Next
        Worksheet.Columns("A", "K").AdjustToContents()
        workbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory & $"{nombreArchivo}.xlsx")
    End Sub

    Public Shared Sub CrearExcelHistoricoPST(dt As Data.DataTable, nombreArchivo As String, rangofecha As String)
        Dim workbook As New XLWorkbook
        Dim rowTablaIndex As Integer = 1

        Dim _excel = workbook.Worksheets.Add("HistoricoPreciosSpotTerminal")

        _excel.Cell(rowTablaIndex, 1).Value = "HISTORICO PRECIOS SPOT TERMINAL"
        Dim Titulo As IXLRange = _excel.Range(($"A" & rowTablaIndex & ":E" & rowTablaIndex))
        Titulo.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
        Titulo.Merge()
        Titulo.Style.Font.SetBold(True)
        Titulo.Style.Font.SetFontSize(30)
        rowTablaIndex += 1

        Titulo = _excel.Range(($"A" & rowTablaIndex & ":E" & rowTablaIndex))
        _excel.Cell(rowTablaIndex, 1).Value = "Fecha: " & rangofecha
        Titulo.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
        Titulo.Merge()
        Titulo.Style.Font.SetBold(True)
        Titulo.Style.Font.SetFontSize(15)
        rowTablaIndex += 1

        CrearCabecalhoColumnasPST(_excel)
        rowTablaIndex += 1

        Dim ContData = 0
        Dim ContProveedor = 0
        Dim ContProducto = 0

        Dim grupoTar = dt.AsEnumerable().GroupBy(Function(x) x.Item("TPS_TSU_CODIGODESCRIPICION")).OrderBy(Function(z) z.Key)
        For Each itemTar In grupoTar
            _excel.Cell(rowTablaIndex, 1).Value = itemTar.Key
            Dim intervaloColumnas As IXLCell = _excel.Cell(rowTablaIndex, 1)
            intervaloColumnas.Style.Fill.SetBackgroundColor(XLColor.FromArgb(249, 236, 197))
            Dim grupoProveedor = itemTar.GroupBy(Function(x) x.Item("TPS_PRO_CODIGODESCRIPICION")).OrderBy(Function(z) z.Key)
            For Each itemProveedor In grupoProveedor
                ContProveedor += 1
                _excel.Cell(rowTablaIndex, 2).Value = itemProveedor.Key
                intervaloColumnas = _excel.Cell(rowTablaIndex, 2)
                intervaloColumnas.Style.Fill.SetBackgroundColor(XLColor.FromArgb(215, 228, 239))
                Dim grupoProducto = itemProveedor.GroupBy(Function(x) x.Item("TPS_CON_CODIGODESCRIPICION")).OrderBy(Function(z) z.Key)
                For Each itemProducto In grupoProducto
                    ContProducto += 1
                    _excel.Cell(rowTablaIndex, 3).Value = itemProducto.Key.Replace("<", "‹") _
                                                   .Replace(">", "›")
                    intervaloColumnas = _excel.Cell(rowTablaIndex, 3)
                    intervaloColumnas.Style.Fill.SetBackgroundColor(XLColor.FromArgb(206, 224, 191))
                    Dim grupoFecha = itemProducto.GroupBy(Function(x) CDate(x.Item("TPS_VIGENTEDESDE"))).OrderBy(Function(z) CDate(z.Key))
                    For Each itemFecha In grupoFecha
                        ContData += 1
                        Dim somaImport As String = FormatNumber(itemFecha.Sum(Function(x) x.Item("TPS_IMPORTE")), 6).ToString
                        _excel.Cell(rowTablaIndex, 4).Value = Format(CDate(itemFecha.Key), "yyyy/MM/dd")
                        _excel.Cell(rowTablaIndex, 5).Value = $"' {somaImport} "
                        Dim intervaloColumnasL As IXLRange = _excel.Range(($"D" & rowTablaIndex & ":E" & rowTablaIndex))
                        intervaloColumnasL.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                        If ContData < grupoFecha.Count Then
                            rowTablaIndex += 1
                        End If
                    Next
                    ContData = 0
                    If ContProducto < grupoProducto.Count Then
                        rowTablaIndex += 1
                    End If

                Next
                ContProducto = 0
                If ContProveedor < grupoProveedor.Count Then
                    rowTablaIndex += 1
                End If
            Next
            rowTablaIndex += 1
            ContProveedor = 0
        Next

        _excel.Columns("A", "E").AdjustToContents()

        workbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory & $"{nombreArchivo}.xlsx")
    End Sub

    Private Shared Sub CrearCabecalhoColumnas(ByRef _excel As IXLWorksheet)
        _excel.Cell(2, 1).Value = "TERMINAL"
        _excel.Cell(2, 2).Value = "PRODUCTO"
        _excel.Cell(2, 3).Value = "COSTO TERMINAL"
        _excel.Cell(2, 4).Value = "IMPORTE Y FECHA"
        _excel.Cell(2, 5).Value = "IMPORTE Y FECHA"
        _excel.Cell(2, 6).Value = "IMPORTE Y FECHA"
        _excel.Cell(2, 7).Value = "IMPORTE Y FECHA"
        _excel.Cell(2, 8).Value = "IMPORTE Y FECHA"

        Dim intervaloColumnas As IXLRange = _excel.Range($"A" & 2 & ":H" & 2)
        intervaloColumnas.Style.Fill.SetBackgroundColor(XLColor.FromArgb(217, 217, 217))
        intervaloColumnas.Style.Font.SetBold(True)
        intervaloColumnas.Style.Border.SetBottomBorderColor(XLColor.Black)
        intervaloColumnas.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
    End Sub

    Private Shared Sub CrearCabecalhoColumnasPRT(ByRef _excel As IXLWorksheet, ByVal quantidadeColunas As String())
        _excel.Cell(3, 1).Value = "PRODUCTO"
        _excel.Cell(3, 2).Value = "FECHA"

        For index = 3 To quantidadeColunas.Count
            _excel.Cell(3, index).Value = quantidadeColunas(index - 1)
        Next

        Dim intervaloColumnas As IXLRange = _excel.Range(($"A" & 3 & ":K" & 3))
        intervaloColumnas.Style.Fill.SetBackgroundColor(XLColor.FromArgb(217, 217, 217))
        intervaloColumnas.Style.Font.SetBold(True)
        intervaloColumnas.Style.Border.SetBottomBorderColor(XLColor.Black)
        intervaloColumnas.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
    End Sub

    Private Shared Sub CrearCabecalhoColumnasPST(ByRef _excel As IXLWorksheet)
        _excel.Cell(3, 1).Value = "TERMINAL"
        _excel.Cell(3, 2).Value = "PROVEEDOR"
        _excel.Cell(3, 3).Value = "PRODUCTO"
        _excel.Cell(3, 4).Value = "FECHA"
        _excel.Cell(3, 5).Value = "IMPORTE"

        Dim intervaloColumnas As IXLRange = _excel.Range($"A" & 3 & ":E" & 3)
        intervaloColumnas.Style.Fill.SetBackgroundColor(XLColor.FromArgb(217, 217, 217))
        intervaloColumnas.Style.Font.SetBold(True)
        intervaloColumnas.Style.Border.SetBottomBorderColor(XLColor.Black)
        intervaloColumnas.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)

    End Sub

    Private Shared Sub CrearCabecalhoColumnasDatosFacturacion(ByRef _excel As IXLWorksheet)
        _excel.Cell(2, 1).Value = "PROVEEDOR"
        _excel.Cell(2, 2).Value = "TERMINAL"
        _excel.Cell(2, 3).Value = "PRODUCTO"
        _excel.Cell(2, 4).Value = "NEGOCIO"
        _excel.Cell(2, 5).Value = "ID PEDIDO"
        _excel.Cell(2, 6).Value = "VOL PEDIDO"
        _excel.Cell(2, 7).Value = "PRECIO COSTO"
        _excel.Cell(2, 8).Value = "FLETE COMPRA"
        _excel.Cell(2, 9).Value = "FLETE VENTA"
        _excel.Cell(2, 10).Value = "TERMINAL ENTREGA"
        _excel.Cell(2, 11).Value = "VOL COMPRA DIST"
        _excel.Cell(2, 12).Value = "VOL COMPRA NEGOCIO"
        _excel.Cell(2, 13).Value = "FACTURA"
        _excel.Cell(2, 14).Value = "Nº REMISION"
        _excel.Cell(2, 15).Value = "PEDIMENTO"
        _excel.Cell(2, 16).Value = "FECHA CARGA"
        _excel.Cell(2, 17).Value = "FECHA SUMINISTRO"
        _excel.Cell(2, 18).Value = "COMENTARIOS"

        Dim intervaloColumnas As IXLRange = _excel.Range($"A" & 2 & ":R" & 2)
        intervaloColumnas.Style.Fill.SetBackgroundColor(XLColor.FromArgb(217, 217, 217))
        intervaloColumnas.Style.Font.SetBold(True)
        intervaloColumnas.Style.Border.SetBottomBorderColor(XLColor.Black)

        'For Each item In intervaloColumnas.Columns("A:P")
        '    item.WorksheetColumn().Width = 35
        'Next

        'intervaloColumnas.Columns("A:P").Width = 25
        'intervaloColumnas.Column("O").WorksheetColumn().Width = 25
        intervaloColumnas.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
    End Sub

    Public Shared Sub CrearExcelDatosFacturacion(objdatos As List(Of ObjPedidosFacturacion), nombreArchivo As String, idCliente As String)

        Dim _businessProveedor As New FMBussinesProveedor
        Dim _businessProducto As New FMBussinesProducto
        Dim _businessTerminalSuministro As New FMBussinesTerminalSuministro
        Dim _businessNegocio As New FMBussinesNegocio

        Dim regreso = _businessProveedor.GetProveedoresClienteJoinSQL(idCliente)

        Dim workbook As New XLWorkbook
        Dim worksheet = workbook.Worksheets.Add("Facturacion")
        Dim worksheetProveedor = workbook.Worksheets.Add("Proveedor")
        Dim worksheetTerminal = workbook.Worksheets.Add("Terminal")
        Dim worksheetProducto = workbook.Worksheets.Add("Producto")
        Dim worksheetNegocio = workbook.Worksheets.Add("Negocio")
        Dim rowTablaIndex = 1

        Dim rowTablaIndexA = 1
        For Each item In regreso.AsEnumerable()
            worksheetProveedor.Cell(rowTablaIndexA, 1).Value = item.Item("PRO_NOMCORTO")
            rowTablaIndexA += 1
        Next

        regreso = _businessTerminalSuministro.GetTerminalSuministroClienteJoinSQL(idCliente)
        Dim rowTablaIndexT = 1
        For Each item In regreso.AsEnumerable()
            worksheetTerminal.Cell(rowTablaIndexT, 2).Value = item.Item("TSU_NOMCORTO")
            rowTablaIndexT += 1
        Next

        regreso = _businessProducto.GetProductoJoinSQL(idCliente)
        Dim rowTablaIndexP = 1
        For Each item In regreso.AsEnumerable()
            worksheetProducto.Cell(rowTablaIndexP, 3).Value = item.Item("CON_DESCRIPCION_CORTA")
            rowTablaIndexP += 1
        Next

        regreso = _businessNegocio.GetNegocioCliente(idCliente)
        Dim rowTablaIndexN = 1
        For Each item In regreso.AsEnumerable()
            worksheetNegocio.Cell(rowTablaIndexN, 4).Value = item.Item("CLI_CODNEGOCIO")
            rowTablaIndexN += 1
        Next

        worksheetProveedor.Hide()
        worksheetTerminal.Hide()
        worksheetProducto.Hide()
        worksheetNegocio.Hide()

        worksheet.Column("A").SetDataValidation().List(worksheetProveedor.Range($"A1", $"A{rowTablaIndexA}"), True)
        worksheet.Column("B").SetDataValidation().List(worksheetTerminal.Range($"B1", $"B{rowTablaIndexT}"), True)
        worksheet.Column("J").SetDataValidation().List(worksheetTerminal.Range($"B1", $"B{rowTablaIndexT}"), True)
        worksheet.Column("C").SetDataValidation().List(worksheetProducto.Range($"C1", $"C{rowTablaIndexP}"), True)
        worksheet.Column("D").SetDataValidation().List(worksheetNegocio.Range($"D1", $"D{rowTablaIndexN}"), True)

        worksheet.Cell(rowTablaIndex, 1).Value = "Datos Facturacion"
        Dim Titulo As IXLRange = worksheet.Range(($"A" & rowTablaIndex & ":R" & rowTablaIndex))
        Titulo.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
        Titulo.Merge()
        Titulo.Style.Font.SetBold(True)
        Titulo.Style.Font.SetFontSize(30)
        rowTablaIndex += 1

        CrearCabecalhoColumnasDatosFacturacion(worksheet)
        rowTablaIndex += 1

        For Each item In objdatos
            worksheet.Cell(rowTablaIndex, 1).Value = item.PRO_NOMBRE
            Dim Coluna1 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna1.WorksheetColumn().Width = 15
            Coluna1.Style.Protection.SetHidden(True)
            worksheet.Cell(rowTablaIndex, 2).Value = item.TSU_DESCRIPCION
            Dim Coluna2 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna2.WorksheetColumn().Width = 20
            worksheet.Cell(rowTablaIndex, 3).Value = item.CON_DESCRIPCION_CORTA
            Dim Coluna3 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna3.WorksheetColumn().Width = 15
            worksheet.Cell(rowTablaIndex, 4).Value = item.CLI_DESCRIPCION
            Dim Coluna4 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna4.WorksheetColumn().Width = 20
            worksheet.Cell(rowTablaIndex, 5).Value = item.PED_IDPEDIDOS
            Dim Coluna5 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna5.WorksheetColumn().Width = 10
            worksheet.Cell(rowTablaIndex, 6).Value = item.PRD_VOLUMEN
            Dim Coluna6 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna6.WorksheetColumn().Width = 15
            worksheet.Cell(rowTablaIndex, 7).Value = item.PRD_PRECIO
            Dim Coluna7 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna7.WorksheetColumn().Width = 15
            worksheet.Cell(rowTablaIndex, 8).Value = FormatNumber(0 & item.PRD_FLETE, 6)
            Dim Coluna8 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna8.WorksheetColumn().Width = 15
            worksheet.Cell(rowTablaIndex, 9).Value = item.PRD_FLETE_VENTA
            Dim Coluna9 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna9.WorksheetColumn().Width = 15
            worksheet.Cell(rowTablaIndex, 10).Value = If(item.TERMINALENTREGACODIGO.Equals("VAZIO"), "", item.TERMINALENTREGADESCRICION)
            Dim Coluna10 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna10.WorksheetColumn().Width = 25
            worksheet.Cell(rowTablaIndex, 11).Value = item.VOLCOMPRADIST
            Dim Coluna11 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna11.WorksheetColumn().Width = 15
            worksheet.Cell(rowTablaIndex, 12).Value = item.VOLCOMPRANEGOCIO
            Dim Coluna12 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna12.WorksheetColumn().Width = 15
            worksheet.Cell(rowTablaIndex, 13).Value = item.FACTURA
            Dim Coluna13 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna13.WorksheetColumn().Width = 15
            worksheet.Cell(rowTablaIndex, 14).Value = item.NREMISION
            Dim Coluna14 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna14.WorksheetColumn().Width = 15

            worksheet.Cell(rowTablaIndex, 15).Value = item.FAC_PEDIMENTO
            Dim Coluna15 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna15.WorksheetColumn().Width = 15

            If item.FAC_FECHA_CARGA.Equals("") Then
                worksheet.Cell(rowTablaIndex, 16).Value = ""
            Else
                worksheet.Cell(rowTablaIndex, 16).Value = Format(CDate(item.FAC_FECHA_CARGA), "yyyy/MM/dd").ToString
            End If
            If item.FECHASUMINISTRO.Equals("") Then
                worksheet.Cell(rowTablaIndex, 17).Value = ""
            Else
                worksheet.Cell(rowTablaIndex, 17).Value = Format(CDate(item.FECHASUMINISTRO), "yyyy/MM/dd").ToString
            End If

            Dim Coluna16 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna16.WorksheetColumn().Width = 20

            worksheet.Cell(rowTablaIndex, 18).Value = item.FAC_OBSERVACIONES
            Dim Coluna18 As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            Coluna15.WorksheetColumn().Width = 18

            Dim objDatosLinha = JsonConvert.SerializeObject(item)
            worksheet.Cell(rowTablaIndex, 21).Value = objDatosLinha

            Dim intervaloColumnas As IXLRange = worksheet.Range($"A" & rowTablaIndex, "G" & rowTablaIndex)
            intervaloColumnas.Style.Fill.SetBackgroundColor(XLColor.FromArgb(247, 247, 247))
            Dim intervaloColumnasZ As IXLCell = worksheet.Cell(rowTablaIndex, rowTablaIndex)
            worksheet.Column($"U").Hide()
            rowTablaIndex += 1
        Next

        worksheet.Columns("A", "R").AdjustToContents()

        workbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory & $"{nombreArchivo}.xlsx")
    End Sub

    Private Shared Sub CrearCabecalhoColumnasInformesTar(ByRef _excel As IXLWorksheet)
        _excel.Cell(3, 1).Value = "TERMINAL"
        _excel.Cell(3, 2).Value = "PROVEEDOR"
        _excel.Cell(3, 3).Value = "PRODUCTO"
        _excel.Cell(3, 4).Value = "PERIODO"
        _excel.Cell(3, 5).Value = "VOL. COMPRA"
        _excel.Cell(3, 6).Value = "VOL. COMPROMETIDO"
        _excel.Cell(3, 7).Value = "% DESV"

        Dim intervaloColumnas As IXLRange = _excel.Range($"A" & 3 & ":G" & 3)
        intervaloColumnas.Style.Fill.SetBackgroundColor(XLColor.FromArgb(217, 217, 217))
        intervaloColumnas.Style.Font.SetBold(True)
        intervaloColumnas.Style.Border.SetBottomBorderColor(XLColor.Black)
        intervaloColumnas.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
    End Sub

    Public Shared Sub CrearExcelInformesTAR(ByVal dt As Data.DataTable, ByVal rangoFecha As String)

        Dim workbook As New XLWorkbook
        Dim colIndex As Integer = 1
        Dim rowIndex As Integer = 0
        Dim rowTablaIndex As Integer = 1
        Dim totalDesvio As Double = 0
        Dim totalCompra As Double = 0
        Dim totalObjetivo As Double = 0
        Dim totalTarDesvio As Double = 0
        Dim totalTarCompra As Double = 0
        Dim totalTarObjetivo As Double = 0

        Dim _excel = workbook.Worksheets.Add("InformesTar")
        _excel.Style.NumberFormat.SetFormat("#.000")

        _excel.Cell(rowTablaIndex, 1).Value = "INFORMES TAR"
        Dim Titulo As IXLRange = _excel.Range(($"A" & rowTablaIndex & ":G" & rowTablaIndex))
        Titulo.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
        Titulo.Merge()
        Titulo.Style.Font.SetBold(True)
        Titulo.Style.Font.SetFontSize(30)
        rowTablaIndex += 1

        Titulo = _excel.Range(($"A" & rowTablaIndex & ":G" & rowTablaIndex))
        _excel.Cell(rowTablaIndex, 1).Value = "Fecha: " & rangoFecha
        Titulo.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
        Titulo.Merge()
        Titulo.Style.Font.SetBold(True)
        Titulo.Style.Font.SetFontSize(15)
        rowTablaIndex += 1

        CrearCabecalhoColumnasInformesTar(_excel)
        CrearLinea(_excel, "A1:G1", XLColor.FromArgb(217, 217, 217), True, XLColor.Black, False, False)
        rowTablaIndex += 1

        Dim grupoTar = dt.AsEnumerable().GroupBy(Function(f) f.Item("TAR")).ToList()

        For Each itemTar In grupoTar

            _excel.Cell(rowTablaIndex, colIndex).Value = itemTar.Key
            CrearLinea(_excel, $"A" & rowTablaIndex & ":G" & rowTablaIndex, XLColor.FromArgb(255, 250, 236), False, XLColor.Silver, True, XLAlignmentHorizontalValues.Left)
            rowTablaIndex += 1
            colIndex += 1
            Dim grupoProveedor = itemTar.GroupBy(Function(x) x.Item("PROVEEDOR"))
            Dim totalProveedorDesvio As Double = 0
            Dim totalProveedorCompra As Double = 0
            Dim totalProveedorObjetivo As Double = 0

            For Each itemProveedor In grupoProveedor

                _excel.Cell(rowTablaIndex, colIndex).Value = itemProveedor.Key
                CrearLinea(_excel, ($"B" & rowTablaIndex & ":G" & rowTablaIndex), XLColor.FromArgb(238, 250, 247), False, XLColor.Silver, True, XLAlignmentHorizontalValues.Left)
                rowTablaIndex += 1
                colIndex += 1

                Dim grupoProducto = itemProveedor.GroupBy(Function(x) x.Item("PRODUCTO"))

                For Each itemProducto In grupoProducto

                    _excel.Cell(rowTablaIndex, colIndex).Value = itemProducto.Key
                    CrearLinea(_excel, ($"C" & rowTablaIndex & ":G" & rowTablaIndex), XLColor.FromArgb(235, 235, 246), False, XLColor.Silver, True, XLAlignmentHorizontalValues.Left)
                    rowTablaIndex += 1
                    Dim grupoMes = itemProducto.GroupBy(Function(x) CDate(x.Item("FECHA")).Month)

                    Dim contador = 0
                    For Each itemMes In grupoMes

                        Dim grupoDia = itemMes.GroupBy(Function(x) CDate(x.Item("FECHA")).Day)

                        Dim totalVolumenMes = FormatNumber(itemProducto.Where(Function(p) CDate(p.Item("FECHA")).Month = itemMes.Key).Sum(Function(x) x.Item("PRD_VOLUMEN")), 3)

                        _excel.Cell(rowTablaIndex, 4).Value = "MES " & SiteMaster.RetornaMesEscrito(itemMes.Key.ToString().PadLeft(2, "0"))

                        _excel.Cell(rowTablaIndex, 5).Value = totalVolumenMes

                        CrearLinea(_excel, ($"D" & rowTablaIndex & ":D" & rowTablaIndex), XLColor.FromArgb(255, 255, 241), True, XLColor.Silver, False, XLAlignmentHorizontalValues.Left)

                        CrearLinea(_excel, ($"E" & rowTablaIndex & ":E" & rowTablaIndex), XLColor.FromArgb(255, 255, 241), True, XLColor.Silver, False, XLAlignmentHorizontalValues.Right)
                        _excel.Cell(rowTablaIndex, 5).Value = FormatNumber(itemMes.Sum(Function(x) x.Item("PRD_VOLUMEN")), 3)
                        _excel.Cell(rowTablaIndex, 6).Value = FormatNumber(itemMes.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO")), 3)
                        Dim ColunaF As IXLRange = _excel.Range(($"F" & rowTablaIndex & ":F6"))
                        ColunaF.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right)

                        Dim desvioDia = FormatNumber((((itemMes.Sum(Function(x) x.Item("PRD_VOLUMEN")) / itemMes.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO"))) - 1) * 100), 3)

                        _excel.Cell(rowTablaIndex, 7).Value = desvioDia
                        Dim celulaDesvio1 As IXLRange = _excel.Range(($"G" & rowTablaIndex & ":G" & rowTablaIndex))
                        celulaDesvio1.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right)
                        celulaDesvio1.Style.NumberFormat.SetFormat("#.000")
                        If desvioDia > 0 Then
                            celulaDesvio1.Style.Font.SetFontColor(XLColor.Green)
                        ElseIf desvioDia < 0 Then
                            celulaDesvio1.Style.Font.SetFontColor(XLColor.Red)
                        End If
                        rowTablaIndex += 1

                        If grupoDia.Count() > 0 Then

                            For Each itemDia In grupoDia
                                Dim celuladata As IXLRange = _excel.Range(($"D" & rowTablaIndex & ":D" & rowTablaIndex))
                                celuladata.Style.DateFormat.SetFormat("yyyy/MM/dd")
                                _excel.Cell(rowTablaIndex, 4).Value = Format(CDate(itemDia.FirstOrDefault.Item("FECHA")), "yyyy/MM/dd").ToString()
                                CrearLinea(_excel, ($"D" & rowTablaIndex & ":D" & rowTablaIndex), XLColor.FromArgb(255, 255, 255), False, XLColor.Silver, False, XLAlignmentHorizontalValues.Right)
                                _excel.Cell(rowTablaIndex, 5).Value = FormatNumber(itemDia.Sum(Function(x) x.Item("PRD_VOLUMEN")), 3)
                                CrearLinea(_excel, ($"E" & rowTablaIndex & ":E" & rowTablaIndex), XLColor.FromArgb(255, 255, 255), False, XLColor.Silver, False, XLAlignmentHorizontalValues.Right)
                                rowTablaIndex += 1
                            Next

                        End If
                    Next

                    colIndex = 3
                    _excel.Cell(rowTablaIndex, 3).Value = "TOTAL " & itemProducto.Key
                    _excel.Cell(rowTablaIndex, 5).Value = FormatNumber(itemProducto.Sum(Function(x) x.Item("PRD_VOLUMEN")), 3)
                    _excel.Cell(rowTablaIndex, 6).Value = FormatNumber(itemProducto.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO")), 3)
                    Dim celulaProductoTotalCor As IXLRange = _excel.Range(($"C" & rowTablaIndex & ":G" & rowTablaIndex))
                    celulaProductoTotalCor.Style.Fill.SetBackgroundColor(XLColor.FromArgb(222, 222, 240))
                    Dim desvioProducto = FormatNumber((((itemProducto.Sum(Function(x) x.Item("PRD_VOLUMEN")) / itemProducto.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO"))) - 1) * 100), 3)

                    _excel.Cell(rowTablaIndex, 7).Value = desvioProducto
                    CrearLinea(_excel, ($"E" & rowTablaIndex & ":G" & rowTablaIndex), XLColor.FromArgb(222, 222, 240), False, XLColor.Silver, False, XLAlignmentHorizontalValues.Right)
                    Dim celulaDesvio2 As IXLRange = _excel.Range(($"G" & rowTablaIndex & ":G" & rowTablaIndex))
                    celulaDesvio2.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right)
                    celulaDesvio2.Style.NumberFormat.SetFormat("#.000")
                    If desvioProducto > 0 Then
                        celulaDesvio2.Style.Font.SetFontColor(XLColor.Green)
                    ElseIf desvioProducto < 0 Then
                        celulaDesvio2.Style.Font.SetFontColor(XLColor.Red)
                    End If

                    rowTablaIndex += 1

                Next
                colIndex = 2
                _excel.Cell(rowTablaIndex, 2).Value = "TOTAL" & itemProveedor.Key
                CrearLinea(_excel, ($"B" & rowTablaIndex & ":G" & rowTablaIndex), XLColor.FromArgb(238, 250, 247), False, XLColor.Silver, False, XLAlignmentHorizontalValues.Left)

                _excel.Cell(rowTablaIndex, 5).Value = FormatNumber(itemProveedor.Sum(Function(x) x.Item("PRD_VOLUMEN")), 3)
                _excel.Cell(rowTablaIndex, 6).Value = FormatNumber(itemProveedor.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO")), 3)

                Dim desvioProveedor = FormatNumber((((itemProveedor.Sum(Function(x) x.Item("PRD_VOLUMEN")) / itemProveedor.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO"))) - 1) * 100), 3)
                Dim celulaProveedorTotalCor As IXLRange = _excel.Range(($"B" & rowTablaIndex & ":G" & rowTablaIndex))
                celulaProveedorTotalCor.Style.Fill.SetBackgroundColor(XLColor.FromArgb(216, 235, 231))
                _excel.Cell(rowTablaIndex, 7).Value = desvioProveedor
                CrearLinea(_excel, ($"E" & rowTablaIndex & ":G" & rowTablaIndex), XLColor.FromArgb(216, 235, 231), False, XLColor.Silver, False, XLAlignmentHorizontalValues.Right)
                Dim celulaDesvio3 As IXLRange = _excel.Range(($"G" & rowTablaIndex & ":G" & rowTablaIndex))
                celulaDesvio3.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right)
                celulaDesvio3.Style.NumberFormat.SetFormat("#.000")
                If desvioProveedor > 0 Then
                    celulaDesvio3.Style.Font.SetFontColor(XLColor.Green)
                ElseIf desvioProveedor < 0 Then
                    celulaDesvio3.Style.Font.SetFontColor(XLColor.Red)
                End If
                rowTablaIndex += 1


            Next

            _excel.Cell(rowTablaIndex, 1).Value = "TOTAL" & itemTar.Key
            CrearLinea(_excel, ($"A" & rowTablaIndex & ":A" & rowTablaIndex), XLColor.FromArgb(244, 236, 212), False, XLColor.Silver, False, XLAlignmentHorizontalValues.Left)
            _excel.Cell(rowTablaIndex, 5).Value = FormatNumber(itemTar.Sum(Function(x) x.Item("PRD_VOLUMEN")), 3)
            _excel.Cell(rowTablaIndex, 6).Value = FormatNumber(itemTar.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO")), 3)

            totalCompra += itemTar.Sum(Function(x) x.Item("PRD_VOLUMEN"))
            totalObjetivo += itemTar.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO"))

            Dim desvioTar = FormatNumber((((itemTar.Sum(Function(x) x.Item("PRD_VOLUMEN")) / itemTar.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO"))) - 1) * 100), 3)
            Dim celulaTarTotalCor As IXLRange = _excel.Range(($"A" & rowTablaIndex & ":G" & rowTablaIndex))
            celulaTarTotalCor.Style.Fill.SetBackgroundColor(XLColor.FromArgb(244, 236, 212))
            _excel.Cell(rowTablaIndex, 7).Value = desvioTar
            Dim celulaDesvio As IXLRange = _excel.Range(($"G" & rowTablaIndex & ":G" & rowTablaIndex))

            If desvioTar > 0 Then
                celulaDesvio.Style.Font.SetFontColor(XLColor.Green)
            ElseIf desvioTar < 0 Then
                celulaDesvio.Style.Font.SetFontColor(XLColor.Red)
            End If

            CrearLinea(_excel, ($"B" & rowTablaIndex & ":G" & rowTablaIndex), XLColor.FromArgb(244, 236, 212), False, XLColor.Silver, False, XLAlignmentHorizontalValues.Right)

            rowTablaIndex += 1
            colIndex = 1
        Next

        Dim celulaDesvioGeral As IXLRange = _excel.Range(($"G" & rowTablaIndex & ":G" & rowTablaIndex))

        totalDesvio = FormatNumber((((CDbl(totalCompra) / CDbl(totalObjetivo)) - 1) * 100), 3)

        If totalDesvio > 0 Then
            celulaDesvioGeral.Style.Font.SetFontColor(XLColor.Green)
        ElseIf totalDesvio < 0 Then
            celulaDesvioGeral.Style.Font.SetFontColor(XLColor.Red)
        End If

        CrearLineaTotalizadora(_excel, ($"A" & rowTablaIndex & ":G" & rowTablaIndex), rowTablaIndex, "TOTAL COMPRAS", 1, totalCompra, 5, totalObjetivo, 6, totalDesvio, 7, XLColor.FromArgb(228, 216, 244), ($"A" & rowTablaIndex & ":A" & rowTablaIndex), XLAlignmentHorizontalValues.Left, ($"B" & rowTablaIndex & ":G" & rowTablaIndex), XLAlignmentHorizontalValues.Right)
        rowTablaIndex += 1

        _excel.Columns("A", "G").AdjustToContents()

        workbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory & $"informe.xlsx")
    End Sub


    Private Shared Sub CrearLinea(ByRef _excel As IXLWorksheet, ByVal intervalo As String, ByVal color As XLColor, ByVal negrito As Boolean, ByVal bordacolor As XLColor,
                                 Optional ByVal merge As Boolean = False, Optional align As XLAlignmentHorizontalValues = XLAlignmentHorizontalValues.Center)

        Dim intervaloColumnas As IXLRange = _excel.Range(intervalo)
        intervaloColumnas.Style.Fill.SetBackgroundColor(color)
        intervaloColumnas.Style.Font.SetBold(True)
        intervaloColumnas.Style.Alignment.SetHorizontal(align)
        intervaloColumnas.Style.Border.SetTopBorderColor(bordacolor)
        intervaloColumnas.Style.Border.SetBottomBorderColor(bordacolor)
        intervaloColumnas.Style.Border.SetLeftBorderColor(bordacolor)
        intervaloColumnas.Style.Border.SetRightBorderColor(bordacolor)

        If merge Then
            intervaloColumnas.Merge()
        End If
    End Sub
    Private Shared Sub CrearLineaTotalizadora(ByRef _excel As IXLWorksheet, ByRef intervalo As String, ByVal lineaIndex As Integer,
                                          ByVal columnaTextoTotal As String, ByVal columnaIndexTotal As Integer,
                                          ByVal columnaTextoUm As Double, ByVal columnaIndexUm As Integer,
                                          ByVal columnaTextoDois As Double, ByVal columnaIndexDois As Integer,
                                          ByVal columnaTextoTres As Double, ByVal columnaIndexTres As Integer,
                                          ByVal color As XLColor,
                                          ByVal intervalAlignText As String, ByVal intervalAlignContentText As XLAlignmentHorizontalValues,
                                          ByVal intervalContentText As String, ByVal intervalAlignContent As XLAlignmentHorizontalValues)

        Dim intervaloTotal As IXLRange = _excel.Range(intervalo)
        _excel.Cell(lineaIndex, columnaIndexTotal).Value = columnaTextoTotal
        _excel.Cell(lineaIndex, columnaIndexUm).Value = columnaTextoUm
        _excel.Cell(lineaIndex, columnaIndexDois).Value = columnaTextoDois
        _excel.Cell(lineaIndex, columnaIndexTres).Value = columnaTextoTres
        intervaloTotal.Style.NumberFormat.SetFormat("#.000")
        intervaloTotal.Style.Fill.SetBackgroundColor(color)

        Dim intervaloAlignTotal As IXLRange = _excel.Range(intervalAlignText)
        intervaloAlignTotal.Style.Alignment.SetHorizontal(intervalAlignContentText)
        intervaloAlignTotal.Style.NumberFormat.SetFormat("#.000")

        Dim intervaloContentAlignTotal As IXLRange = _excel.Range(intervalContentText)
        intervaloContentAlignTotal.Style.Alignment.SetHorizontal(intervalAlignContent)
    End Sub

    Public Shared Sub CrearExcelInformesProveedor(dt As Data.DataTable, ByVal rangoFecha As String)
        Dim workbook As New XLWorkbook
        Dim colIndex As Integer = 7
        Dim rowIndex As Integer = 0
        Dim rowTablaIndex As Integer = 1

        Dim _excel = workbook.Worksheets.Add("InformesProveedores")
        _excel.Worksheet.Style.NumberFormat.SetFormat("#.000")

        _excel.Cell(rowTablaIndex, 1).Value = "INFORMES PROVEEDOR"
        Dim Titulo As IXLRange = _excel.Range(($"A" & rowTablaIndex & ":H" & rowTablaIndex))
        Titulo.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
        Titulo.Merge()
        Titulo.Style.Font.SetBold(True)
        Titulo.Style.Font.SetFontSize(30)
        rowTablaIndex += 1

        Titulo = _excel.Range(($"A" & rowTablaIndex & ":H" & rowTablaIndex))
        _excel.Cell(rowTablaIndex, 1).Value = "Año: " & rangoFecha
        Titulo.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left)
        Titulo.Merge()
        Titulo.Style.Font.SetBold(True)
        Titulo.Style.Font.SetFontSize(15)
        rowTablaIndex += 1

        CrearCabecalhoColumnasProveedor(_excel)
        CrearLinea(_excel, $"A3:H3", XLColor.FromArgb(217, 217, 217), True, XLColor.Black, False, XLAlignmentHorizontalValues.Center)
        rowTablaIndex += 1
        Dim grupoProveedor = dt.AsEnumerable().GroupBy(Function(x) x.Item("PROVEEDOR"))

        Dim somatoriaVolCompraTotalGeral As Double = 0
        Dim somatoriaVolObjetivoTotalGeral As Double = 0
        Dim somatoriaDesvProveedorTotalGeral As Double = 0
        Dim linhaFin As String = ""
        For Each provedor In grupoProveedor
            Dim celulaValores As IXLRange = _excel.Range(($"F" & rowTablaIndex & ":H" & rowTablaIndex))

            _excel.Cell(rowTablaIndex, 1).Value = provedor.Key
            CrearLinea(_excel, ($"A" & rowTablaIndex & ":H" & rowTablaIndex), XLColor.FromArgb(255, 250, 236), False, XLColor.Silver, True, XLAlignmentHorizontalValues.Left)
            rowTablaIndex += 1

            Dim somatoriaVolCompraTotal As Double = provedor.Sum(Function(x) x.Item("PRD_VOLUMEN"))
            Dim somatoriaVolObjetivoTotal As Double = provedor.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO"))
            Dim somatoriaDesvProveedorTotal As Double = (somatoriaVolCompraTotal / somatoriaVolObjetivoTotal) - 1
            Dim linhaProveedor = ""
            Dim idProveedor = provedor.FirstOrDefault().Item("CODIGOPROVEEDOR")
            Dim nomeCluster = "PRO-" & idProveedor

            somatoriaVolCompraTotalGeral += somatoriaVolCompraTotal
            somatoriaVolObjetivoTotalGeral += somatoriaVolObjetivoTotal
            somatoriaDesvProveedorTotalGeral += somatoriaDesvProveedorTotal

            Dim grupoCluster = provedor.AsEnumerable().GroupBy(Function(x) x.Item("CLU_DESCRIPCION")).OrderBy(Function(x) x.Key)
            Dim linhaCluster = ""
            For Each itemCluster In grupoCluster

                Dim somatoriaVolCompraCluster As Double = itemCluster.Sum(Function(x) x.Item("PRD_VOLUMEN"))
                Dim somatoriaVolObjetivoCluster As Double = itemCluster.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO"))
                Dim somatoriaDesvProveedorCluster As Double = (somatoriaVolCompraCluster / somatoriaVolObjetivoCluster) - 1
                Dim idCluster = itemCluster.FirstOrDefault().Item("CLU_IDCLUSTER")
                Dim nomeTar = "CLU-" & idCluster & "-PRO-" & idProveedor
                Dim totalComprometidoProveedor = 0
                _excel.Cell(rowTablaIndex, 2).Value = itemCluster.Key
                CrearLinea(_excel, ($"B" & rowTablaIndex & ":H" & rowTablaIndex), XLColor.FromArgb(210, 236, 246), False, XLColor.Silver, True, XLAlignmentHorizontalValues.Left)
                rowTablaIndex += 1

                Dim grupoTar = itemCluster.AsEnumerable().GroupBy(Function(x) x.Item("TSU_DESCRIPCION"))
                Dim linhaTar As String = ""
                For Each itemTar In grupoTar

                    Dim linhaProducto As String = ""
                    Dim somatoriaTar As Double = 0
                    Dim somatoriaVolCompraTAR As Double = itemTar.Sum(Function(x) x.Item("PRD_VOLUMEN"))
                    Dim somatoriaVolObjetivoTAR As Double = itemTar.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO"))
                    Dim somatoriaDesvProveedorTAR As Double = (somatoriaVolCompraTAR / somatoriaVolObjetivoTAR) - 1
                    Dim idTar = itemTar.FirstOrDefault().Item("CODIGOTAR").ToString()
                    _excel.Cell(rowTablaIndex, 3).Value = itemTar.Key
                    CrearLinea(_excel, ($"C" & rowTablaIndex & ":H" & rowTablaIndex), XLColor.FromArgb(220, 239, 235), False, XLColor.Silver, True, XLAlignmentHorizontalValues.Left)
                    rowTablaIndex += 1


                    Dim nomeProducto = "PRV-" & itemCluster.FirstOrDefault.Item("CODIGOPROVEEDOR").ToString() & "-CLU-" & itemCluster.FirstOrDefault.Item("CLU_IDCLUSTER").ToString() & "-TAR-" & itemTar.FirstOrDefault.Item("CODIGOTAR").ToString()
                    Dim GrupoProducto = itemTar.AsEnumerable().GroupBy(Function(x) x.Item("PRODUCTO")).OrderBy(Function(x) x.Key)

                    For Each producto In GrupoProducto
                        Dim linhaMes = ""
                        Dim totalVolumenesProductos As Double = 0
                        Dim totalObjetivoProductos As Double = 0

                        Dim descripcionProducto = producto.Key.Replace("<", "‹") _
                                                             .Replace(">", "›")
                        Dim idProduto = producto.FirstOrDefault().Item("CODIGOPRODUTO").ToString()
                        Dim nomeMes = "CON-" & idProduto & "-TAR-" & idTar & "-CLU-" & idCluster & "-PRO-" & idProveedor
                        descripcionProducto = descripcionProducto.Replace("<", "‹") _
                                                             .Replace(">", "›")

                        _excel.Cell(rowTablaIndex, 4).Value = producto.Key
                        CrearLinea(_excel, ($"D" & rowTablaIndex & ":H" & rowTablaIndex), XLColor.FromArgb(228, 228, 239), False, XLColor.Silver, True, XLAlignmentHorizontalValues.Left)
                        rowTablaIndex += 1

                        'agrupa los productos por fecha
                        Dim fechaProductos = producto.GroupBy(Function(x) CDate(x.Item("FECHA")).Month()).Distinct().ToList()

                        For Each fechaMes In fechaProductos
                            Dim totalObjetivoMes As Double = 0
                            Dim totalVolumenMes As Double = 0
                            'Dim fechaProducto As String = ""
                            Dim linhaDia = ""
                            Dim nomeDia = "MES-" & fechaMes.Key & "-CON-" & idProduto & "-TAR-" & idTar & "-CLU-" & idCluster & "-PRO-" & idProveedor
                            'escribe el encabezamiento de lo mes
                            Dim desvio = (((fechaMes.Sum(Function(x) x.Item("PRD_VOLUMEN")) / fechaMes.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO"))) - 1) * 100)
                            _excel.Cell(rowTablaIndex, 5).Value = "MES " & SiteMaster.RetornaMesEscrito(fechaMes.Key.ToString().PadLeft(2, "0"))
                            _excel.Cell(rowTablaIndex, 6).Value = FormatNumber(fechaMes.Sum(Function(x) x.Item("PRD_VOLUMEN")), 3)
                            _excel.Cell(rowTablaIndex, 7).Value = FormatNumber(fechaMes.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO")), 3)
                            _excel.Cell(rowTablaIndex, 8).Value = FormatNumber(desvio, 3)
                            Dim celulaMesCor As IXLRange = _excel.Range(($"E" & rowTablaIndex & ":H" & rowTablaIndex))
                            celulaMesCor.Style.Fill.SetBackgroundColor(XLColor.FromArgb(248, 248, 234))

                            celulaValores = _excel.Range(($"F" & rowTablaIndex & ":H" & rowTablaIndex))
                            celulaValores.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right)

                            Dim celulafimMes As IXLRange = _excel.Range(($"H" & rowTablaIndex & ":H" & rowTablaIndex))
                            If desvio > 0 Then
                                celulafimMes.Style.Font.SetFontColor(XLColor.Green)
                            ElseIf somatoriaDesvProveedorTotalGeral < 0 Then
                                celulafimMes.Style.Font.SetFontColor(XLColor.Red)
                            End If

                            rowTablaIndex += 1

                            Dim grupodia = fechaMes.GroupBy(Function(x) CDate(x.Item("FECHA")).Day()).Distinct().ToList()
                            For Each fechaDia In grupodia
                                'escribe los dias
                                Dim sumVolumenProducto = fechaDia.Sum(Function(x) x.Item("PRD_VOLUMEN"))
                                Dim volumenObjetivo = fechaDia.Sum(Function(x) x.Item("PVO_VOLUMENCOMPROMETIDO"))
                                totalVolumenesProductos += sumVolumenProducto
                                totalObjetivoMes += volumenObjetivo
                                totalVolumenMes += sumVolumenProducto
                                Dim adiccionales = $"<tr name=""{nomeDia}"" style=""display: none; background-color: transparent;"">"

                                _excel.Cell(rowTablaIndex, 5).Value = fechaDia.FirstOrDefault().Item("FECHA")
                                Dim celuladata As IXLRange = _excel.Range(($"D" & rowTablaIndex & ":E" & rowTablaIndex))
                                celuladata.Style.DateFormat.SetFormat("yyyy/MM/dd")

                                celulaValores = _excel.Range(($"F" & rowTablaIndex & ":F" & rowTablaIndex))
                                celulaValores.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right)

                                _excel.Cell(rowTablaIndex, 6).Value = FormatNumber(sumVolumenProducto, 3)
                                rowTablaIndex += 1
                            Next
                            totalObjetivoProductos += +totalObjetivoMes
                        Next
                        'escribe el encabezamiento
                        _excel.Cell(rowTablaIndex, 4).Value = "TOTAL " & producto.Key
                        _excel.Cell(rowTablaIndex, 6).Value = FormatNumber(producto.Sum(Function(x) x.Item("PRD_VOLUMEN")), 3).ToString()
                        _excel.Cell(rowTablaIndex, 7).Value = FormatNumber(totalObjetivoProductos, 3)
                        _excel.Cell(rowTablaIndex, 8).Value = FormatNumber((((totalVolumenesProductos / totalObjetivoProductos) - 1) * 100), 3)
                        Dim celulaProductoTotalCor As IXLRange = _excel.Range(($"D" & rowTablaIndex & ":H" & rowTablaIndex))
                        celulaProductoTotalCor.Style.Fill.SetBackgroundColor(XLColor.FromArgb(204, 204, 222))

                        celulaValores = _excel.Range(($"F" & rowTablaIndex & ":H" & rowTablaIndex))
                        celulaValores.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right)

                        Dim celulafimProd As IXLRange = _excel.Range(($"H" & rowTablaIndex & ":H" & rowTablaIndex))

                        If (((totalVolumenesProductos / totalObjetivoProductos) - 1) * 100) > 0 Then
                            celulafimProd.Style.Font.SetFontColor(XLColor.Green)
                        ElseIf somatoriaDesvProveedorTotalGeral < 0 Then
                            celulafimProd.Style.Font.SetFontColor(XLColor.Red)
                        End If
                        rowTablaIndex += 1
                        totalComprometidoProveedor += totalObjetivoProductos
                    Next

                    _excel.Cell(rowTablaIndex, 3).Value = "TOTAL " & itemTar.Key
                    _excel.Cell(rowTablaIndex, 6).Value = FormatNumber(somatoriaVolCompraTAR, 3).ToString()
                    _excel.Cell(rowTablaIndex, 7).Value = FormatNumber(somatoriaVolObjetivoTAR, 3).ToString()
                    _excel.Cell(rowTablaIndex, 8).Value = FormatNumber(somatoriaDesvProveedorTAR, 3).ToString()
                    Dim celulaTarTotalCor As IXLRange = _excel.Range(($"C" & rowTablaIndex & ":H" & rowTablaIndex))
                    celulaTarTotalCor.Style.Fill.SetBackgroundColor(XLColor.FromArgb(183, 228, 219))

                    celulaValores = _excel.Range(($"F" & rowTablaIndex & ":H" & rowTablaIndex))
                    celulaValores.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right)

                    Dim celulafimTar As IXLRange = _excel.Range(($"H" & rowTablaIndex & ":H" & rowTablaIndex))

                    If somatoriaDesvProveedorTAR > 0 Then
                        celulafimTar.Style.Font.SetFontColor(XLColor.Green)
                    ElseIf somatoriaDesvProveedorTotalGeral < 0 Then
                        celulafimTar.Style.Font.SetFontColor(XLColor.Red)
                    End If
                    rowTablaIndex += 1
                Next

                _excel.Cell(rowTablaIndex, 2).Value = "TOTAL " & itemCluster.Key
                _excel.Cell(rowTablaIndex, 6).Value = FormatNumber(somatoriaVolCompraCluster, 3).ToString()
                _excel.Cell(rowTablaIndex, 7).Value = FormatNumber(somatoriaVolObjetivoCluster, 3).ToString()
                _excel.Cell(rowTablaIndex, 8).Value = FormatNumber(somatoriaDesvProveedorCluster, 3).ToString()
                Dim celulaClusterTotalCor As IXLRange = _excel.Range(($"B" & rowTablaIndex & ":H" & rowTablaIndex))
                celulaClusterTotalCor.Style.Fill.SetBackgroundColor(XLColor.FromArgb(163, 216, 235))

                celulaValores = _excel.Range(($"F" & rowTablaIndex & ":H" & rowTablaIndex))
                celulaValores.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right)

                Dim celulafimC As IXLRange = _excel.Range(($"H" & rowTablaIndex & ":H" & rowTablaIndex))

                If somatoriaDesvProveedorCluster > 0 Then
                    celulafimC.Style.Font.SetFontColor(XLColor.Green)
                ElseIf somatoriaDesvProveedorTotalGeral < 0 Then
                    celulafimC.Style.Font.SetFontColor(XLColor.Red)
                End If
                rowTablaIndex += 1
            Next
            _excel.Cell(rowTablaIndex, 1).Value = "TOTAL " & provedor.Key
            _excel.Cell(rowTablaIndex, 6).Value = FormatNumber(somatoriaVolCompraTotal, 3).ToString()
            _excel.Cell(rowTablaIndex, 7).Value = FormatNumber(somatoriaVolObjetivoTotal, 3).ToString()
            _excel.Cell(rowTablaIndex, 8).Value = FormatNumber(somatoriaDesvProveedorTotal, 3).ToString()
            Dim celulafimCor As IXLRange = _excel.Range(($"A" & rowTablaIndex & ":H" & rowTablaIndex))
            celulafimCor.Style.Fill.SetBackgroundColor(XLColor.FromArgb(248, 240, 216))

            celulaValores = _excel.Range(($"F" & rowTablaIndex & ":H" & rowTablaIndex))
            celulaValores.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Right)

            Dim celulafim As IXLRange = _excel.Range(($"H" & rowTablaIndex & ":H" & rowTablaIndex))
            If somatoriaDesvProveedorTotal > 0 Then
                celulafim.Style.Font.SetFontColor(XLColor.Green)
            ElseIf somatoriaDesvProveedorTotalGeral < 0 Then
                celulafim.Style.Font.SetFontColor(XLColor.Red)
            End If
            rowTablaIndex += 1
        Next

        Dim celulaDesvioGeral As IXLRange = _excel.Range(($"H" & rowTablaIndex & ":H" & rowTablaIndex))

        If somatoriaDesvProveedorTotalGeral > 0 Then
            celulaDesvioGeral.Style.Font.SetFontColor(XLColor.Green)
        ElseIf somatoriaDesvProveedorTotalGeral < 0 Then
            celulaDesvioGeral.Style.Font.SetFontColor(XLColor.Red)
        End If

        CrearLineaTotalizadora(_excel, ($"A" & rowTablaIndex & ":H" & rowTablaIndex), rowTablaIndex, "TOTAL GERAL COMPRAS", 1, somatoriaVolCompraTotalGeral.ToString(), 6, somatoriaVolObjetivoTotalGeral.ToString(), 7, somatoriaDesvProveedorTotalGeral.ToString(), 8, XLColor.FromArgb(228, 216, 244), ($"A" & rowTablaIndex & ":A" & rowTablaIndex), XLAlignmentHorizontalValues.Left, ($"B" & rowTablaIndex & ":G" & rowTablaIndex), XLAlignmentHorizontalValues.Right)
        rowTablaIndex += 1

        _excel.Columns("A", "G").AdjustToContents()

        workbook.SaveAs(AppDomain.CurrentDomain.BaseDirectory & $"informeProveedor.xlsx")
    End Sub
    Private Shared Sub CrearCabecalhoColumnasProveedor(ByRef _excel As IXLWorksheet)
        _excel.Cell(3, 1).Value = "PROVEEDOR"
        _excel.Cell(3, 2).Value = "CLUSTER"
        _excel.Cell(3, 3).Value = "TERMINAL"
        _excel.Cell(3, 4).Value = "PRODUCTO"
        _excel.Cell(3, 5).Value = "PERIODO"
        _excel.Cell(3, 6).Value = "VOL. COMPRA"
        _excel.Cell(3, 7).Value = "VOL. OBJETIVO"
        _excel.Cell(3, 8).Value = "% DESV"

        Dim intervaloColumnas As IXLRange = _excel.Range(($"A" & 3 & ":H" & 3))
        intervaloColumnas.Style.Fill.SetBackgroundColor(XLColor.FromArgb(217, 217, 217))
        intervaloColumnas.Style.Font.SetBold(True)
        intervaloColumnas.Style.Border.SetBottomBorderColor(XLColor.Black)
        intervaloColumnas.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
    End Sub
End Class


