Attribute VB_Name = "modCasaOrganizada"
' Casa Organizada - ERP de gestão familiar (painel, análise, lançamentos, contas e metas)
Option Explicit

Public gOcupado As Boolean
Private gAgendado As Date

Private Const SH_DASH As String = "Painel"
Private Const SH_ANA As String = "Análise"
Private Const SH_CAD As String = "Cadastro"
Private Const SH_CON As String = "Contas"
Private Const SH_MET As String = "Metas"
Private Const SH_LANC As String = "Lançamentos"
Private Const SH_ORC As String = "Orçamento"
Private Const SH_REND As String = "Rendas"
Private Const SH_CAT As String = "Categorias"
Private Const SH_CALC As String = "Calc"
Private Const SH_PIV As String = "Dinâmicas"
Private Const FONTE As String = "Segoe UI"
Private Const LARG As Single = 1235
Private Const Y0 As Single = 84           ' início do conteúdo abaixo do cabeçalho + navegação

Private ErroMontagem As String

' paleta executiva
Private cHeader As Long, cGold As Long, cFundo As Long, cCard As Long, cTexto As Long, cMuted As Long, cLinha As Long
Private cPrim As Long, cSec As Long, cTeal As Long, cPrev As Long, cBom As Long, cRuim As Long, cAlerta As Long, cInput As Long, cSuave As Long

Private Sub Cores()
    cHeader = RGB(11, 31, 58): cGold = RGB(201, 162, 39)
    cFundo = RGB(244, 246, 249): cCard = RGB(255, 255, 255): cLinha = RGB(226, 232, 240)
    cTexto = RGB(26, 32, 44): cMuted = RGB(100, 116, 139): cSuave = RGB(241, 245, 249)
    cPrim = RGB(42, 111, 184): cSec = RGB(200, 110, 23): cTeal = RGB(17, 145, 127): cPrev = RGB(184, 194, 204)
    cBom = RGB(46, 125, 50): cRuim = RGB(198, 40, 40): cAlerta = RGB(183, 121, 31): cInput = RGB(248, 250, 252)
End Sub

'===============================================================================
' MONTAGEM (idempotente: pode rodar de novo sem perder dados)
'===============================================================================
Public Sub ConfigurarPainel()
    If Not MontarPainel() Then MsgBox "Erro ao montar: " & ErroMontagem, vbCritical, "Casa Organizada"
End Sub

Public Function ConfigurarPainelAuto() As String
    If Not MontarPainel() Then ConfigurarPainelAuto = ErroMontagem
End Function

Private Function MontarPainel() As Boolean
    Dim etapa As String
    On Error GoTo erro
    Cores
    gOcupado = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    etapa = "estilos": CriarEstilos
    etapa = "calc": PrepararCalc
    etapa = "limpeza": RemoverObsoletos
    etapa = "categorias": PrepararCategorias
    etapa = "contas": PrepararContas
    etapa = "metas": PrepararMetas
    etapa = "apagar": ApagarAba SH_DASH: ApagarAba "Dashboard": ApagarAba SH_ANA: ApagarSegmentacoes: ApagarAba SH_PIV
    etapa = "dinâmicas": CriarDinamicas
    etapa = "cálculos": PreencherCalc
    etapa = "painel": CriarPainel
    etapa = "análise": CriarAnalise
    etapa = "cadastro": CriarCadastro
    etapa = "tela contas": TelaContas
    etapa = "tela metas": TelaMetas
    etapa = "abas de dados": ArrumarAbasDeDados
    etapa = "ordem": OrdenarAbas
    etapa = "eventos": InstalarEventos

    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    gOcupado = False
    AtualizarPainel
    Application.ScreenUpdating = True
    ThisWorkbook.ShowPivotTableFieldList = False
    ThisWorkbook.ShowPivotChartActiveFields = False
    IrDashboard
    MontarPainel = True
    Exit Function
erro:
    ErroMontagem = "[" & etapa & "] " & Err.Description & " (" & Err.Number & ")"
    gOcupado = False
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Function

' Estilos próprios de tabela, tabela dinâmica e segmentação (identidade visual única).
Private Sub CriarEstilos()
    Dim ts As TableStyle
    On Error Resume Next

    Set ts = Nothing: Set ts = ThisWorkbook.TableStyles("CO Tabela")
    If ts Is Nothing Then Set ts = ThisWorkbook.TableStyles.Add("CO Tabela")
    ts.ShowAsAvailableTableStyle = True
    ts.ShowAsAvailablePivotTableStyle = True
    ts.ShowAsAvailableSlicerStyle = False
    ts.TableStyleElements(xlWholeTable).Clear
    ts.TableStyleElements(xlWholeTable).Borders(xlEdgeBottom).Color = cLinha
    ts.TableStyleElements(xlWholeTable).Borders(xlInsideHorizontal).Color = cLinha
    With ts.TableStyleElements(xlHeaderRow)
        .Clear
        .Interior.Color = cHeader
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
    End With
    With ts.TableStyleElements(xlRowStripe1)
        .Clear
        .Interior.Color = cInput
    End With
    With ts.TableStyleElements(xlTotalRow)
        .Clear
        .Interior.Color = cSuave
        .Font.Bold = True
    End With
    With ts.TableStyleElements(xlGrandTotalRow)
        .Clear
        .Interior.Color = RGB(226, 232, 240)
        .Font.Bold = True
        .Font.Color = cTexto
    End With
    With ts.TableStyleElements(xlGrandTotalColumn)
        .Clear
        .Interior.Color = cSuave
        .Font.Bold = True
    End With
    With ts.TableStyleElements(xlSubtotalRow1)
        .Clear
        .Font.Bold = True
        .Interior.Color = cSuave
    End With
    With ts.TableStyleElements(xlRowSubheading1)
        .Clear
        .Font.Bold = True
        .Font.Color = cHeader
    End With

    Set ts = Nothing: Set ts = ThisWorkbook.TableStyles("CO Filtro")
    If ts Is Nothing Then Set ts = ThisWorkbook.TableStyles.Add("CO Filtro")
    ts.ShowAsAvailableSlicerStyle = True
    ts.ShowAsAvailableTableStyle = False
    ts.ShowAsAvailablePivotTableStyle = False
    With ts.TableStyleElements(xlHeaderRow)
        .Clear
        .Font.Bold = True
        .Font.Color = cTexto
    End With
    With ts.TableStyleElements(xlSlicerSelectedItemWithData)
        .Clear
        .Interior.Color = cPrim
        .Font.Color = RGB(255, 255, 255)
    End With
    With ts.TableStyleElements(xlSlicerUnselectedItemWithData)
        .Clear
        .Interior.Color = RGB(255, 255, 255)
        .Font.Color = cTexto
        .Borders(xlEdgeTop).Color = cLinha: .Borders(xlEdgeBottom).Color = cLinha
        .Borders(xlEdgeLeft).Color = cLinha: .Borders(xlEdgeRight).Color = cLinha
    End With
    With ts.TableStyleElements(xlSlicerSelectedItemWithNoData)
        .Clear
        .Interior.Color = RGB(203, 213, 225)
        .Font.Color = cMuted
    End With
    With ts.TableStyleElements(xlSlicerUnselectedItemWithNoData)
        .Clear
        .Interior.Color = cSuave
        .Font.Color = RGB(160, 170, 185)
    End With
    With ts.TableStyleElements(xlSlicerHoveredSelectedItemWithData)
        .Clear
        .Interior.Color = cHeader
        .Font.Color = RGB(255, 255, 255)
    End With
    With ts.TableStyleElements(xlSlicerHoveredUnselectedItemWithData)
        .Clear
        .Interior.Color = cSuave
        .Font.Color = cTexto
    End With
End Sub

' Aba oculta com parâmetros, listas e cálculos.
Private Sub PrepararCalc()
    Dim wc As Worksheet, ano As Variant, dia As Variant
    On Error Resume Next
    ano = ThisWorkbook.Names("Ano").RefersToRange.Value
    Set wc = ThisWorkbook.Worksheets(SH_CALC)
    If Not wc Is Nothing Then dia = wc.Range("B3").Value
    On Error GoTo 0
    If IsEmpty(ano) Or Not IsNumeric(ano) Then ano = Year(Date)
    If IsEmpty(dia) Or Not IsNumeric(dia) Or dia < 1 Or dia > 28 Then dia = 10
    If wc Is Nothing Then
        Set wc = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wc.Name = SH_CALC
    End If
    wc.Cells.Clear
    wc.Range("A1").Value = "Ano": wc.Range("B1").Value = ano
    wc.Range("A3").Value = "Dia de vencimento padrão": wc.Range("B3").Value = dia
    ThisWorkbook.Names.Add Name:="Ano", RefersTo:="='" & SH_CALC & "'!$B$1"
    ThisWorkbook.Names.Add Name:="DiaVenc", RefersTo:="='" & SH_CALC & "'!$B$3"
    wc.Range("H1:H6").Value = Application.WorksheetFunction.Transpose(Array("Pix", "Débito", "Crédito", "Dinheiro", "Boleto", "Transferência"))
    wc.Range("I1:I3").Value = Application.WorksheetFunction.Transpose(Array("Matheus", "Jhenis", "Casa"))
    wc.Range("J1:J2").Value = Application.WorksheetFunction.Transpose(Array("Aberta", "Paga"))
    wc.Range("K1:K2").Value = Application.WorksheetFunction.Transpose(Array("Mensal", "Única"))
    wc.Range("L1:L2").Value = Application.WorksheetFunction.Transpose(Array("Sim", "Não"))
    wc.Visible = xlSheetHidden
End Sub

Private Sub RemoverObsoletos()
    Dim nm As Variant, ws As Worksheet, lo As ListObject
    For Each nm In Array("Instruções", "Qualidade", "Log de Alterações", "Painel Mensal", "Resumo Anual")
        ApagarAba CStr(nm)
    Next nm
    On Error Resume Next
    ThisWorkbook.Worksheets(SH_LANC).Range("A2").ClearContents
    ThisWorkbook.Worksheets(SH_ORC).Range("A2").ClearContents
    Set ws = ThisWorkbook.Worksheets(SH_REND)
    ws.Range("G3:K8").UnMerge: ws.Range("G3:K8").Clear
    Set lo = ThisWorkbook.Worksheets(SH_CAT).ListObjects("tbCategorias")
    lo.ListColumns("Observação").Delete
    ThisWorkbook.Names("gCat").Delete: ThisWorkbook.Names("gPrev").Delete: ThisWorkbook.Names("gReal").Delete
    On Error GoTo 0
End Sub

' Categorias: coluna Lista (para escolher no formulário) e coluna Fixa (gera contas a pagar).
Private Sub PrepararCategorias()
    Dim lo As ListObject, h As Range, v As Variant, i As Long, fixas As Variant, novaFixa As Boolean
    Set lo = ThisWorkbook.Worksheets(SH_CAT).ListObjects("tbCategorias")
    If Not ColunaExiste(lo, "Lista") Then lo.ListColumns.Add.Name = "Lista"
    lo.ListColumns("Lista").DataBodyRange.Formula = "=[@Código]&"" - ""&[@Subcategoria]"
    If Not ColunaExiste(lo, "Fixa") Then lo.ListColumns.Add(lo.ListColumns("Lista").Index).Name = "Fixa": novaFixa = True
    If novaFixa Then
        fixas = Array(2, 3, 4, 6, 7, 8, 9, 15, 16, 17, 18, 19, 20, 21, 55)
        v = lo.ListColumns("Código").DataBodyRange.Value
        Dim saida() As Variant: ReDim saida(1 To UBound(v, 1), 1 To 1)
        For i = 1 To UBound(v, 1)
            saida(i, 1) = IIf(IsNumeric(Application.Match(v(i, 1), fixas, 0)), "Sim", "Não")
        Next i
        lo.ListColumns("Fixa").DataBodyRange.Value = saida
    End If
    With lo.ListColumns("Fixa").DataBodyRange.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="='" & SH_CALC & "'!$L$1:$L$2"
    End With
    Set h = ThisWorkbook.Worksheets(SH_CAT).Rows(3).Find(What:="Grupos", LookIn:=xlValues, LookAt:=xlWhole)
    If h Is Nothing Then Err.Raise vbObjectError + 1, , "Lista de grupos não encontrada na aba Categorias."
    PrepararRegras h
    ThisWorkbook.Names.Add Name:="ListaCategorias", RefersTo:="=tbCategorias[Lista]"
    ThisWorkbook.Names.Add Name:="ListaGrupos", RefersTo:="='" & SH_CAT & "'!" & h.Offset(1, 0).Resize(37, 1).Address
End Sub

' Contas a pagar: cria a aba e a tabela na primeira vez (dados preservados nas próximas).
Private Sub PrepararContas()
    Dim ws As Worksheet, lo As ListObject
    Set ws = PegarOuCriarAba(SH_CON)
    On Error Resume Next
    Set lo = ws.ListObjects("tbContas")
    On Error GoTo 0
    If lo Is Nothing Then
        ws.Range("B11:J11").Value = Array("Vencimento", "Descrição", "Código", "Categoria", "Valor", "Recorrência", "Status", "Pago em", "Situação")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("B11:J12"), , xlYes)
        lo.Name = "tbContas"
    End If
    lo.ListColumns("Categoria").DataBodyRange.Formula = "=IF([@Código]="""","""",IFERROR(INDEX(tbCategorias[Subcategoria],MATCH([@Código],tbCategorias[Código],0)),""CÓDIGO INVÁLIDO""))"
    lo.ListColumns("Situação").DataBodyRange.Formula = "=IF([@Descrição]="""","""",IF([@Status]=""Paga"",""Paga"",IF([@Vencimento]<TODAY(),""Vencida"",IF([@Vencimento]-TODAY()<=7,""Vence em 7 dias"",""A vencer""))))"
    lo.ListColumns("Vencimento").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    lo.ListColumns("Pago em").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    lo.ListColumns("Valor").DataBodyRange.NumberFormat = """R$ ""#,##0.00"
    With lo.ListColumns("Status").DataBodyRange.Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="='" & SH_CALC & "'!$J$1:$J$2"
    End With
    With lo.ListColumns("Recorrência").DataBodyRange.Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="='" & SH_CALC & "'!$K$1:$K$2"
    End With
    With lo.ListColumns("Código").DataBodyRange.Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=INDIRECT(""tbCategorias[Código]"")"
    End With
    lo.TableStyle = "CO Tabela"
End Sub

' Metas: vinculadas a uma categoria de reserva/investimento; o acumulado vem dos lançamentos.
Private Sub PrepararMetas()
    Dim ws As Worksheet, lo As ListObject, codigos As Variant, i As Long, lc As ListObject, r As Variant, alvo As Double, n As Long
    Set ws = PegarOuCriarAba(SH_MET)
    On Error Resume Next
    Set lo = ws.ListObjects("tbMetas")
    On Error GoTo 0
    If lo Is Nothing Then
        ws.Range("B11:J11").Value = Array("Meta", "Código", "Alvo", "Prazo", "Acumulado", "% concluído", "Falta", "Aporte mensal", "Situação")
        Set lo = ws.ListObjects.Add(xlSrcRange, ws.Range("B11:J12"), , xlYes)
        lo.Name = "tbMetas"
        ' metas iniciais = categorias de reserva/investimento com orçamento; alvo = previsto do ano
        Set lc = ThisWorkbook.Worksheets(SH_CAT).ListObjects("tbCategorias")
        codigos = Array(1, 5, 34, 35, 36, 37)
        For i = 0 To UBound(codigos)
            alvo = Application.WorksheetFunction.SumIfs(ThisWorkbook.Worksheets(SH_ORC).Range("P5:P300"), ThisWorkbook.Worksheets(SH_ORC).Range("A5:A300"), codigos(i))
            r = Application.Match(codigos(i), lc.ListColumns("Código").DataBodyRange, 0)
            If alvo > 0 And Not IsError(r) Then
                n = n + 1
                If n > 1 Then lo.ListRows.Add
                With lo.ListRows(n).Range
                    .Cells(1, 1).Value = lc.ListColumns("Subcategoria").DataBodyRange.Cells(r, 1).Value
                    .Cells(1, 2).Value = codigos(i)
                    .Cells(1, 3).Value = alvo
                    .Cells(1, 4).Value = DateSerial(ThisWorkbook.Names("Ano").RefersToRange.Value, 12, 31)
                End With
            End If
        Next i
    End If
    lo.ListColumns("Acumulado").DataBodyRange.Formula = "=IF([@Código]="""",0,SUMIFS(tbLancamentos[Valor],tbLancamentos[Código],[@Código]))"
    lo.ListColumns("% concluído").DataBodyRange.Formula = "=IFERROR(MIN(1,[@Acumulado]/[@Alvo]),0)"
    lo.ListColumns("Falta").DataBodyRange.Formula = "=MAX(0,N([@Alvo])-[@Acumulado])"
    lo.ListColumns("Aporte mensal").DataBodyRange.Formula = "=IF([@Meta]="""",0,IF([@Prazo]<=TODAY(),[@Falta],[@Falta]/(DATEDIF(TODAY(),[@Prazo],""m"")+1)))"
    lo.ListColumns("Situação").DataBodyRange.Formula = "=IF([@Meta]="""","""",IF([@Acumulado]>=[@Alvo],""Concluída"",IF([@Prazo]<TODAY(),""Atrasada"",""Em andamento"")))"
    lo.ListColumns("Alvo").DataBodyRange.NumberFormat = """R$ ""#,##0.00"
    lo.ListColumns("Acumulado").DataBodyRange.NumberFormat = """R$ ""#,##0.00"
    lo.ListColumns("Falta").DataBodyRange.NumberFormat = """R$ ""#,##0.00"
    lo.ListColumns("Aporte mensal").DataBodyRange.NumberFormat = """R$ ""#,##0.00"
    lo.ListColumns("% concluído").DataBodyRange.NumberFormat = "0%"
    lo.ListColumns("Prazo").DataBodyRange.NumberFormat = "dd/mm/yyyy"
    With lo.ListColumns("Código").DataBodyRange.Validation
        .Delete: .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=INDIRECT(""tbCategorias[Código]"")"
    End With
    lo.TableStyle = "CO Tabela"
    ThisWorkbook.Names.Add Name:="ListaMetas", RefersTo:="=tbMetas[Meta]"
    ThisWorkbook.Names.Add Name:="nmMeta", RefersTo:="=tbMetas[Meta]"
    ThisWorkbook.Names.Add Name:="nmAlvo", RefersTo:="=tbMetas[Alvo]"
    ThisWorkbook.Names.Add Name:="nmAcum", RefersTo:="=tbMetas[Acumulado]"
End Sub

'===============================================================================
' TABELAS DINÂMICAS (um único cache sobre tbLancamentos)
'===============================================================================
Private Sub CriarDinamicas()
    Dim ws As Worksheet, pc As PivotCache, pt As PivotTable
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    ws.Name = SH_PIV
    Set pc = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:="tbLancamentos")

    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("A4"), TableName:="ptMes")
    pt.PivotFields("Mês Ref").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Valor"), "Gasto no mês", xlSum

    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("D4"), TableName:="ptGrupo")
    pt.PivotFields("Grupo").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Valor"), "Gasto por grupo", xlSum

    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("G4"), TableName:="ptSub")
    pt.PivotFields("Subcategoria").Orientation = xlRowField
    pt.AddDataField pt.PivotFields("Valor"), "Gasto", xlSum
    pt.PivotFields("Subcategoria").AutoSort xlDescending, "Gasto"
    On Error Resume Next
    pt.PivotFields("Subcategoria").PivotFilters.Add2 Type:=xlTopCount, DataField:=pt.DataFields("Gasto"), Value1:=10
    On Error GoTo 0

    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("J4"), TableName:="ptAux")
    pt.PivotFields("Pagamento").Orientation = xlRowField
    pt.PivotFields("Responsável").Orientation = xlColumnField
    pt.AddDataField pt.PivotFields("Valor"), "Total", xlSum
    RenomearVazios pt

    For Each pt In ws.PivotTables
        pt.DataBodyRange.NumberFormat = """R$ ""#,##0"
        pt.ManualUpdate = False
    Next pt
    ws.Visible = xlSheetHidden
End Sub

Private Sub RenomearVazios(pt As PivotTable)
    Dim pf As PivotField, pi As PivotItem
    On Error Resume Next
    For Each pf In pt.PivotFields
        If pf.Name = "Pagamento" Or pf.Name = "Responsável" Then
            For Each pi In pf.PivotItems
                If pi.Name = "(blank)" Or pi.Name = "(vazio)" Then pi.Caption = "Não informado"
            Next pi
        End If
    Next pf
End Sub

'===============================================================================
' CÁLCULOS (aba Calc)
'===============================================================================
Private Sub PreencherCalc()
    Dim wc As Worksheet, m As Long, r As Long, meses As Variant, up As String, dn As String, ok As String, av As String
    Dim a() As Variant
    Set wc = ThisWorkbook.Worksheets(SH_CALC)
    meses = Array("JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ")
    up = ChrW(9650) & " ": dn = ChrW(9660) & " ": ok = ChrW(10003) & " ": av = ChrW(9888) & " "

    wc.Range("A4:K4").Value = Array("Mês", "Rótulo", "Ref", "Filtro", "Renda", "Renda prev.", "Despesa", "Previsto", "gRenda", "gDespesa", "gSaldoAcum")
    ReDim a(1 To 12, 1 To 4)
    For m = 1 To 12
        a(m, 1) = m: a(m, 2) = meses(m - 1): a(m, 3) = "'" & Format(m, "00") & "-" & meses(m - 1): a(m, 4) = 1
    Next m
    wc.Range("A5:D16").Value = a
    wc.Range("E5:E16").Formula = "=SUMIFS(tbRendas[Real],tbRendas[Mês],A5)"
    wc.Range("F5:F16").Formula = "=SUMIFS(tbRendas[Previsto],tbRendas[Mês],A5)"
    wc.Range("G5:G16").Formula = "=IFERROR(GETPIVOTDATA(""Gasto no mês"",'" & SH_PIV & "'!$A$4,""Mês Ref"",C5),0)"
    wc.Range("H5:H16").Formula = "=SUMPRODUCT(SUMIFS(INDEX('" & SH_ORC & "'!$D$5:$O$300,0,A5),'" & SH_ORC & "'!$B$5:$B$300,$B$21:$B$40)*$C$21:$C$40*($B$21:$B$40<>""""))"
    wc.Range("I5:I16").Formula = "=IF(D5=1,E5,NA())"
    wc.Range("J5:J16").Formula = "=IF(D5=1,G5,NA())"
    wc.Range("K5:K16").Formula = "=IF(D5=1,SUMPRODUCT($D$5:D5,$E$5:E5)-SUMPRODUCT($D$5:D5,$G$5:G5),NA())"
    For m = 1 To 12: wc.Cells(2, 3 + m).Formula = "=INDEX($D$5:$D$16," & m & ")": Next m

    wc.Range("A20:G20").Value = Array("", "Grupo", "Filtro", "Previsto", "Realizado", "gPrevisto", "gRealizado")
    wc.Range("A19").Value = "Qtde grupos": wc.Range("B19").Formula = "=COUNTIF(B21:B40,""?*"")"
    For r = 21 To 40
        wc.Cells(r, 2).Formula = "=IFERROR(IF(INDEX(ListaGrupos," & r - 20 & ")=0,"""",INDEX(ListaGrupos," & r - 20 & ")),"""")"
    Next r
    wc.Range("C21:C40").Value = 1
    wc.Range("D21:D40").Formula = "=IF(B21="""","""",SUMPRODUCT(('" & SH_ORC & "'!$B$5:$B$300=B21)*'" & SH_ORC & "'!$D$5:$O$300*$D$2:$O$2))"
    wc.Range("E21:E40").Formula = "=IF(B21="""","""",IFERROR(GETPIVOTDATA(""Gasto por grupo"",'" & SH_PIV & "'!$D$4,""Grupo"",B21),0))"
    wc.Range("F21:F40").Formula = "=IF(OR(B21="""",C21=0),NA(),D21)"
    wc.Range("G21:G40").Formula = "=IF(OR(B21="""",C21=0),NA(),E21)"
    wc.Range("D21:G40").NumberFormat = """R$ ""#,##0"
    ' faixas dos gráficos sem funções voláteis
    ThisWorkbook.Names.Add Name:="gCat", RefersTo:="='" & SH_CALC & "'!$B$21:INDEX('" & SH_CALC & "'!$B$21:$B$40,MAX(1,'" & SH_CALC & "'!$B$19))"
    ThisWorkbook.Names.Add Name:="gPrev", RefersTo:="='" & SH_CALC & "'!$F$21:INDEX('" & SH_CALC & "'!$F$21:$F$40,MAX(1,'" & SH_CALC & "'!$B$19))"
    ThisWorkbook.Names.Add Name:="gReal", RefersTo:="='" & SH_CALC & "'!$G$21:INDEX('" & SH_CALC & "'!$G$21:$G$40,MAX(1,'" & SH_CALC & "'!$B$19))"

    ' indicadores: B número | C texto | D situação (1 bom, -1 ruim, 2 atenção, 0 neutro) | E comparação
    wc.Range("A44:E44").Value = Array("Indicador", "Valor", "Texto", "Situação", "Comparação")
    wc.Range("A45:A54").Value = Application.WorksheetFunction.Transpose(Array("Renda", "Renda prevista", "Despesas", "Despesa prevista", "Saldo", "Saldo previsto", "% renda gasta", "% orçamento", "Meses com gasto", "Média mensal"))
    wc.Range("B45").Formula = "=SUMPRODUCT(D5:D16,E5:E16)"
    wc.Range("B46").Formula = "=SUMPRODUCT(D5:D16,F5:F16)"
    wc.Range("B47").Formula = "=IFERROR(GETPIVOTDATA(""Gasto no mês"",'" & SH_PIV & "'!$A$4),0)"
    wc.Range("B48").Formula = "=SUMPRODUCT(D5:D16,H5:H16)"
    wc.Range("B49").Formula = "=B45-B47"
    wc.Range("B50").Formula = "=B46-B48"
    wc.Range("B51").Formula = "=IFERROR(B47/B45,0)"
    wc.Range("B52").Formula = "=IFERROR(B47/B48,0)"
    wc.Range("B53").Formula = "=SUMPRODUCT(D5:D16*(G5:G16>0))"
    wc.Range("B54").Formula = "=IFERROR(B47/B53,0)"
    wc.Range("C45").Formula = "=DOLLAR(B45,0)"
    wc.Range("C47").Formula = "=DOLLAR(B47,0)"
    wc.Range("C49").Formula = "=DOLLAR(B49,0)"
    wc.Range("C51").Formula = "=TEXT(B51,""0%"")"
    wc.Range("C52").Formula = "=TEXT(B52,""0%"")"
    wc.Range("D45").Formula = "=IF(B46=0,0,IF(B45>=B46,1,-1))"
    wc.Range("D47").Formula = "=IF(B48=0,0,IF(B47<=B48,1,-1))"
    wc.Range("D49").Formula = "=IF(B49>=0,1,-1)"
    wc.Range("D51").Formula = "=IF(B45=0,0,IF(B51<=0.8,1,IF(B51<=1,2,-1)))"
    wc.Range("D52").Formula = "=IF(B48=0,0,IF(B52<=0.9,1,IF(B52<=1,2,-1)))"
    wc.Range("E45").Formula = "=IF(B46=0,""Sem renda prevista"",IF(B45>=B46,""" & up & """,""" & dn & """)&DOLLAR(ABS(B45-B46),0)&"" vs previsto"")"
    wc.Range("E47").Formula = "=IF(B48=0,""Sem orçamento no período"",IF(B47<=B48,""" & dn & """&DOLLAR(B48-B47,0)&"" abaixo do previsto"",""" & up & """&DOLLAR(B47-B48,0)&"" acima do previsto""))"
    wc.Range("E49").Formula = "=""Previsto ""&DOLLAR(B50,0)"
    wc.Range("E51").Formula = "=IF(B45=0,""Sem renda no período"",IF(B51<=0.8,""Saudável"",IF(B51<=1,""Perto do limite"",""Gastando mais do que ganha"")))"
    wc.Range("E52").Formula = "=""de ""&DOLLAR(B48,0)&"" · média ""&DOLLAR(B54,0)&""/mês"""

    ' consistência e período
    wc.Range("A56").Value = "Problemas"
    wc.Range("B56").Formula = "=COUNTIF(tbLancamentos[Grupo],""CÓDIGO INVÁLIDO"")+COUNTIFS(tbLancamentos[Data],""<""&DATE(Ano,1,1))+COUNTIFS(tbLancamentos[Data],"">""&DATE(Ano,12,31))+COUNTIFS(tbLancamentos[Valor],""<=0"")+COUNTIF(tbContas[Categoria],""CÓDIGO INVÁLIDO"")"
    wc.Range("C56").Formula = "=IF(B56=0,""" & ok & "Dados consistentes"",""" & av & """&B56&"" registro(s) a revisar"")"
    wc.Range("D56").Formula = "=IF(B56=0,1,-1)"
    wc.Range("A58").Value = "Período": wc.Range("C58").Value = "Exercício " & wc.Range("B1").Value

    ' contas a pagar
    wc.Range("A60:A66").Value = Application.WorksheetFunction.Transpose(Array("Em aberto", "Vencidas (valor)", "Vencidas (qtd)", "7 dias (valor)", "7 dias (qtd)", "Pagas", "Qtd aberta"))
    wc.Range("B60").Formula = "=SUMIFS(tbContas[Valor],tbContas[Status],""Aberta"")"
    wc.Range("B61").Formula = "=SUMIFS(tbContas[Valor],tbContas[Status],""Aberta"",tbContas[Vencimento],""<""&TODAY())"
    wc.Range("B62").Formula = "=COUNTIFS(tbContas[Status],""Aberta"",tbContas[Vencimento],""<""&TODAY())"
    wc.Range("B63").Formula = "=SUMIFS(tbContas[Valor],tbContas[Status],""Aberta"",tbContas[Vencimento],"">=""&TODAY(),tbContas[Vencimento],""<=""&TODAY()+7)"
    wc.Range("B64").Formula = "=COUNTIFS(tbContas[Status],""Aberta"",tbContas[Vencimento],"">=""&TODAY(),tbContas[Vencimento],""<=""&TODAY()+7)"
    wc.Range("B65").Formula = "=SUMIFS(tbContas[Valor],tbContas[Status],""Paga"")"
    wc.Range("B66").Formula = "=COUNTIFS(tbContas[Status],""Aberta"")"
    wc.Range("C60").Formula = "=DOLLAR(B60,0)"
    wc.Range("C61").Formula = "=DOLLAR(B61,0)"
    wc.Range("C63").Formula = "=DOLLAR(B63,0)"
    wc.Range("C65").Formula = "=DOLLAR(B65,0)"
    wc.Range("D60").Formula = "=IF(B62>0,-1,IF(B64>0,2,IF(B66>0,0,1)))"
    wc.Range("D61").Formula = "=IF(B62>0,-1,1)"
    wc.Range("D63").Formula = "=IF(B64>0,2,1)"
    wc.Range("D65").Formula = "=0"
    wc.Range("E60").Formula = "=IF(B62>0,B62&"" vencida(s) · ""&DOLLAR(B61,0),IF(B64>0,B64&"" vence(m) em 7 dias"",IF(B66=0,""Nenhuma conta em aberto"",B66&"" conta(s) em dia"")))"
    wc.Range("E61").Formula = "=B62&"" conta(s)"""
    wc.Range("E63").Formula = "=B64&"" conta(s)"""
    wc.Range("E65").Formula = "=COUNTIFS(tbContas[Status],""Paga"")&"" paga(s) no exercício"""

    ' metas
    wc.Range("A70:A73").Value = Application.WorksheetFunction.Transpose(Array("Alvo", "Acumulado", "% geral", "Aporte mensal"))
    wc.Range("B70").Formula = "=SUM(tbMetas[Alvo])"
    wc.Range("B71").Formula = "=SUM(tbMetas[Acumulado])"
    wc.Range("B72").Formula = "=IFERROR(MIN(1,B71/B70),0)"
    wc.Range("B73").Formula = "=SUM(tbMetas[Aporte mensal])"
    wc.Range("C70").Formula = "=DOLLAR(B70,0)"
    wc.Range("C71").Formula = "=DOLLAR(B71,0)"
    wc.Range("C72").Formula = "=TEXT(B72,""0%"")"
    wc.Range("C73").Formula = "=DOLLAR(B73,0)"
    wc.Range("D70").Formula = "=0"
    wc.Range("D71").Formula = "=0"
    wc.Range("D72").Formula = "=IF(B70=0,0,IF(B72>=1,1,IF(COUNTIF(tbMetas[Situação],""Atrasada"")>0,-1,2)))"
    wc.Range("D73").Formula = "=0"
    wc.Range("E70").Formula = "=COUNTIF(tbMetas[Meta],""?*"")&"" meta(s)"""
    wc.Range("E71").Formula = "=COUNTIF(tbMetas[Situação],""Concluída"")&"" concluída(s)"""
    wc.Range("E72").Formula = "=COUNTIF(tbMetas[Situação],""Atrasada"")&"" atrasada(s)"""
    wc.Range("E73").Formula = "=""para cumprir no prazo"""
End Sub

'===============================================================================
' LAYOUT COMUM: cabeçalho + barra de navegação (todas as telas)
'===============================================================================
Private Function Telas() As Variant
    Telas = Array(Array("Painel", SH_DASH), Array("Análise", SH_ANA), Array("Lançar", SH_CAD), Array("Contas", SH_CON), _
                  Array("Metas", SH_MET), Array("Lançamentos", SH_LANC), Array("Orçamento", SH_ORC), Array("Rendas", SH_REND), Array("Categorias", SH_CAT))
End Function

Private Sub Cabecalho(ws As Worksheet, ativo As String)
    Dim shp As Shape, t As Variant, i As Long, x As Single, w As Single
    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, LARG + 40, 44)
    shp.Name = "cabFundo": shp.Fill.ForeColor.RGB = cHeader: shp.Line.Visible = msoFalse: shp.Placement = xlFreeFloating
    Texto ws, 18, 6, 300, 20, "Casa Organizada", 14, RGB(255, 255, 255), True
    Set shp = Texto(ws, 18, 26, 300, 12, "", 8.5, RGB(165, 180, 200), False, "='" & SH_CALC & "'!$C$58")
    shp.Name = "cabPeriodo"
    Set shp = Texto(ws, 250, 16, 260, 14, "", 9, RGB(134, 239, 172), True, "='" & SH_CALC & "'!$C$56")
    shp.Name = "chipStatus"
    Set shp = BotaoForma(ws, 740, 9, 132, 26, "+ Lançar gasto", "IrCadastro", cGold, cHeader, 9.5)

    Set shp = ws.Shapes.AddShape(msoShapeRectangle, 0, 44, LARG + 40, 30)
    shp.Name = "navFundo": shp.Fill.ForeColor.RGB = cCard: shp.Line.Visible = msoFalse: shp.Placement = xlFreeFloating
    Set shp = ws.Shapes.AddLine(0, 74, LARG + 40, 74)
    shp.Name = "navLinha": shp.Line.ForeColor.RGB = cLinha: shp.Placement = xlFreeFloating
    t = Telas(): x = 12
    For i = 0 To UBound(t)
        w = 22 + Len(t(i)(0)) * 6.4
        Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, 48, w, 24)
        With shp
            .Name = "nav_" & i
            .OnAction = "'" & ThisWorkbook.Name & "'!Navegar"
            .Line.Visible = msoFalse: .Fill.Visible = msoFalse
            With .TextFrame2
                .MarginLeft = 0: .MarginRight = 0: .VerticalAnchor = msoAnchorMiddle: .WordWrap = msoFalse
                .TextRange.Text = t(i)(0)
                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
                .TextRange.Font.Name = FONTE: .TextRange.Font.Size = 9.5
                .TextRange.Font.Bold = IIf(t(i)(1) = ativo, msoTrue, msoFalse)
                .TextRange.Font.Fill.ForeColor.RGB = IIf(t(i)(1) = ativo, cHeader, cMuted)
            End With
            .Placement = xlFreeFloating
        End With
        If t(i)(1) = ativo Then
            Set shp = ws.Shapes.AddShape(msoShapeRectangle, x + 6, 71, w - 12, 3)
            shp.Name = "navAtivo": shp.Fill.ForeColor.RGB = cGold: shp.Line.Visible = msoFalse: shp.Placement = xlFreeFloating
        End If
        x = x + w + 4
        If i = 4 Then
            Set shp = ws.Shapes.AddLine(x + 2, 52, x + 2, 68)
            shp.Name = "navSep": shp.Line.ForeColor.RGB = cLinha: shp.Placement = xlFreeFloating
            x = x + 8
        End If
    Next i
    BotaoForma ws, x + 12, 49, 80, 20, "Atualizar", "AtualizarPainel", cCard, cPrim, 8.5, True
End Sub

Public Sub Navegar()
    Dim t As Variant, i As Long
    On Error Resume Next
    i = CLng(Mid(Application.Caller, 5))
    t = Telas()
    Select Case t(i)(1)
        Case SH_DASH: IrDashboard
        Case SH_ANA: Ir SH_ANA, "A1": AjustarZoom LARG + 10
        Case SH_CAD: IrCadastro
        Case Else: Ir CStr(t(i)(1)), "A1"
    End Select
End Sub

Private Sub PrepararTela(ws As Worksheet, Optional linhas As Boolean = True)
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.DisplayHeadings = False
    ws.Cells.Font.Name = FONTE
    ws.Cells.Font.Size = 10
    If linhas Then ws.Range("A1:AZ150").Interior.Color = cFundo
    ws.Rows(1).RowHeight = 44
    ws.Rows(2).RowHeight = 30
End Sub

'===============================================================================
' PAINEL (executivo)
'===============================================================================
Private Sub CriarPainel()
    Dim ws As Worksheet, wp As Worksheet, wc As Worksheet, ch As ChartObject, s As Series
    Dim x0 As Single, w As Single, gap As Single, i As Long, cards As Variant, cw As Single

    Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
    ws.Name = SH_DASH
    Set wp = ThisWorkbook.Worksheets(SH_PIV)
    Set wc = ThisWorkbook.Worksheets(SH_CALC)
    PrepararTela ws
    Cabecalho ws, SH_DASH

    ' painel de filtros
    Cartao ws, 8, Y0, 188, 628, "pnFiltros"
    Texto ws, 20, Y0 + 10, 100, 18, "Filtros", 11, cTexto, True
    BotaoForma ws, 134, Y0 + 8, 52, 20, "Limpar", "LimparFiltros", cCard, cPrim, 8.5, True

    ' 6 indicadores
    x0 = 205: gap = 10
    cw = (LARG - x0 - 5 * gap) / 6
    cards = Array( _
        Array("RENDA", "$C$45", "$E$45", "45"), Array("DESPESAS", "$C$47", "$E$47", "47"), Array("SALDO", "$C$49", "$E$49", "49"), _
        Array("% DA RENDA GASTA", "$C$51", "$E$51", "51"), Array("ORÇAMENTO USADO", "$C$52", "$E$52", "52"), Array("CONTAS EM ABERTO", "$C$60", "$E$60", "60"))
    For i = 0 To 5
        Kpi ws, x0 + i * (cw + gap), Y0, cw, 84, CStr(cards(i)(0)), CStr(cards(i)(1)), CStr(cards(i)(2)), CStr(cards(i)(3))
    Next i

    ' Renda x despesas por mês + saldo acumulado
    Set ch = ws.ChartObjects.Add(x0, Y0 + 94, 600, 262)
    ch.Name = "gMeses"
    With ch.Chart
        .ChartType = xlColumnClustered
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        Set s = .SeriesCollection.NewSeries: s.Name = "Renda": s.Values = wc.Range("I5:I16"): s.XValues = wc.Range("B5:B16")
        s.Format.Fill.ForeColor.RGB = cPrim
        Set s = .SeriesCollection.NewSeries: s.Name = "Despesas": s.Values = wc.Range("J5:J16"): s.XValues = wc.Range("B5:B16")
        s.Format.Fill.ForeColor.RGB = cSec
        Set s = .SeriesCollection.NewSeries: s.Name = "Saldo acumulado": s.Values = wc.Range("K5:K16"): s.XValues = wc.Range("B5:B16")
        s.ChartType = xlLineMarkers
        s.Format.Line.ForeColor.RGB = cTeal: s.Format.Line.Weight = 2.25
        s.MarkerStyle = xlMarkerStyleCircle: s.MarkerSize = 5
        s.MarkerBackgroundColor = cTeal: s.MarkerForegroundColor = cTeal
        .ChartGroups(1).GapWidth = 70: .ChartGroups(1).Overlap = -5
    End With
    EstiloGrafico ch, "Renda x despesas por mês", True

    ' Top 10 subcategorias
    Set ch = ws.ChartObjects.Add(x0 + 610, Y0 + 94, LARG - x0 - 610, 262)
    ch.Name = "gTop"
    ch.Chart.SetSourceData wp.PivotTables("ptSub").TableRange1
    ch.Chart.ChartType = xlBarClustered
    With ch.Chart
        .SeriesCollection(1).Format.Fill.ForeColor.RGB = cPrim
        .ChartGroups(1).GapWidth = 45
        .Axes(xlCategory).ReversePlotOrder = True
    End With
    EstiloGrafico ch, "Top 10 subcategorias", False
    With ch.Chart
        .HasAxis(xlValue) = False
        .SeriesCollection(1).ApplyDataLabels
        .SeriesCollection(1).DataLabels.NumberFormatLocal = FormatoMoedaLocal()
        .SeriesCollection(1).DataLabels.Position = xlLabelPositionOutsideEnd
        .SeriesCollection(1).DataLabels.Font.Size = 8.5
        .SeriesCollection(1).DataLabels.Font.Color = cTexto
    End With

    ' Previsto x realizado por grupo
    Set ch = ws.ChartObjects.Add(x0, Y0 + 366, 600, 262)
    ch.Name = "gGrupos"
    With ch.Chart
        .ChartType = xlBarClustered
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        Set s = .SeriesCollection.NewSeries: s.Name = "Previsto"
        s.Values = "='" & ThisWorkbook.Name & "'!gPrev": s.XValues = "='" & ThisWorkbook.Name & "'!gCat"
        s.Format.Fill.ForeColor.RGB = cPrev
        Set s = .SeriesCollection.NewSeries: s.Name = "Realizado"
        s.Values = "='" & ThisWorkbook.Name & "'!gReal": s.XValues = "='" & ThisWorkbook.Name & "'!gCat"
        s.Format.Fill.ForeColor.RGB = cPrim
        .ChartGroups(1).GapWidth = 40: .ChartGroups(1).Overlap = 0
        .Axes(xlCategory).ReversePlotOrder = True
    End With
    EstiloGrafico ch, "Previsto x realizado por grupo", True

    ' Metas e reservas (barra de progresso: acumulado sobre o alvo)
    Set ch = ws.ChartObjects.Add(x0 + 610, Y0 + 366, LARG - x0 - 610, 262)
    ch.Name = "gMetas"
    With ch.Chart
        .ChartType = xlBarClustered
        Do While .SeriesCollection.Count > 0: .SeriesCollection(1).Delete: Loop
        Set s = .SeriesCollection.NewSeries: s.Name = "Alvo"
        s.Values = "='" & ThisWorkbook.Name & "'!nmAlvo": s.XValues = "='" & ThisWorkbook.Name & "'!nmMeta"
        s.Format.Fill.ForeColor.RGB = RGB(226, 232, 240)
        Set s = .SeriesCollection.NewSeries: s.Name = "Acumulado"
        s.Values = "='" & ThisWorkbook.Name & "'!nmAcum": s.XValues = "='" & ThisWorkbook.Name & "'!nmMeta"
        s.Format.Fill.ForeColor.RGB = cTeal
        .ChartGroups(1).GapWidth = 55: .ChartGroups(1).Overlap = 100
        .Axes(xlCategory).ReversePlotOrder = True
    End With
    EstiloGrafico ch, "Metas e reservas", True

    CriarSegmentacoes ws, wp
End Sub

Private Sub CriarSegmentacoes(ws As Worksheet, wp As Worksheet)
    Dim campos As Variant, rotulos As Variant, alturas As Variant, i As Long, y As Single, sc As SlicerCache, pt As PivotTable, sl As Slicer
    campos = Array("Mês Ref", "Grupo", "Responsável", "Pagamento")
    rotulos = Array("Mês", "Grupo", "Responsável", "Pagamento")
    alturas = Array(166, 206, 90, 110)
    y = Y0 + 34
    On Error Resume Next
    For i = 0 To 3
        Set sc = Nothing
        Set sc = ThisWorkbook.SlicerCaches.Add2(wp.PivotTables("ptMes"), campos(i))
        If Not sc Is Nothing Then
            For Each pt In wp.PivotTables
                If pt.Name <> "ptMes" Then sc.PivotTables.AddPivotTable pt
            Next pt
            Set sl = sc.Slicers.Add(ws, , "slPainel" & i, rotulos(i), y, 16, 172, alturas(i))
            EstiloSegmentacao sl, campos(i) = "Mês Ref"
            y = y + alturas(i) + 8
        End If
    Next i
    On Error GoTo 0
End Sub

Private Sub EstiloSegmentacao(sl As Slicer, mes As Boolean)
    On Error Resume Next
    sl.Style = "CO Filtro"
    If Err.Number <> 0 Then Err.Clear: sl.Style = "SlicerStyleLight1"
    If mes Then sl.NumberOfColumns = 2
    sl.RowHeight = 18
    sl.DisableMoveResizeUI = True
End Sub

'===============================================================================
' ANÁLISE (tabelas dinâmicas visíveis, mesmo cache, mesmos filtros)
'===============================================================================
Private Sub CriarAnalise()
    Dim ws As Worksheet, wp As Worksheet, pc As PivotCache, pt As PivotTable, sc As SlicerCache, sl As Slicer, fc As Object, i As Long, y As Single

    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(SH_DASH))
    ws.Name = SH_ANA
    Set wp = ThisWorkbook.Worksheets(SH_PIV)
    PrepararTela ws
    ws.Columns("A").ColumnWidth = 30
    ws.Rows(3).RowHeight = 12
    Cabecalho ws, SH_ANA
    Set pc = wp.PivotTables("ptMes").PivotCache

    ' matriz Grupo > Subcategoria x Mês com mapa de calor
    ws.Range("B4").Value = "Despesas por grupo e mês"
    Titulo ws.Range("B4")
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("B6"), TableName:="ptMatriz")
    With pt
        .RowAxisLayout xlCompactRow
        .PivotFields("Grupo").Orientation = xlRowField
        .PivotFields("Subcategoria").Orientation = xlRowField
        .PivotFields("Mês Ref").Orientation = xlColumnField
        .AddDataField .PivotFields("Valor"), "Total ", xlSum
        .CompactLayoutRowHeader = "Grupo / Subcategoria"
        .CompactLayoutColumnHeader = "Mês"
        .PivotFields("Grupo").ShowDetail = False
        .ShowDrillIndicators = True
        .HasAutoFormat = False
        .DataBodyRange.NumberFormat = """R$ ""#,##0;-""R$ ""#,##0;""-"""
    End With
    EstiloPivot pt
    Set fc = pt.DataBodyRange.FormatConditions.AddColorScale(2)
    fc.ScopeType = xlDataFieldScope
    fc.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 255, 255)
    fc.ColorScaleCriteria(2).FormatColor.Color = RGB(147, 180, 214)

    ' participação por grupo
    ws.Range("Q4").Value = "Participação e ticket médio por grupo"
    Titulo ws.Range("Q4")
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("Q6"), TableName:="ptParticipacao")
    With pt
        .RowAxisLayout xlCompactRow
        .PivotFields("Grupo").Orientation = xlRowField
        .AddDataField .PivotFields("Valor"), "Total  ", xlSum
        .AddDataField .PivotFields("Valor"), "% do total", xlSum
        .AddDataField .PivotFields("Valor"), "Lançamentos", xlCount
        .AddDataField .PivotFields("Valor"), "Ticket médio", xlAverage
        .DataFields("% do total").Calculation = xlPercentOfTotal
        .DataFields("% do total").NumberFormat = "0.0%"
        .DataFields("Total  ").NumberFormat = """R$ ""#,##0"
        .DataFields("Ticket médio").NumberFormat = """R$ ""#,##0"
        .DataFields("Lançamentos").NumberFormat = "0"
        .PivotFields("Grupo").AutoSort xlDescending, "Total  "
        .CompactLayoutRowHeader = "Grupo"
    End With
    EstiloPivot pt
    Set fc = pt.DataFields("% do total").DataRange.FormatConditions.AddDatabar
    fc.ScopeType = xlDataFieldScope
    fc.BarColor.Color = RGB(147, 180, 214)

    ' responsável x pagamento
    ws.Range("Q22").Value = "Por responsável e forma de pagamento"
    Titulo ws.Range("Q22")
    Set pt = pc.CreatePivotTable(TableDestination:=ws.Range("Q24"), TableName:="ptRespPag")
    With pt
        .RowAxisLayout xlCompactRow
        .PivotFields("Responsável").Orientation = xlRowField
        .PivotFields("Pagamento").Orientation = xlColumnField
        .AddDataField .PivotFields("Valor"), "Total   ", xlSum
        .DataBodyRange.NumberFormat = """R$ ""#,##0"
        .CompactLayoutRowHeader = "Responsável"
        .CompactLayoutColumnHeader = "Pagamento"
    End With
    RenomearVazios pt
    EstiloPivot pt

    ws.Columns("C:O").ColumnWidth = 10.5
    ws.Columns("P").ColumnWidth = 3
    ws.Columns("Q").ColumnWidth = 24
    ws.Columns("R:X").ColumnWidth = 12.5

    ' os mesmos filtros do Painel (mesmas segmentações, sincronizadas)
    y = Y0 + 10
    On Error Resume Next
    For Each sc In ThisWorkbook.SlicerCaches
        If sc.Slicers(1).Caption = "Mês" Or sc.Slicers(1).Caption = "Grupo" Then
            For Each pt In ws.PivotTables
                sc.PivotTables.AddPivotTable pt
            Next pt
        Else
            For Each pt In ws.PivotTables
                sc.PivotTables.AddPivotTable pt
            Next pt
        End If
    Next sc
    On Error GoTo 0
    ws.Range("A1").Select
End Sub

Private Sub EstiloPivot(pt As PivotTable)
    On Error Resume Next
    pt.TableStyle2 = "CO Tabela"
    pt.ShowTableStyleRowStripes = False
    pt.ShowTableStyleColumnHeaders = True
    pt.ShowTableStyleRowHeaders = True
    pt.TableRange1.Font.Name = FONTE
    pt.TableRange1.Font.Size = 9.5
    pt.DisplayErrorString = True: pt.ErrorString = "-"
    pt.DisplayNullString = True: pt.NullString = ""
End Sub

Private Sub Titulo(c As Range)
    c.Font.Name = FONTE: c.Font.Size = 12: c.Font.Bold = True: c.Font.Color = cTexto
End Sub

'===============================================================================
' LANÇAR (cadastro)
'===============================================================================
Private Sub CriarCadastro()
    Dim ws As Worksheet, w As Variant, i As Long
    ApagarAba SH_CAD
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(SH_ANA))
    ws.Name = SH_CAD
    PrepararTela ws
    ' colunas: [rótulo | campo | respiro] por card, com vão entre os cards
    w = Array(2, 14, 38, 2, 3, 14, 26, 2, 3, 12, 22, 2)
    For i = 0 To UBound(w): ws.Columns(i + 1).ColumnWidth = w(i): Next i
    ws.Rows(3).RowHeight = 12
    ws.Rows("4:15").RowHeight = 24
    ws.Rows(5).RowHeight = 8: ws.Rows(13).RowHeight = 8: ws.Rows(15).RowHeight = 10
    ws.Rows(16).RowHeight = 14
    ws.Rows("17:21").RowHeight = 24
    ws.Rows(18).RowHeight = 8: ws.Rows(21).RowHeight = 10
    ws.Rows(22).RowHeight = 14
    Cabecalho ws, SH_CAD

    ' card 1: novo gasto
    Painel ws, "B4:D15", "Novo gasto"
    Campo ws, 6, 2, "Data", "frmData", "dd/mm/yyyy"
    ws.Range("frmData").Value = DataPadrao()
    Campo ws, 7, 2, "Categoria", "frmCat"
    Lista ws.Range("frmCat"), "=ListaCategorias"
    Campo ws, 8, 2, "Valor", "frmValor", """R$ ""#,##0.00"
    With ws.Range("frmValor").Validation
        .Delete
        .Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreater, Formula1:="0"
        .ErrorMessage = "O valor deve ser maior que zero."
    End With
    Campo ws, 9, 2, "Descrição", "frmDesc"
    Campo ws, 10, 2, "Pagamento", "frmPag"
    Lista ws.Range("frmPag"), "='" & SH_CALC & "'!$H$1:$H$6"
    Campo ws, 11, 2, "Responsável", "frmResp"
    Lista ws.Range("frmResp"), "='" & SH_CALC & "'!$I$1:$I$3"
    ws.Range("C12").Formula = "=IFERROR(IF(frmCat="""","""",""Previsto no mês ""&DOLLAR(IFERROR(INDEX('" & SH_ORC & "'!$D$5:$O$300,MATCH(VALUE(LEFT(frmCat,FIND("" "",frmCat)-1)),'" & SH_ORC & "'!$A$5:$A$300,0),MONTH(frmData)),0),2)&""  ·  gasto ""&DOLLAR(SUMIFS(tbLancamentos[Valor],tbLancamentos[Código],VALUE(LEFT(frmCat,FIND("" "",frmCat)-1)),tbLancamentos[Mês],MONTH(frmData)),2)),"""")"
    ws.Range("C12").Font.Size = 8.5: ws.Range("C12").Font.Color = cMuted
    BotaoForma ws, ws.Range("B14").Left + 10, ws.Range("B14").Top - 1, 110, 26, "Lançar gasto", "LancarGasto", cPrim, RGB(255, 255, 255), 9.5
    BotaoForma ws, ws.Range("B14").Left + 126, ws.Range("B14").Top - 1, 62, 26, "Limpar", "LimparLancamento", cCard, cTexto, 9.5, True
    BotaoForma ws, ws.Range("B14").Left + 194, ws.Range("B14").Top - 1, 104, 26, "Desfazer último", "DesfazerUltimoLancamento", cCard, cTexto, 9.5, True

    ' card 2: nova categoria
    Painel ws, "F4:H15", "Nova categoria"
    Campo ws, 6, 6, "Grupo", "catGrupo"
    Lista ws.Range("catGrupo"), "=ListaGrupos", False
    Campo ws, 7, 6, "Subcategoria", "catSub"
    Campo ws, 8, 6, "Previsto/mês", "catPrev", """R$ ""#,##0.00"
    Campo ws, 9, 6, "Conta fixa?", "catFixa"
    Lista ws.Range("catFixa"), "='" & SH_CALC & "'!$L$1:$L$2"
    ws.Range("catFixa").Value = "Não"
    BotaoForma ws, ws.Range("F14").Left + 10, ws.Range("F14").Top - 1, 150, 26, "Cadastrar categoria", "CadastrarCategoria", cPrim, RGB(255, 255, 255), 9.5

    ' card 3: nova renda
    Painel ws, "J4:L15", "Nova renda"
    Campo ws, 6, 10, "Mês", "rMes", "0"
    ws.Range("rMes").Value = MesPadrao()
    With ws.Range("rMes").Validation
        .Delete
        .Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="12"
        .ErrorMessage = "Informe o mês de 1 a 12."
    End With
    Campo ws, 7, 10, "Fonte", "rFonte"
    Campo ws, 8, 10, "Previsto", "rPrev", """R$ ""#,##0.00"
    Campo ws, 9, 10, "Real", "rReal", """R$ ""#,##0.00"
    BotaoForma ws, ws.Range("J14").Left + 10, ws.Range("J14").Top - 1, 130, 26, "Registrar renda", "RegistrarRenda", cPrim, RGB(255, 255, 255), 9.5

    ' card 4: automação (Python)
    Painel ws, "B17:H21", "Importar e auditar"
    BotaoForma ws, ws.Range("B19").Left + 10, ws.Range("B19").Top + 1, 190, 26, "Importar extrato (Python)", "ImportarExtrato", cTeal, RGB(255, 255, 255), 9.5
    BotaoForma ws, ws.Range("B19").Left + 208, ws.Range("B19").Top + 1, 170, 26, "Auditar dados (Python)", "AuditarDados", cCard, cTexto, 9.5, True
    BotaoForma ws, ws.Range("B19").Left + 386, ws.Range("B19").Top + 1, 140, 26, "Gerar contas do mês", "IrContas", cCard, cTexto, 9.5, True

    ' card 5: última ação
    Painel ws, "J17:L21", "Última ação"
    NomeFaixa "msgStatus", ws.Range("J19")
    With ws.Range("J19:K20")
        .Merge
        .WrapText = True
        .VerticalAlignment = xlTop
        .IndentLevel = 1
        .Font.Color = cBom: .Font.Bold = True: .Font.Size = 9
    End With
    ws.Range("frmCat").Select
End Sub

'===============================================================================
' CONTAS A PAGAR
'===============================================================================
Private Sub TelaContas()
    Dim ws As Worksheet, i As Long, cards As Variant, x As Single
    Set ws = ThisWorkbook.Worksheets(SH_CON)
    LimparFormas ws
    PrepararTela ws
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 13: ws.Columns("C").ColumnWidth = 34: ws.Columns("D").ColumnWidth = 9
    ws.Columns("E").ColumnWidth = 28: ws.Columns("F").ColumnWidth = 13: ws.Columns("G").ColumnWidth = 13
    ws.Columns("H").ColumnWidth = 11: ws.Columns("I").ColumnWidth = 12: ws.Columns("J").ColumnWidth = 17
    ws.Rows(3).RowHeight = 10
    ws.Rows("4:7").RowHeight = 17
    ws.Rows(8).RowHeight = 16: ws.Rows(9).RowHeight = 26: ws.Rows(10).RowHeight = 10
    ws.Range("A11:K500").Interior.Pattern = xlNone
    Cabecalho ws, SH_CON

    cards = Array(Array("EM ABERTO", "$C$60", "$E$60", "60"), Array("VENCIDAS", "$C$61", "$E$61", "61"), _
                  Array("VENCEM EM 7 DIAS", "$C$63", "$E$63", "63"), Array("PAGAS", "$C$65", "$E$65", "65"))
    x = 12
    For i = 0 To 3
        Kpi ws, x + i * 222, Y0, 212, 68, CStr(cards(i)(0)), CStr(cards(i)(1)), CStr(cards(i)(2)), "c" & CStr(cards(i)(3))
    Next i

    ws.Range("B8").Value = "Mês": ws.Range("B8").Font.Size = 8.5: ws.Range("B8").Font.Color = cMuted
    With ws.Range("B9")
        .Interior.Color = cCard: .Borders.Color = RGB(209, 213, 219): .HorizontalAlignment = xlCenter
        .Validation.Delete
        .Validation.Add Type:=xlValidateWholeNumber, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="1", Formula2:="12"
        If Len(.Value) = 0 Then .Value = MesPadrao()
    End With
    NomeFaixa "conMes", ws.Range("B9")
    x = ws.Range("C9").Left + 6
    BotaoForma ws, x, ws.Range("B9").Top, 150, 26, "Gerar contas do mês", "GerarContasMes", cPrim, RGB(255, 255, 255), 9
    BotaoForma ws, x + 158, ws.Range("B9").Top, 150, 26, "Pagar selecionadas", "PagarSelecionadas", cTeal, RGB(255, 255, 255), 9
    BotaoForma ws, x + 316, ws.Range("B9").Top, 100, 26, "Nova conta", "NovaConta", cCard, cTexto, 9, True

    With ws.ListObjects("tbContas").ListColumns("Situação").DataBodyRange.FormatConditions
        .Delete
        With .Add(xlCellValue, xlEqual, "=""Vencida""")
            .Font.Color = cRuim: .Font.Bold = True
        End With
        With .Add(xlCellValue, xlEqual, "=""Vence em 7 dias""")
            .Font.Color = cAlerta: .Font.Bold = True
        End With
        With .Add(xlCellValue, xlEqual, "=""Paga""")
            .Font.Color = cBom
        End With
    End With
    ws.Range("B12").Select
    ActiveWindow.FreezePanes = False
End Sub

Public Sub GerarContasMes()
    Dim wc As Worksheet, lo As ListObject, lcat As ListObject, m As Variant, ano As Long, dia As Long
    Dim cods As Variant, subs As Variant, fixa As Variant, orcCod As Variant, orcVal As Variant
    Dim exist As Object, i As Long, j As Long, n As Long, v As Double, novos() As Variant, k As String, vencs As Variant, cc As Variant
    Set lo = ThisWorkbook.Worksheets(SH_CON).ListObjects("tbContas")
    Set lcat = ThisWorkbook.Worksheets(SH_CAT).ListObjects("tbCategorias")
    m = ThisWorkbook.Names("conMes").RefersToRange.Value
    If Not IsNumeric(m) Or m < 1 Or m > 12 Then MsgBox "Informe o mês (1 a 12).", vbExclamation: Exit Sub
    ano = ThisWorkbook.Names("Ano").RefersToRange.Value
    dia = ThisWorkbook.Names("DiaVenc").RefersToRange.Value

    ' contas já existentes no mês (evita duplicar)
    Set exist = CreateObject("Scripting.Dictionary")
    If Not lo.DataBodyRange Is Nothing Then
        vencs = Col2D(lo.ListColumns("Vencimento").DataBodyRange)
        cc = Col2D(lo.ListColumns("Código").DataBodyRange)
        For i = 1 To UBound(vencs, 1)
            If IsDate(vencs(i, 1)) Then
                If Month(vencs(i, 1)) = m And Year(vencs(i, 1)) = ano Then exist(CStr(cc(i, 1))) = True
            End If
        Next i
    End If

    cods = Col2D(lcat.ListColumns("Código").DataBodyRange)
    subs = Col2D(lcat.ListColumns("Subcategoria").DataBodyRange)
    fixa = Col2D(lcat.ListColumns("Fixa").DataBodyRange)
    orcCod = ThisWorkbook.Worksheets(SH_ORC).Range("A5:A300").Value
    orcVal = ThisWorkbook.Worksheets(SH_ORC).Range("D5:O300").Value

    ReDim novos(1 To UBound(cods, 1), 1 To 7)
    For i = 1 To UBound(cods, 1)
        If fixa(i, 1) = "Sim" And Not exist.Exists(CStr(cods(i, 1))) Then
            v = 0
            For j = 1 To UBound(orcCod, 1)
                If orcCod(j, 1) = cods(i, 1) Then v = Val(orcVal(j, m)): Exit For
            Next j
            If v > 0 Then
                n = n + 1
                novos(n, 1) = DateSerial(ano, m, dia)
                novos(n, 2) = subs(i, 1) & " " & Format(m, "00") & "/" & ano
                novos(n, 3) = cods(i, 1)
                novos(n, 4) = v
                novos(n, 5) = "Mensal"
                novos(n, 6) = "Aberta"
            End If
        End If
    Next i
    If n = 0 Then Status "Nenhuma conta nova para " & Format(m, "00") & "/" & ano & ".": MsgBox "Nenhuma conta nova para gerar neste mês.", vbInformation: Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Dim ini As Long
    ini = AcrescentarLinhas(lo, n)
    For i = 1 To n
        With lo.ListRows(ini + i - 1).Range
            .Cells(1, lo.ListColumns("Vencimento").Index).Value = novos(i, 1)
            .Cells(1, lo.ListColumns("Descrição").Index).Value = novos(i, 2)
            .Cells(1, lo.ListColumns("Código").Index).Value = novos(i, 3)
            .Cells(1, lo.ListColumns("Valor").Index).Value = novos(i, 4)
            .Cells(1, lo.ListColumns("Recorrência").Index).Value = novos(i, 5)
            .Cells(1, lo.ListColumns("Status").Index).Value = novos(i, 6)
        End With
    Next i
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    AtualizarVisual
    Status n & " conta(s) gerada(s) para " & Format(m, "00") & "/" & ano & "."
End Sub

Public Sub PagarSelecionadas()
    Dim lo As ListObject, ll As ListObject, alvo As Range, c As Range, r As Long, rows As Object, k As Variant
    Dim n As Long, ini As Long, dt As Date, ano As Long, total As Double
    Set lo = ThisWorkbook.Worksheets(SH_CON).ListObjects("tbContas")
    Set ll = ThisWorkbook.Worksheets(SH_LANC).ListObjects("tbLancamentos")
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If TypeName(Selection) <> "Range" Then Exit Sub
    Set alvo = Intersect(Selection, lo.DataBodyRange)
    If alvo Is Nothing Then MsgBox "Selecione as linhas das contas a pagar.", vbInformation: Exit Sub
    ano = ThisWorkbook.Names("Ano").RefersToRange.Value
    Set rows = CreateObject("Scripting.Dictionary")
    For Each c In alvo.Rows
        r = c.Row - lo.DataBodyRange.Row + 1
        If lo.ListRows(r).Range.Cells(1, lo.ListColumns("Status").Index).Value = "Aberta" And _
           Len(lo.ListRows(r).Range.Cells(1, lo.ListColumns("Descrição").Index).Value) > 0 Then rows(r) = True
    Next c
    If rows.Count = 0 Then MsgBox "Nenhuma conta em aberto na seleção.", vbInformation: Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ini = AcrescentarLinhas(ll, rows.Count)
    n = 0
    For Each k In rows.Keys
        With lo.ListRows(k).Range
            dt = IIf(Year(Date) = ano, Date, .Cells(1, lo.ListColumns("Vencimento").Index).Value)
            With ll.ListRows(ini + n).Range
                .Cells(1, ll.ListColumns("Data").Index).Value = dt
                .Cells(1, ll.ListColumns("Código").Index).Value = lo.ListRows(k).Range.Cells(1, lo.ListColumns("Código").Index).Value
                .Cells(1, ll.ListColumns("Descrição").Index).Value = "Conta: " & lo.ListRows(k).Range.Cells(1, lo.ListColumns("Descrição").Index).Value
                .Cells(1, ll.ListColumns("Valor").Index).Value = lo.ListRows(k).Range.Cells(1, lo.ListColumns("Valor").Index).Value
                .Cells(1, ll.ListColumns("Pagamento").Index).Value = "Boleto"
                .Cells(1, ll.ListColumns("Responsável").Index).Value = "Casa"
                .Cells(1, ll.ListColumns("Origem").Index).Value = "Contas"
            End With
            total = total + .Cells(1, lo.ListColumns("Valor").Index).Value
            .Cells(1, lo.ListColumns("Status").Index).Value = "Paga"
            .Cells(1, lo.ListColumns("Pago em").Index).Value = dt
        End With
        n = n + 1
    Next k
    Application.EnableEvents = True
    AtualizarPainel
    Application.ScreenUpdating = True
    Status n & " conta(s) paga(s) · " & Format(total, """R$ ""#,##0.00") & " lançado(s) em despesas."
End Sub

Public Sub NovaConta()
    Dim lo As ListObject, r As Long
    Set lo = ThisWorkbook.Worksheets(SH_CON).ListObjects("tbContas")
    Application.EnableEvents = False
    r = AcrescentarLinhas(lo, 1)
    With lo.ListRows(r).Range
        .Cells(1, lo.ListColumns("Vencimento").Index).Value = DateSerial(ThisWorkbook.Names("Ano").RefersToRange.Value, ThisWorkbook.Names("conMes").RefersToRange.Value, ThisWorkbook.Names("DiaVenc").RefersToRange.Value)
        .Cells(1, lo.ListColumns("Recorrência").Index).Value = "Única"
        .Cells(1, lo.ListColumns("Status").Index).Value = "Aberta"
        Application.EnableEvents = True
        .Cells(1, lo.ListColumns("Descrição").Index).Select
    End With
End Sub

'===============================================================================
' METAS
'===============================================================================
Private Sub TelaMetas()
    Dim ws As Worksheet, i As Long, cards As Variant, x As Single, fc As Object
    Set ws = ThisWorkbook.Worksheets(SH_MET)
    LimparFormas ws
    PrepararTela ws
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 30: ws.Columns("C").ColumnWidth = 9: ws.Columns("D").ColumnWidth = 14
    ws.Columns("E").ColumnWidth = 12: ws.Columns("F").ColumnWidth = 14: ws.Columns("G").ColumnWidth = 13
    ws.Columns("H").ColumnWidth = 14: ws.Columns("I").ColumnWidth = 15: ws.Columns("J").ColumnWidth = 15
    ws.Rows(3).RowHeight = 10
    ws.Rows("4:7").RowHeight = 17
    ws.Rows(8).RowHeight = 16: ws.Rows(9).RowHeight = 26: ws.Rows(10).RowHeight = 10
    ws.Range("A11:K500").Interior.Pattern = xlNone
    Cabecalho ws, SH_MET

    cards = Array(Array("ALVO TOTAL", "$C$70", "$E$70", "70"), Array("ACUMULADO", "$C$71", "$E$71", "71"), _
                  Array("% CONCLUÍDO", "$C$72", "$E$72", "72"), Array("APORTE MENSAL NECESSÁRIO", "$C$73", "$E$73", "73"))
    For i = 0 To 3
        Kpi ws, 12 + i * 222, Y0, 212, 68, CStr(cards(i)(0)), CStr(cards(i)(1)), CStr(cards(i)(2)), "m" & CStr(cards(i)(3))
    Next i

    ws.Range("B8").Value = "Meta": ws.Range("D8").Value = "Valor do aporte": ws.Range("E8").Value = "Data"
    ws.Range("B8,D8,E8").Font.Size = 8.5: ws.Range("B8,D8,E8").Font.Color = cMuted
    With ws.Range("B9")
        .Interior.Color = cCard: .Borders.Color = RGB(209, 213, 219)
        .Validation.Delete
        .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=ListaMetas"
    End With
    With ws.Range("D9")
        .Interior.Color = cCard: .Borders.Color = RGB(209, 213, 219): .NumberFormat = """R$ ""#,##0.00"
    End With
    With ws.Range("E9")
        .Interior.Color = cCard: .Borders.Color = RGB(209, 213, 219): .NumberFormat = "dd/mm/yyyy"
        If Len(.Value) = 0 Then .Value = DataPadrao()
    End With
    NomeFaixa "metSel", ws.Range("B9")
    NomeFaixa "metValor", ws.Range("D9")
    NomeFaixa "metData", ws.Range("E9")
    x = ws.Range("F9").Left + 6
    BotaoForma ws, x, ws.Range("B9").Top, 140, 26, "Registrar aporte", "RegistrarAporte", cTeal, RGB(255, 255, 255), 9
    BotaoForma ws, x + 148, ws.Range("B9").Top, 100, 26, "Nova meta", "NovaMeta", cCard, cTexto, 9, True

    With ws.ListObjects("tbMetas").ListColumns("% concluído").DataBodyRange
        .FormatConditions.Delete
        Set fc = .FormatConditions.AddDatabar
        fc.BarColor.Color = cTeal
        fc.MinPoint.Modify xlConditionValueNumber, 0
        fc.MaxPoint.Modify xlConditionValueNumber, 1
    End With
    With ws.ListObjects("tbMetas").ListColumns("Situação").DataBodyRange.FormatConditions
        .Delete
        With .Add(xlCellValue, xlEqual, "=""Atrasada""")
            .Font.Color = cRuim: .Font.Bold = True
        End With
        With .Add(xlCellValue, xlEqual, "=""Concluída""")
            .Font.Color = cBom: .Font.Bold = True
        End With
    End With
    ws.Range("B9").Select
End Sub

Public Sub RegistrarAporte()
    Dim lo As ListObject, ll As ListObject, meta As String, r As Variant, v As Variant, d As Variant, cod As Variant, ini As Long
    Set lo = ThisWorkbook.Worksheets(SH_MET).ListObjects("tbMetas")
    Set ll = ThisWorkbook.Worksheets(SH_LANC).ListObjects("tbLancamentos")
    meta = CStr(ThisWorkbook.Names("metSel").RefersToRange.Value)
    v = ThisWorkbook.Names("metValor").RefersToRange.Value
    d = ThisWorkbook.Names("metData").RefersToRange.Value
    If Len(meta) = 0 Then MsgBox "Escolha a meta.", vbExclamation: Exit Sub
    r = Application.Match(meta, lo.ListColumns("Meta").DataBodyRange, 0)
    If IsError(r) Then MsgBox "Meta não encontrada.", vbExclamation: Exit Sub
    cod = lo.ListColumns("Código").DataBodyRange.Cells(r, 1).Value
    If Len(CStr(cod)) = 0 Then MsgBox "Esta meta não tem categoria (Código) vinculada.", vbExclamation: Exit Sub
    If Not IsNumeric(v) Or Len(CStr(v)) = 0 Then MsgBox "Informe o valor do aporte.", vbExclamation: Exit Sub
    If CDbl(v) <= 0 Then MsgBox "O valor deve ser maior que zero.", vbExclamation: Exit Sub
    If Not IsDate(d) Then MsgBox "Informe a data.", vbExclamation: Exit Sub
    If Year(d) <> ThisWorkbook.Names("Ano").RefersToRange.Value Then MsgBox "A data precisa ser do exercício " & ThisWorkbook.Names("Ano").RefersToRange.Value & ".", vbExclamation: Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    ini = AcrescentarLinhas(ll, 1)
    With ll.ListRows(ini).Range
        .Cells(1, ll.ListColumns("Data").Index).Value = CDate(d)
        .Cells(1, ll.ListColumns("Código").Index).Value = cod
        .Cells(1, ll.ListColumns("Descrição").Index).Value = "Aporte: " & meta
        .Cells(1, ll.ListColumns("Valor").Index).Value = Round(CDbl(v), 2)
        .Cells(1, ll.ListColumns("Pagamento").Index).Value = "Transferência"
        .Cells(1, ll.ListColumns("Responsável").Index).Value = "Casa"
        .Cells(1, ll.ListColumns("Origem").Index).Value = "Metas"
    End With
    ThisWorkbook.Names("metValor").RefersToRange.ClearContents
    Application.EnableEvents = True
    AtualizarPainel
    Application.ScreenUpdating = True
    Status "Aporte de " & Format(v, """R$ ""#,##0.00") & " em " & meta & "."
End Sub

Public Sub NovaMeta()
    Dim lo As ListObject, r As Long
    Set lo = ThisWorkbook.Worksheets(SH_MET).ListObjects("tbMetas")
    Application.EnableEvents = False
    r = AcrescentarLinhas(lo, 1)
    lo.ListRows(r).Range.Cells(1, lo.ListColumns("Prazo").Index).Value = DateSerial(ThisWorkbook.Names("Ano").RefersToRange.Value, 12, 31)
    Application.EnableEvents = True
    lo.ListRows(r).Range.Cells(1, 1).Select
End Sub

'===============================================================================
' ABAS DE DADOS
'===============================================================================
Private Sub ArrumarAbasDeDados()
    Dim nomes As Variant, i As Long, ws As Worksheet, lo As ListObject, t As Range
    nomes = Array(SH_LANC, SH_ORC, SH_REND, SH_CAT)
    For i = 0 To UBound(nomes)
        Set ws = ThisWorkbook.Worksheets(nomes(i))
        LimparFormas ws
        ws.Activate
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        ws.Cells.Font.Name = FONTE
        ws.Range("A1").ClearContents
        ws.Rows(1).RowHeight = 44
        ws.Rows(2).RowHeight = 30
        For Each lo In ws.ListObjects
            lo.TableStyle = "CO Tabela"
        Next lo
        Cabecalho ws, CStr(nomes(i))
        If nomes(i) = SH_LANC Then
            ws.Rows(3).RowHeight = 34
            BotaoForma ws, 12, 79, 190, 26, "Importar extrato (Python)", "ImportarExtrato", cPrim, RGB(255, 255, 255), 9
            BotaoForma ws, 210, 79, 170, 26, "Auditar dados (Python)", "AuditarDados", cCard, cTexto, 9, True
        End If
    Next i
    ' Orçamento: cabeçalho e células de digitação na identidade executiva
    Set ws = ThisWorkbook.Worksheets(SH_ORC)
    Set t = ws.Columns(3).Find(What:="TOTAL PREVISTO", LookIn:=xlValues, LookAt:=xlWhole)
    With ws.Range("A4:P4")
        .Interior.Color = cHeader: .Font.Color = RGB(255, 255, 255): .Font.Bold = True
    End With
    If Not t Is Nothing Then
        ws.Range(ws.Cells(5, 4), ws.Cells(t.Row - 1, 15)).Interior.Color = cInput
        ws.Range(ws.Cells(5, 4), ws.Cells(t.Row - 1, 15)).Font.Color = cPrim
        ws.Range(ws.Cells(t.Row, 3), ws.Cells(t.Row, 16)).Interior.Color = RGB(226, 232, 240)
        ws.Range(ws.Cells(5, 1), ws.Cells(t.Row, 16)).Borders.Color = cLinha
    End If
    ' lista de grupos em Categorias
    Set t = ThisWorkbook.Worksheets(SH_CAT).Rows(3).Find(What:="Grupos", LookIn:=xlValues, LookAt:=xlWhole)
    If Not t Is Nothing Then
        t.Interior.Color = cHeader: t.Font.Color = RGB(255, 255, 255)
    End If
End Sub

' Regras de categorização automática do extrato (palavra-chave -> código), editáveis pela família.
Private Sub PrepararRegras(hGrupos As Range)
    Dim ws As Worksheet, lo As ListObject, c As Range, base As Variant, i As Long, n As Long, cods As Range
    Set ws = hGrupos.Worksheet
    On Error Resume Next
    Set lo = ws.ListObjects("tbRegras")
    On Error GoTo 0
    If lo Is Nothing Then
        Set c = ws.Cells(3, hGrupos.Column + 2)
        c.Resize(1, 3).Value = Array("Palavra-chave", "Código", "Categoria")
        Set lo = ws.ListObjects.Add(xlSrcRange, c.Resize(2, 3), , xlYes)
        lo.Name = "tbRegras"
        Set cods = ws.ListObjects("tbCategorias").ListColumns("Código").DataBodyRange
        base = Array("IFOOD", 14, "RAPPI", 14, "UBER", 40, "99APP", 40, "99 POP", 40, "NETFLIX", 18, "PRIME VIDEO", 19, "AMAZON PRIME", 19, _
                     "HBO", 20, "MAX.COM", 20, "SPOTIFY", 17, "DEEZER", 17, "YOUTUBE", 17, "SUPERMERC", 11, "ATACAD", 11, "ASSAI", 11, _
                     "CARREFOUR", 11, "PAO DE ACUCAR", 11, "HORTIFRUTI", 11, "ACOUGUE", 11, "PADARIA", 13, "LANCHONETE", 13, "RESTAURANTE", 12, _
                     "DROGA", 27, "FARMACIA", 27, "DROGASIL", 27, "RAIA", 27, "ENEL", 6, "CEMIG", 6, "CPFL", 6, "COPEL", 6, "ENERGISA", 6, _
                     "SABESP", 7, "COPASA", 7, "SANEPAR", 7, "CEDAE", 7, "ULTRAGAZ", 8, "LIQUIGAS", 8, "SUPERGASBRAS", 8, "VIVO", 15, "CLARO", 15, _
                     "SMART FIT", 21, "SMARTFIT", 21, "ACADEMIA", 21, "CINEMARK", 60, "CINEMA", 60, "INGRESSO", 60, "RENNER", 31, "RIACHUELO", 31, "ZARA", 31)
        For i = 0 To UBound(base) Step 2
            If IsNumeric(Application.Match(base(i + 1), cods, 0)) Then
                n = n + 1
                If n > 1 Then lo.ListRows.Add
                lo.ListRows(n).Range.Cells(1, 1).Value = base(i)
                lo.ListRows(n).Range.Cells(1, 2).Value = base(i + 1)
            End If
        Next i
    End If
    lo.ListColumns("Categoria").DataBodyRange.Formula = "=IF([@Código]="""","""",IFERROR(INDEX(tbCategorias[Subcategoria],MATCH([@Código],tbCategorias[Código],0)),""CÓDIGO INVÁLIDO""))"
    lo.TableStyle = "CO Tabela"
    ws.Columns(lo.Range.Column).ColumnWidth = 18
    ws.Columns(lo.Range.Column + 2).ColumnWidth = 24
End Sub

'===============================================================================
' BACK-END PYTHON (processamento em lote com pandas)
'===============================================================================
Public Sub ImportarExtrato()
    Dim f As Variant, pasta As String, res As Object, ll As ListObject, linhas As Variant, n As Long, i As Long, ini As Long
    Dim cols() As String, dts() As Variant, cds() As Variant, dsc() As Variant, vls() As Variant, org() As Variant, padrao As Long
    f = Application.GetOpenFilename("Extratos bancários (*.csv;*.xlsx;*.ofx),*.csv;*.xlsx;*.ofx", , "Importar extrato")
    If VarType(f) = vbBoolean Then Exit Sub
    ImportarArquivo CStr(f)
End Sub

Public Function ImportarArquivo(f As String) As String
    Dim pasta As String, res As Object, ll As ListObject, linhas As Variant, n As Long, i As Long, ini As Long
    Dim cols() As String, dts() As Variant, cds() As Variant, dsc() As Variant, vls() As Variant, org() As Variant, padrao As Long, lc As ListObject
    Set lc = ThisWorkbook.Worksheets(SH_CAT).ListObjects("tbCategorias")
    padrao = 10
    If IsError(Application.Match(10, lc.ListColumns("Código").DataBodyRange, 0)) Then padrao = lc.ListColumns("Código").DataBodyRange.Cells(1, 1).Value
    pasta = PastaTemp()
    Application.StatusBar = "Python: importando extrato..."
    ExportarLancamentos pasta & "\lanc.csv"
    ExportarTabela ThisWorkbook.Worksheets(SH_CAT).ListObjects("tbRegras"), Array("Palavra-chave", "Código"), "palavra;codigo", pasta & "\regras.csv"
    ExportarTabela lc, Array("Código"), "codigo", pasta & "\cats.csv"
    Set res = RodarPython("importar --extrato """ & f & """ --lanc """ & pasta & "\lanc.csv"" --regras """ & pasta & "\regras.csv"" --cats """ & pasta & "\cats.csv"" --saida """ & pasta & "\novos.csv"" --ano " & ThisWorkbook.Names("Ano").RefersToRange.Value & " --padrao " & padrao, pasta)
    Application.StatusBar = False
    If res Is Nothing Then GoTo fim
    If res.Exists("erro") Then
        ImportarArquivo = "erro: " & res("erro")
        If Application.Interactive Then MsgBox "Não foi possível importar: " & res("erro"), vbExclamation, "Importar extrato"
        GoTo fim
    End If

    linhas = LerLinhas(pasta & "\novos.csv")
    n = 0
    If IsArray(linhas) Then n = UBound(linhas) + 1
    If n > 0 Then
        ReDim dts(1 To n, 1 To 1): ReDim cds(1 To n, 1 To 1): ReDim dsc(1 To n, 1 To 1): ReDim vls(1 To n, 1 To 1): ReDim org(1 To n, 1 To 1)
        For i = 1 To n
            cols = Split(linhas(i - 1), ";")
            dts(i, 1) = DateSerial(CLng(Left(cols(0), 4)), CLng(Mid(cols(0), 6, 2)), CLng(Right(cols(0), 2)))
            cds(i, 1) = CLng(cols(1)): dsc(i, 1) = cols(2): vls(i, 1) = Val(cols(3)): org(i, 1) = "Importado"
        Next i
        Set ll = ThisWorkbook.Worksheets(SH_LANC).ListObjects("tbLancamentos")
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        ini = AcrescentarLinhas(ll, n)
        ll.ListColumns("Data").DataBodyRange.Cells(ini, 1).Resize(n, 1).Value = dts
        ll.ListColumns("Código").DataBodyRange.Cells(ini, 1).Resize(n, 1).Value = cds
        ll.ListColumns("Descrição").DataBodyRange.Cells(ini, 1).Resize(n, 1).Value = dsc
        ll.ListColumns("Valor").DataBodyRange.Cells(ini, 1).Resize(n, 1).Value = vls
        ll.ListColumns("Origem").DataBodyRange.Cells(ini, 1).Resize(n, 1).Value = org
        Application.EnableEvents = True
        AtualizarPainel
        Application.ScreenUpdating = True
    End If
    ImportarArquivo = "novos=" & n & ";duplicados=" & res("duplicados") & ";revisar=" & res("revisar") & ";creditos=" & res("creditos") & ";fora=" & res("fora_do_ano")
    Status "Extrato: " & n & " lançamento(s) novo(s), " & res("duplicados") & " duplicado(s) ignorado(s), " & res("revisar") & " para revisar."
    If Application.Interactive Then
        MsgBox "Extrato importado." & vbCrLf & vbCrLf & _
               "Lançamentos novos: " & n & vbCrLf & _
               "Duplicados ignorados: " & res("duplicados") & vbCrLf & _
               "Sem categoria (marcados [revisar]): " & res("revisar") & vbCrLf & _
               "Créditos ignorados: " & res("creditos") & vbCrLf & _
               "Fora do exercício: " & res("fora_do_ano"), vbInformation, "Importar extrato"
    End If
fim:
    Application.StatusBar = False
    LimparPastaTemp pasta
End Function

Public Sub AuditarDados()
    AuditarExecutar
End Sub

Public Function AuditarExecutar() As String
    Dim pasta As String, res As Object, ll As ListObject, v As Variant, msg As String, nd As Long, na As Long
    pasta = PastaTemp()
    Application.StatusBar = "Python: auditando lançamentos..."
    ExportarLancamentos pasta & "\lanc.csv"
    Set res = RodarPython("auditar --lanc """ & pasta & "\lanc.csv""", pasta)
    Application.StatusBar = False
    If res Is Nothing Then GoTo fim
    If res.Exists("erro") Then
        AuditarExecutar = "erro: " & res("erro")
        If Application.Interactive Then MsgBox "Não foi possível auditar: " & res("erro"), vbExclamation, "Auditoria"
        GoTo fim
    End If
    Set ll = ThisWorkbook.Worksheets(SH_LANC).ListObjects("tbLancamentos")
    Application.ScreenUpdating = False
    ll.DataBodyRange.Interior.Pattern = xlNone
    If Len(res("duplicados")) > 0 Then
        For Each v In Split(res("duplicados"), ",")
            ll.ListRows(CLng(v)).Range.Interior.Color = RGB(254, 243, 199): nd = nd + 1
        Next v
    End If
    If Len(res("atipicos")) > 0 Then
        For Each v In Split(res("atipicos"), ",")
            ll.ListRows(CLng(v)).Range.Interior.Color = RGB(254, 226, 226): na = na + 1
        Next v
    End If
    Application.ScreenUpdating = True
    msg = "Lançamentos analisados: " & res("total") & vbCrLf & vbCrLf & _
          "Possíveis duplicados (amarelo): " & nd & vbCrLf & _
          "Valores atípicos (vermelho): " & na & vbCrLf & _
          "Sem forma de pagamento: " & res("sem_pagamento") & vbCrLf & _
          "Sem responsável: " & res("sem_responsavel")
    If Len(res("variacoes")) > 0 Then msg = msg & vbCrLf & vbCrLf & "Maiores variações no último mês:" & vbCrLf & Replace(res("variacoes"), " | ", vbCrLf)
    AuditarExecutar = "total=" & res("total") & ";duplicados=" & nd & ";atipicos=" & na & ";sem_pag=" & res("sem_pagamento") & ";variacoes=" & res("variacoes")
    Status "Auditoria: " & nd & " duplicado(s), " & na & " atípico(s)."
    If Application.Interactive Then MsgBox msg, vbInformation, "Auditoria dos dados"
fim:
    LimparPastaTemp pasta
End Function

' Executa o back-end Python sem abrir janela e devolve as linhas "chave=valor" do resultado.
Private Function RodarPython(args As String, pasta As String) As Object
    Dim sh As Object, script As String, saida As String, rc As Long, linhas As Variant, i As Long, p As Long, d As Object, exe As Variant, tudo As String
    script = ThisWorkbook.Path & "\backend\casa_backend.py"
    Set d = CreateObject("Scripting.Dictionary")
    If Len(Dir(script)) = 0 Then d("erro") = "back-end Python não encontrado em " & script: Set RodarPython = d: Exit Function
    saida = pasta & "\resultado.txt"
    Set sh = CreateObject("WScript.Shell")
    For Each exe In Array("python", "py -3")
        rc = sh.Run("cmd /c " & exe & " """ & script & """ " & args & " > """ & saida & """ 2>&1", 0, True)
        linhas = LerLinhas(saida)
        tudo = ""
        If IsArray(linhas) Then tudo = Join(linhas, " ")
        If InStr(1, tudo, "reconhecido", vbTextCompare) = 0 And InStr(1, tudo, "not recognized", vbTextCompare) = 0 Then Exit For
    Next exe
    If IsArray(linhas) Then
        For i = 0 To UBound(linhas)
            p = InStr(1, linhas(i), "=")
            If p > 1 Then d(Left(linhas(i), p - 1)) = Mid(linhas(i), p + 1)
        Next i
    End If
    If rc <> 0 And Not d.Exists("erro") Then
        If Len(tudo) > 0 Then d("erro") = Left(tudo, 400) Else d("erro") = "Python não respondeu (código " & rc & ")."
    End If
    For Each exe In Array("duplicados", "revisar", "creditos", "fora_do_ano", "atipicos", "sem_pagamento", "sem_responsavel", "variacoes", "total")
        If Not d.Exists(exe) Then d(exe) = ""
    Next exe
    Set RodarPython = d
End Function

' Exporta os lançamentos para o Python em um CSV temporário (apagado ao final).
Private Sub ExportarLancamentos(caminho As String)
    Dim ll As ListObject, i As Long, f As Integer, a As Variant, cD As Long, cC As Long, cS As Long, cV As Long, cB As Long, cP As Long, cR As Long
    Set ll = ThisWorkbook.Worksheets(SH_LANC).ListObjects("tbLancamentos")
    f = FreeFile
    Open caminho For Output As #f
    Print #f, "data;codigo;descricao;valor;subcategoria;pagamento;responsavel"
    If Not ll.DataBodyRange Is Nothing Then
        a = ll.DataBodyRange.Value
        cD = ll.ListColumns("Data").Index: cC = ll.ListColumns("Código").Index: cS = ll.ListColumns("Descrição").Index
        cV = ll.ListColumns("Valor").Index: cB = ll.ListColumns("Subcategoria").Index
        cP = ll.ListColumns("Pagamento").Index: cR = ll.ListColumns("Responsável").Index
        For i = 1 To UBound(a, 1)
            Print #f, IIf(IsDate(a(i, cD)), Format(a(i, cD), "yyyy-mm-dd"), "") & ";" & Limpo(a(i, cC)) & ";" & Limpo(a(i, cS)) & ";" & _
                      NumTexto(a(i, cV)) & ";" & Limpo(a(i, cB)) & ";" & Limpo(a(i, cP)) & ";" & Limpo(a(i, cR))
        Next i
    End If
    Close #f
End Sub

Private Function NumTexto(v As Variant) As String
    If IsNumeric(v) And Not IsEmpty(v) Then NumTexto = Replace(CStr(CDbl(v)), ",", ".") Else NumTexto = ""
End Function

Private Sub ExportarTabela(lo As ListObject, colunas As Variant, cabecalho As String, caminho As String)
    Dim f As Integer, a As Variant, i As Long, j As Long, lin As String, idx() As Long
    ReDim idx(0 To UBound(colunas))
    For j = 0 To UBound(colunas): idx(j) = lo.ListColumns(colunas(j)).Index: Next j
    f = FreeFile
    Open caminho For Output As #f
    Print #f, cabecalho
    If Not lo.DataBodyRange Is Nothing Then
        a = Col2D(lo.DataBodyRange)
        For i = 1 To UBound(a, 1)
            lin = ""
            For j = 0 To UBound(idx): lin = lin & IIf(j > 0, ";", "") & Limpo(a(i, idx(j))): Next j
            Print #f, lin
        Next i
    End If
    Close #f
End Sub

Private Function Limpo(v As Variant) As String
    If IsError(v) Then Exit Function
    Limpo = Replace(Replace(Replace(CStr(v), ";", ","), vbCr, " "), vbLf, " ")
End Function

Private Function LerLinhas(caminho As String) As Variant
    Dim f As Integer, txt As String, t As String
    If Len(Dir(caminho)) = 0 Then Exit Function
    f = FreeFile
    Open caminho For Input As #f
    Do While Not EOF(f)
        Line Input #f, t
        If Len(Trim(t)) > 0 Then txt = txt & IIf(Len(txt) > 0, vbLf, "") & t
    Loop
    Close #f
    If Len(txt) > 0 Then LerLinhas = Split(txt, vbLf)
End Function

Private Function PastaTemp() As String
    PastaTemp = Environ("TEMP") & "\casa_organizada_" & Format(Now, "yyyymmddhhnnss")
    On Error Resume Next
    MkDir PastaTemp
End Function

Private Sub LimparPastaTemp(pasta As String)
    On Error Resume Next
    Kill pasta & "\*.*"
    RmDir pasta
End Sub

Private Sub OrdenarAbas()
    Dim ordem As Variant, i As Long
    ordem = Array(SH_DASH, SH_ANA, SH_CAD, SH_CON, SH_MET, SH_LANC, SH_ORC, SH_REND, SH_CAT, SH_CALC, SH_PIV)
    On Error Resume Next
    For i = UBound(ordem) To 0 Step -1
        ThisWorkbook.Worksheets(ordem(i)).Move Before:=ThisWorkbook.Worksheets(1)
    Next i
    ThisWorkbook.Worksheets(SH_CALC).Visible = xlSheetHidden
    ThisWorkbook.Worksheets(SH_PIV).Visible = xlSheetHidden
    ThisWorkbook.Worksheets(SH_CON).Tab.Color = cGold
    ThisWorkbook.Worksheets(SH_MET).Tab.Color = cGold
    ThisWorkbook.Worksheets(SH_DASH).Tab.Color = cHeader
    ThisWorkbook.Worksheets(SH_ANA).Tab.Color = cHeader
    ThisWorkbook.Worksheets(SH_CAD).Tab.Color = cHeader
End Sub

'===============================================================================
' COMPONENTES VISUAIS
'===============================================================================
Private Sub Cartao(ws As Worksheet, x As Single, y As Single, w As Single, h As Single, nome As String)
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    With shp
        .Name = nome
        .Adjustments.Item(1) = 0.05
        .Fill.ForeColor.RGB = cCard
        .Line.ForeColor.RGB = cLinha: .Line.Weight = 0.75
        .Shadow.Visible = msoTrue
        .Shadow.OffsetX = 0: .Shadow.OffsetY = 1: .Shadow.Blur = 4
        .Shadow.ForeColor.RGB = RGB(0, 0, 0): .Shadow.Transparency = 0.92
        .Placement = xlFreeFloating
    End With
End Sub

Private Sub Kpi(ws As Worksheet, x As Single, y As Single, w As Single, h As Single, titulo As String, celValor As String, celDelta As String, chave As String)
    Dim shp As Shape, alto As Boolean
    alto = (h >= 80)
    Cartao ws, x, y, w, h, "kpiCard" & chave
    Texto ws, x + 14, y + 10, w - 28, 13, titulo, 7.5, cMuted, True
    Set shp = Texto(ws, x + 14, y + IIf(alto, 26, 23), w - 28, 28, "", IIf(alto, 20, 17), cTexto, True, "='" & SH_CALC & "'!" & celValor)
    shp.Name = "kpiValor" & chave
    Set shp = Texto(ws, x + 14, y + h - 22, w - 28, 13, "", 8, cMuted, False, "='" & SH_CALC & "'!" & celDelta)
    shp.Name = "kpiDelta" & chave
End Sub

Private Function Texto(ws As Worksheet, x As Single, y As Single, w As Single, h As Single, txt As String, tam As Single, cor As Long, negrito As Boolean, Optional vinculo As String = "") As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    With shp
        .Line.Visible = msoFalse
        .Fill.Visible = msoFalse
        If Len(vinculo) > 0 Then .DrawingObject.Formula = vinculo
        With .TextFrame2
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .WordWrap = msoFalse
            .AutoSize = msoAutoSizeNone
            .VerticalAnchor = msoAnchorTop
            If Len(vinculo) = 0 Then .TextRange.Text = txt
            With .TextRange.Font
                .Name = FONTE
                .Size = tam
                .Bold = IIf(negrito, msoTrue, msoFalse)
                .Fill.ForeColor.RGB = cor
            End With
        End With
        .Placement = xlFreeFloating
        .Locked = True
    End With
    Set Texto = shp
End Function

Private Function BotaoForma(ws As Worksheet, x As Single, y As Single, w As Single, h As Single, txt As String, macro As String, fundo As Long, corTexto As Long, tam As Single, Optional contorno As Boolean = False) As Shape
    Dim shp As Shape
    Set shp = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    With shp
        .Name = "bt" & macro & ws.Shapes.Count
        .OnAction = "'" & ThisWorkbook.Name & "'!" & macro
        .Adjustments.Item(1) = 0.2
        .Fill.ForeColor.RGB = fundo
        If contorno Then
            .Line.Visible = msoTrue: .Line.ForeColor.RGB = cLinha: .Line.Weight = 0.75
        Else
            .Line.Visible = msoFalse
        End If
        With .TextFrame2
            .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoFalse
            With .TextRange
                .Text = txt
                .ParagraphFormat.Alignment = msoAlignCenter
                .Font.Name = FONTE: .Font.Size = tam: .Font.Bold = msoTrue
                .Font.Fill.ForeColor.RGB = corTexto
            End With
        End With
        .Placement = xlFreeFloating
    End With
    Set BotaoForma = shp
End Function

Private Sub EstiloGrafico(ch As ChartObject, titulo As String, legenda As Boolean)
    Dim ax As Axis
    ch.RoundedCorners = False
    ch.Placement = xlFreeFloating
    With ch.Chart
        On Error Resume Next
        .ShowAllFieldButtons = False
        On Error GoTo 0
        .ChartArea.Format.Fill.ForeColor.RGB = cCard
        .ChartArea.Format.Line.ForeColor.RGB = cLinha
        .ChartArea.Format.Line.Weight = 0.75
        .ChartArea.Font.Name = FONTE
        .ChartArea.Font.Size = 9
        .ChartArea.Font.Color = cMuted
        .PlotArea.Format.Fill.Visible = msoFalse
        .HasTitle = True
        .ChartTitle.Text = titulo
        With .ChartTitle.Format.TextFrame2.TextRange.Font
            .Name = FONTE: .Size = 11: .Bold = msoTrue: .Fill.ForeColor.RGB = cTexto
        End With
        On Error Resume Next
        .ChartTitle.Left = 12
        .ChartTitle.Top = 8
        On Error GoTo 0
        .HasLegend = legenda
        If legenda Then
            .Legend.Position = xlLegendPositionTop
            .Legend.Font.Size = 9
            .Legend.Font.Color = cMuted
        End If
        On Error Resume Next
        For Each ax In .Axes
            ax.Format.Line.Visible = msoFalse
            ax.MajorTickMark = xlNone
            ax.TickLabels.Font.Color = cMuted
            ax.TickLabels.Font.Size = 9
        Next ax
        .Axes(xlValue).HasMajorGridlines = True
        .Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = cLinha
        .Axes(xlValue).MajorGridlines.Format.Line.Weight = 0.5
        .Axes(xlValue).TickLabels.NumberFormatLocal = FormatoMoedaLocal()
        .Axes(xlValue).MinimumScale = 0
        On Error GoTo 0
    End With
    With ch.ShapeRange.Shadow
        .Visible = msoTrue: .OffsetX = 0: .OffsetY = 1: .Blur = 4
        .ForeColor.RGB = RGB(0, 0, 0): .Transparency = 0.92
    End With
End Sub

Private Function FormatoMoedaLocal() As String
    FormatoMoedaLocal = """R$ ""#" & Application.International(xlThousandsSeparator) & "##0"
End Function

Private Sub Painel(ws As Worksheet, ende As String, titulo As String)
    Dim b As Variant, r As Range
    Set r = ws.Range(ende)
    r.Interior.Color = cCard
    r.Borders.LineStyle = xlNone
    For Each b In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        With r.Borders(b)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .Color = RGB(203, 213, 225)
        End With
    Next b
    With r.Rows(1)
        .Borders(xlEdgeTop).Color = cPrim
        .Borders(xlEdgeTop).Weight = xlThick
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Weight = xlHairline
        .Borders(xlEdgeBottom).Color = cLinha
    End With
    With r.Cells(1, 1)
        .Value = titulo
        .Font.Size = 11.5: .Font.Bold = True: .Font.Color = cTexto
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
    End With
End Sub

Private Sub Campo(ws As Worksheet, linha As Long, col As Long, rotulo As String, nome As String, Optional formato As String = "@")
    With ws.Cells(linha, col)
        .Value = rotulo
        .Font.Size = 9: .Font.Color = cMuted
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
    End With
    With ws.Cells(linha, col + 1)
        .Interior.Color = cInput
        .Borders.Color = RGB(209, 213, 219)
        .Font.Color = cTexto
        .NumberFormat = formato
        .VerticalAlignment = xlCenter
        .Locked = False
    End With
    NomeFaixa nome, ws.Cells(linha, col + 1)
End Sub

Private Sub Lista(alvo As Range, fonteLista As String, Optional bloqueia As Boolean = True)
    With alvo.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=IIf(bloqueia, xlValidAlertStop, xlValidAlertInformation), Formula1:=fonteLista
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = bloqueia
        .ErrorMessage = "Escolha um item da lista."
    End With
End Sub

'===============================================================================
' ATUALIZAÇÃO, FILTROS E ALERTAS
'===============================================================================
Public Sub AtualizarPainel()
    Dim pc As PivotCache, calc As XlCalculation
    On Error Resume Next
    gOcupado = True
    Application.EnableEvents = False
    calc = Application.Calculation
    Application.Calculation = xlCalculationManual
    For Each pc In ThisWorkbook.PivotCaches
        pc.Refresh
    Next pc
    Application.Calculation = calc
    Application.EnableEvents = True
    gOcupado = False
    SincronizarFiltros
End Sub

' Agrupa várias edições seguidas numa única atualização (alto desempenho ao colar dados).
Public Sub AgendarAtualizacao()
    On Error Resume Next
    If gAgendado <> 0 Then Application.OnTime gAgendado, "modCasaOrganizada.AtualizarPainel", , False
    gAgendado = Now + TimeSerial(0, 0, 1)
    Application.OnTime gAgendado, "modCasaOrganizada.AtualizarPainel"
End Sub

' Lê as segmentações de Mês e Grupo e grava os filtros na aba Calc (indicadores e gráficos acompanham).
Public Sub SincronizarFiltros()
    Dim wc As Worksheet, sc As SlicerCache, si As SlicerItem, r As Long, m As Variant, cap As String
    Dim fm(1 To 12, 1 To 1) As Variant, fg(1 To 20, 1 To 1) As Variant, nomesG As Variant, refs As Variant
    Dim txt As String, nSel As Long, nGr As Long, nGrSel As Long, filtrouM As Boolean, filtrouG As Boolean
    On Error Resume Next
    Set wc = ThisWorkbook.Worksheets(SH_CALC)
    If wc Is Nothing Then Exit Sub
    Application.EnableEvents = False
    refs = wc.Range("C5:C16").Value
    nomesG = wc.Range("B21:B40").Value
    For r = 1 To 12: fm(r, 1) = 1: Next r
    For r = 1 To 20: fg(r, 1) = IIf(Len(nomesG(r, 1)) = 0, 0, 1): Next r
    For Each sc In ThisWorkbook.SlicerCaches
        cap = "": cap = sc.Slicers(1).Caption
        If cap = "Mês" And Not sc.FilterCleared Then
            filtrouM = True
            For r = 1 To 12: fm(r, 1) = 0: Next r
            For Each si In sc.SlicerItems
                If si.Selected Then
                    m = Application.Match(si.Name, wc.Range("C5:C16"), 0)
                    If Not IsError(m) Then fm(m, 1) = 1
                End If
            Next si
        ElseIf cap = "Grupo" And Not sc.FilterCleared Then
            filtrouG = True
            For r = 1 To 20: fg(r, 1) = 0: Next r
            For Each si In sc.SlicerItems
                If si.Selected Then
                    m = Application.Match(si.Name, wc.Range("B21:B40"), 0)
                    If Not IsError(m) Then fg(m, 1) = 1
                End If
            Next si
        End If
    Next sc
    wc.Range("D5:D16").Value = fm
    wc.Range("C21:C40").Value = fg

    nSel = 0: For r = 1 To 12: nSel = nSel + fm(r, 1): Next r
    If nSel = 12 Or nSel = 0 Then
        txt = "Exercício " & wc.Range("B1").Value
    Else
        For r = 1 To 12
            If fm(r, 1) = 1 Then txt = txt & IIf(Len(txt) > 0, ", ", "") & wc.Cells(4 + r, 2).Value
        Next r
        txt = txt & " de " & wc.Range("B1").Value
    End If
    nGr = wc.Range("B19").Value
    nGrSel = 0: For r = 1 To 20: nGrSel = nGrSel + fg(r, 1): Next r
    If nGrSel < nGr Then txt = txt & "  ·  " & nGrSel & " de " & nGr & " grupos"
    wc.Range("C58").Value = txt
    Application.Calculate
    Application.EnableEvents = True
    AtualizarVisual
End Sub

' Cores semânticas nos cartões (verde bom, âmbar atenção, vermelho ruim) e no selo de consistência.
Public Sub AtualizarVisual()
    Dim wc As Worksheet, ws As Worksheet, shp As Shape, chave As String, st As Variant, cor As Long, lin As String
    On Error Resume Next
    Cores
    Set wc = ThisWorkbook.Worksheets(SH_CALC)
    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            If Left(shp.Name, 8) = "kpiDelta" Then
                chave = Mid(shp.Name, 9)
                lin = chave: If Not IsNumeric(Left(lin, 1)) Then lin = Mid(lin, 2)
                st = wc.Range("D" & lin).Value
                Select Case st
                    Case 1: cor = cBom
                    Case -1: cor = cRuim
                    Case 2: cor = cAlerta
                    Case Else: cor = cMuted
                End Select
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = cor
            ElseIf shp.Name = "kpiValor49" Then
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(wc.Range("B49").Value < 0, cRuim, cTexto)
            ElseIf shp.Name = "kpiValorc61" Then
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(wc.Range("B62").Value > 0, cRuim, cTexto)
            ElseIf shp.Name = "chipStatus" Then
                shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = IIf(wc.Range("D56").Value = 1, RGB(134, 239, 172), RGB(252, 165, 165))
            End If
        Next shp
    Next ws
End Sub

Public Sub LimparFiltros()
    Dim sc As SlicerCache
    On Error Resume Next
    gOcupado = True
    Application.ScreenUpdating = False
    For Each sc In ThisWorkbook.SlicerCaches
        sc.ClearManualFilter
    Next sc
    gOcupado = False
    SincronizarFiltros
    Application.ScreenUpdating = True
End Sub

'===============================================================================
' AÇÕES DO LANÇAR
'===============================================================================
Public Sub LancarGasto()
    Dim ws As Worksheet, lo As ListObject, r As Long
    Dim d As Variant, cod As Long, v As Variant, txtCat As String, ano As Long
    Set ws = ThisWorkbook.Worksheets(SH_CAD)
    d = ws.Range("frmData").Value
    txtCat = CStr(ws.Range("frmCat").Value)
    v = ws.Range("frmValor").Value
    ano = ThisWorkbook.Names("Ano").RefersToRange.Value

    If Not IsDate(d) Then Aviso "Informe uma data válida.", ws.Range("frmData"): Exit Sub
    If Year(d) <> ano Then Aviso "A data precisa ser do exercício " & ano & ".", ws.Range("frmData"): Exit Sub
    If Len(txtCat) = 0 Then Aviso "Escolha a categoria.", ws.Range("frmCat"): Exit Sub
    cod = CodigoDaLista(txtCat)
    If cod = 0 Then Aviso "Categoria inválida. Escolha uma da lista.", ws.Range("frmCat"): Exit Sub
    If Not IsNumeric(v) Or Len(CStr(v)) = 0 Then Aviso "Informe o valor.", ws.Range("frmValor"): Exit Sub
    If CDbl(v) <= 0 Then Aviso "O valor deve ser maior que zero.", ws.Range("frmValor"): Exit Sub

    Set lo = ThisWorkbook.Worksheets(SH_LANC).ListObjects("tbLancamentos")
    If ExisteDuplicado(lo, CDate(d), cod, CDbl(v)) Then
        If MsgBox("Já existe um lançamento com mesma data, categoria e valor." & vbCrLf & "Lançar mesmo assim?", vbQuestion + vbYesNo, "Possível duplicidade") = vbNo Then Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    r = AcrescentarLinhas(lo, 1)
    With lo.ListRows(r).Range
        .Cells(1, lo.ListColumns("Data").Index).Value = CDate(d)
        .Cells(1, lo.ListColumns("Código").Index).Value = cod
        .Cells(1, lo.ListColumns("Descrição").Index).Value = ws.Range("frmDesc").Value
        .Cells(1, lo.ListColumns("Valor").Index).Value = Round(CDbl(v), 2)
        .Cells(1, lo.ListColumns("Pagamento").Index).Value = ws.Range("frmPag").Value
        .Cells(1, lo.ListColumns("Responsável").Index).Value = ws.Range("frmResp").Value
        .Cells(1, lo.ListColumns("Origem").Index).Value = "Manual"
    End With
    ws.Range("frmValor,frmDesc").ClearContents
    Application.EnableEvents = True
    AtualizarPainel
    Status "Lançado: " & txtCat & " · " & Format(v, """R$ ""#,##0.00") & " em " & Format(d, "dd/mm/yyyy")
    Application.ScreenUpdating = True
    On Error Resume Next
    If ActiveSheet.Name = ws.Name Then ws.Range("frmValor").Select
End Sub

Public Sub LimparLancamento()
    With ThisWorkbook.Worksheets(SH_CAD)
        .Range("frmCat,frmValor,frmDesc,frmPag,frmResp").ClearContents
        .Range("frmData").Value = DataPadrao()
    End With
    Status "Formulário limpo."
End Sub

Public Sub DesfazerUltimoLancamento()
    Dim lo As ListObject, n As Long, txt As String
    Set lo = ThisWorkbook.Worksheets(SH_LANC).ListObjects("tbLancamentos")
    n = lo.ListRows.Count
    If n = 0 Then Exit Sub
    With lo.ListRows(n).Range
        If .Cells(1, lo.ListColumns("Origem").Index).Value <> "Manual" Then
            MsgBox "O último lançamento não foi feito pela tela Lançar. Exclua pela aba Lançamentos.", vbInformation: Exit Sub
        End If
        txt = Format(.Cells(1, 1).Value, "dd/mm/yyyy") & " · " & .Cells(1, lo.ListColumns("Subcategoria").Index).Value & " · " & _
              Format(.Cells(1, lo.ListColumns("Valor").Index).Value, """R$ ""#,##0.00")
    End With
    If MsgBox("Excluir o último lançamento?" & vbCrLf & txt, vbQuestion + vbYesNo) = vbNo Then Exit Sub
    Application.EnableEvents = False
    lo.ListRows(n).Delete
    Application.EnableEvents = True
    AtualizarPainel
    Status "Excluído: " & txt
End Sub

Public Sub CadastrarCategoria()
    Dim ws As Worksheet, lo As ListObject, r As Long, wo As Worksheet, lg As Range
    Dim grupo As String, subc As String, prev As Double, cod As Long, c As Range, linha As Long, m As Long
    Set ws = ThisWorkbook.Worksheets(SH_CAD)
    Set wo = ThisWorkbook.Worksheets(SH_ORC)
    Set lo = ThisWorkbook.Worksheets(SH_CAT).ListObjects("tbCategorias")
    grupo = Trim(CStr(ws.Range("catGrupo").Value))
    subc = Trim(CStr(ws.Range("catSub").Value))
    If Len(grupo) = 0 Then Aviso "Informe o grupo.", ws.Range("catGrupo"): Exit Sub
    If Len(subc) = 0 Then Aviso "Informe a subcategoria.", ws.Range("catSub"): Exit Sub
    If Not IsEmpty(ws.Range("catPrev").Value) And Not IsNumeric(ws.Range("catPrev").Value) Then Aviso "Previsto inválido.", ws.Range("catPrev"): Exit Sub
    prev = Num(ws.Range("catPrev").Value)
    If prev < 0 Then Aviso "Previsto não pode ser negativo.", ws.Range("catPrev"): Exit Sub
    If Not IsError(Application.Match(subc, lo.ListColumns("Subcategoria").DataBodyRange, 0)) Then
        Aviso "Já existe a subcategoria """ & subc & """.", ws.Range("catSub"): Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    cod = Application.WorksheetFunction.Max(lo.ListColumns("Código").DataBodyRange) + 1
    r = AcrescentarLinhas(lo, 1)
    With lo.ListRows(r).Range
        .Cells(1, lo.ListColumns("Código").Index).Value = cod
        .Cells(1, lo.ListColumns("Grupo").Index).Value = grupo
        .Cells(1, lo.ListColumns("Subcategoria").Index).Value = subc
        .Cells(1, lo.ListColumns("Fixa").Index).Value = IIf(ws.Range("catFixa").Value = "Sim", "Sim", "Não")
    End With
    Set lg = ThisWorkbook.Names("ListaGrupos").RefersToRange
    If IsError(Application.Match(grupo, lg, 0)) Then
        Set c = lg.Cells(1, 1)
        Do While Len(c.Value) > 0: Set c = c.Offset(1, 0): Loop
        c.Value = grupo
    End If
    Set c = wo.Columns(3).Find(What:="TOTAL PREVISTO", LookIn:=xlValues, LookAt:=xlWhole)
    If Not c Is Nothing Then
        linha = c.Row - 1
        wo.Rows(linha).Copy
        wo.Rows(linha).Insert Shift:=xlDown
        Application.CutCopyMode = False
        wo.Cells(linha, 1).Value = cod
        wo.Cells(linha, 2).Formula = "=IFERROR(INDEX(Categorias!$B$4:$B$300,MATCH($A" & linha & ",Categorias!$A$4:$A$300,0)),""CÓDIGO INVÁLIDO"")"
        wo.Cells(linha, 3).Formula = "=IFERROR(INDEX(Categorias!$C$4:$C$300,MATCH($A" & linha & ",Categorias!$A$4:$A$300,0)),""CÓDIGO INVÁLIDO"")"
        For m = 4 To 15: wo.Cells(linha, m).Value = prev: Next m
        wo.Cells(linha, 16).Formula = "=SUM(D" & linha & ":O" & linha & ")"
    End If
    ws.Range("catSub,catPrev").ClearContents
    Application.EnableEvents = True
    AtualizarPainel
    Application.ScreenUpdating = True
    Status "Categoria " & cod & " - " & subc & " cadastrada em " & grupo & "."
End Sub

Public Sub RegistrarRenda()
    Dim ws As Worksheet, lo As ListObject, r As ListRow, nr As Long
    Dim m As Variant, fonteR As String, p As Variant, rl As Variant
    Set ws = ThisWorkbook.Worksheets(SH_CAD)
    m = ws.Range("rMes").Value: fonteR = Trim(CStr(ws.Range("rFonte").Value))
    p = ws.Range("rPrev").Value: rl = ws.Range("rReal").Value
    If Not IsNumeric(m) Or Len(CStr(m)) = 0 Then Aviso "Informe o mês (1 a 12).", ws.Range("rMes"): Exit Sub
    If m < 1 Or m > 12 Then Aviso "Mês deve ser de 1 a 12.", ws.Range("rMes"): Exit Sub
    If Len(fonteR) = 0 Then Aviso "Informe a fonte da renda.", ws.Range("rFonte"): Exit Sub
    If Len(CStr(p)) > 0 And Not IsNumeric(p) Then Aviso "Previsto inválido.", ws.Range("rPrev"): Exit Sub
    If Len(CStr(rl)) > 0 And Not IsNumeric(rl) Then Aviso "Real inválido.", ws.Range("rReal"): Exit Sub

    Set lo = ThisWorkbook.Worksheets(SH_REND).ListObjects("tbRendas")
    For Each r In lo.ListRows
        If r.Range.Cells(1, 1).Value = CLng(m) And LCase(r.Range.Cells(1, 2).Value) = LCase(fonteR) Then
            If MsgBox("Já existe """ & fonteR & """ no mês " & m & ". Atualizar os valores?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
            Application.EnableEvents = False
            r.Range.Cells(1, 3).Value = Num(p): r.Range.Cells(1, 4).Value = Num(rl)
            GoTo fim
        End If
    Next r
    Application.EnableEvents = False
    nr = AcrescentarLinhas(lo, 1)
    With lo.ListRows(nr).Range
        .Cells(1, 1).Value = CLng(m)
        .Cells(1, 2).Value = fonteR
        .Cells(1, 3).Value = Num(p)
        .Cells(1, 4).Value = Num(rl)
    End With
fim:
    ws.Range("rFonte,rPrev,rReal").ClearContents
    Application.EnableEvents = True
    AtualizarPainel
    Status "Renda """ & fonteR & """ registrada no mês " & m & "."
End Sub

'===============================================================================
' NAVEGAÇÃO
'===============================================================================
Public Sub IrCadastro(): Ir SH_CAD, "frmCat": AjustarZoom 890: End Sub
Public Sub IrDashboard(): Ir SH_DASH, "A1": AjustarZoom LARG + 10: End Sub
Public Sub IrLancamentos(): Ir SH_LANC, "A1": End Sub
Public Sub IrOrcamento(): Ir SH_ORC, "A1": End Sub
Public Sub IrRendas(): Ir SH_REND, "A1": End Sub
Public Sub IrContas(): Ir SH_CON, "conMes": End Sub

Private Sub Ir(aba As String, cel As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(aba).Activate
    ThisWorkbook.Worksheets(aba).Range(cel).Select
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

Public Sub AjustarZoom(largura As Single)
    Dim ws As Worksheet, c As Long, sel As Range
    On Error Resume Next
    Set ws = ActiveSheet
    Set sel = Selection
    c = 1
    Do While ws.Cells(1, c).Left + ws.Cells(1, c).Width < largura And c < 200: c = c + 1: Loop
    Application.ScreenUpdating = True
    ws.Range(ws.Cells(1, 1), ws.Cells(1, c)).Select
    ActiveWindow.Zoom = True
    If ActiveWindow.Zoom > 100 Then ActiveWindow.Zoom = 100
    If ActiveWindow.Zoom < 55 Then ActiveWindow.Zoom = 55
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    If Not sel Is Nothing And ws.Name = SH_CAD Then sel.Select Else ws.Range("A1").Select
End Sub

'===============================================================================
' EVENTOS (gravados no módulo EstaPasta_de_trabalho)
'===============================================================================
Private Sub InstalarEventos()
    Dim cm As Object, cod As String
    cod = "Private Sub Workbook_Open()" & vbCrLf & _
          "    Me.ShowPivotTableFieldList = False" & vbCrLf & _
          "    modCasaOrganizada.AtualizarPainel" & vbCrLf & _
          "    modCasaOrganizada.IrDashboard" & vbCrLf & _
          "End Sub" & vbCrLf & vbCrLf & _
          "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)" & vbCrLf & _
          "    Select Case Sh.Name" & vbCrLf & _
          "        Case """ & SH_LANC & """, """ & SH_CAT & """" & vbCrLf & _
          "            modCasaOrganizada.AgendarAtualizacao" & vbCrLf & _
          "        Case """ & SH_CON & """, """ & SH_MET & """, """ & SH_REND & """, """ & SH_ORC & """" & vbCrLf & _
          "            modCasaOrganizada.AtualizarVisual" & vbCrLf & _
          "    End Select" & vbCrLf & _
          "End Sub" & vbCrLf & vbCrLf & _
          "Private Sub Workbook_SheetPivotTableUpdate(ByVal Sh As Object, ByVal Target As PivotTable)" & vbCrLf & _
          "    If Not modCasaOrganizada.gOcupado Then modCasaOrganizada.SincronizarFiltros" & vbCrLf & _
          "End Sub"
    On Error GoTo semAcesso
    Set cm = ThisWorkbook.VBProject.VBComponents(ThisWorkbook.CodeName).CodeModule
    If cm.CountOfLines > 0 Then cm.DeleteLines 1, cm.CountOfLines
    cm.AddFromString cod
    Exit Sub
semAcesso:
End Sub

'===============================================================================
' AUXILIARES
'===============================================================================
' Acrescenta n linhas ao fim da tabela de uma vez e devolve o índice da primeira nova linha.
' Reaproveita a linha vazia que o Excel mantém em tabelas recém-criadas.
Private Function AcrescentarLinhas(lo As ListObject, n As Long) As Long
    Dim ini As Long, vazia As Boolean
    If lo.ListRows.Count = 1 Then vazia = (Application.WorksheetFunction.CountA(lo.ListRows(1).Range) = 0) Or _
        (Len(CStr(lo.ListRows(1).Range.Cells(1, 1).Value)) = 0 And Len(CStr(lo.ListRows(1).Range.Cells(1, 2).Value)) = 0)
    If lo.ListRows.Count = 0 Or vazia Then
        ini = 1
        If lo.ListRows.Count = 0 Then lo.ListRows.Add
        If n > 1 Then lo.Resize lo.Range.Resize(lo.Range.Rows.Count + n - 1)
    Else
        ini = lo.ListRows.Count + 1
        lo.Resize lo.Range.Resize(lo.Range.Rows.Count + n)
    End If
    AcrescentarLinhas = ini
End Function

Private Function DataPadrao() As Date
    Dim ano As Long
    ano = ThisWorkbook.Names("Ano").RefersToRange.Value
    If Year(Date) = ano Then DataPadrao = Date Else DataPadrao = DateSerial(ano, 12, 31)
End Function

Private Function MesPadrao() As Long
    MesPadrao = Month(DataPadrao())
End Function

Private Function PegarOuCriarAba(nome As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nome)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nome
    End If
    Set PegarOuCriarAba = ws
End Function

Private Sub LimparFormas(ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
End Sub

Private Sub ApagarAba(nome As String)
    Application.DisplayAlerts = False
    On Error Resume Next
    ThisWorkbook.Worksheets(nome).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Private Sub ApagarSegmentacoes()
    Dim i As Long
    On Error Resume Next
    For i = ThisWorkbook.SlicerCaches.Count To 1 Step -1
        ThisWorkbook.SlicerCaches(i).Delete
    Next i
End Sub

' Valores de um intervalo sempre como matriz 2D (um intervalo de 1 célula devolve valor simples no Excel).
Private Function Col2D(r As Range) As Variant
    Dim a() As Variant
    If r.Cells.Count = 1 Then
        ReDim a(1 To 1, 1 To 1): a(1, 1) = r.Value: Col2D = a
    Else
        Col2D = r.Value
    End If
End Function

Private Function Num(v As Variant) As Double
    If IsNumeric(v) And Len(CStr(v)) > 0 Then Num = CDbl(v)
End Function

Private Sub NomeFaixa(nome As String, alvo As Range)
    On Error Resume Next
    ThisWorkbook.Names(nome).Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:=nome, RefersTo:="='" & alvo.Worksheet.Name & "'!" & alvo.Address
End Sub

Private Function ColunaExiste(lo As ListObject, nome As String) As Boolean
    Dim c As ListColumn
    For Each c In lo.ListColumns
        If c.Name = nome Then ColunaExiste = True: Exit Function
    Next c
End Function

Private Function CodigoDaLista(txt As String) As Long
    Dim p As Long, n As Variant, lo As ListObject
    p = InStr(1, txt, " ")
    If p = 0 Then n = txt Else n = Left(txt, p - 1)
    If Not IsNumeric(n) Then Exit Function
    Set lo = ThisWorkbook.Worksheets(SH_CAT).ListObjects("tbCategorias")
    If IsError(Application.Match(CLng(n), lo.ListColumns("Código").DataBodyRange, 0)) Then Exit Function
    CodigoDaLista = CLng(n)
End Function

Private Function ExisteDuplicado(lo As ListObject, d As Date, cod As Long, v As Double) As Boolean
    If lo.ListRows.Count = 0 Then Exit Function
    ExisteDuplicado = Application.WorksheetFunction.CountIfs( _
        lo.ListColumns("Data").DataBodyRange, CDbl(d), _
        lo.ListColumns("Código").DataBodyRange, cod, _
        lo.ListColumns("Valor").DataBodyRange, v) > 0
End Function

Private Sub Aviso(msg As String, alvo As Range)
    MsgBox msg, vbExclamation, "Casa Organizada"
    On Error Resume Next
    alvo.Worksheet.Activate
    alvo.Select
End Sub

Private Sub Status(msg As String)
    On Error Resume Next
    ThisWorkbook.Names("msgStatus").RefersToRange.Value = Format(Now, "hh:mm") & " · " & msg
End Sub
