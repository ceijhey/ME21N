Attribute VB_Name = "criar_ped"
Sub Digitar_ped()

Dim MATERIAL As String
Dim objExcel
Dim quantidade As String
Dim objSheet, intRow, i
Dim linhasnaremessa As Long
Dim itvl As String
Dim number As String
Dim finer As String
Dim bome As String
Dim re As Integer
Dim qtrem As String
Dim qtreme As String
Dim litem As String
Dim klitem As String
Dim Psan As String
Dim Dayhux As String
Dim maximalinha As Long
Dim PagAtual As Long
Dim Patos As String
Dim Peixes As String
Dim Lote As String
Dim From_Centro As String
Dim To_centro As String
Dim karleo As String
Dim truwe As String


Set objExcel = GetObject(, "Excel.Application")

Set objSheet = objExcel.ActiveWorkbook.ActiveSheet

Call Conexao_SAP("ME21N")

'#Linhas

On Error Resume Next

'For i = 1 To objSheet.UsedRange.Rows.Count

 

 '   MATERIAL = Trim(CStr(objSheet.Cells(i, 1).Value))
  '  Lote = Trim(CStr(objSheet.Cells(i, 2).Value))
  '  quantidade = Trim(CStr(objSheet.Cells(i, 3).Value))
    From_Centro = Trim(CStr(objSheet.Cells(1, 4).Value))
    'To_centro = Trim(CStr(objSheet.Cells(i, 6).Value))


Session.findById("wnd[0]").maximize

'-----------_----__-____--______--____------___-----__-___--_----_-------_---_____----__---_---_---___--__--_--__--_-____----___--__-__-_----____-
'
'------deposito origem -------
If From_Centro = "2009" Then Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").Text = "2009"

If From_Centro = "2005" Then Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").Text = "2005"

If From_Centro = "2001" Then Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").Text = "2001"


'--------------material----------
'Session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]").Text = MATERIAL
'--------lote-----
'Session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-CHARG[5,0]").Text = Lote
'-------quantidade----------
'Session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,0]").Text = quantidade

'Organizacao
Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221/ctxtMEPO1222-EKORG").Text = "2009"
'-----------_----__-____--______--____------___-----__-___--_----_-------_---_____----__---_---_---___--__--_--__--_-____----___--__-__-_----____-
'FECHAR ITEM
Session.findById("wnd[0]").sendVKey 31
'FECHAR CABECALHO
Session.findById("wnd[0]").sendVKey 29
'abrir sintese
Session.findById("wnd[0]").sendVKey 27

'digitar valor da remessa

'------------------------copiar primeira quantidade
'maximalinha = Session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221").verticalScrollbar.Maximum
 maximalinha = Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").visiblerowcount

pagina = maximalinha
linhasnaremessa = objSheet.UsedRange.Rows.Count
totalpag = linhasnaremessa / maximalinha
'---------------contou as paginas e agora vai se preparar para preencher as colunas
ttlpag = Round(totalpag, 1)
If ttlpag < 1 Then
If ttlpag > 0 Then ttlpag = 1
End If
'###############revisado para entrar na tabela de digitação do pedido
itvl = Trim(CStr("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,"))
'---------centro destino--------
Patos = CStr("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[7,")
'---------deposito destino------
Peixes = CStr("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-LGOBE[8,")


Dayhux = Trim(CStr("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,"))
litem = Trim(CStr("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-CHARG[5,"))
pan = Trim(CStr("]"))
'------agora vai preencher os valores das linhas
For PagAtual = 1 To ttlpag

For i = posicao To -1

bome = i
Psan = Dayhux + bome + pan
klitem = litem + bome + pan
finer = itvl + bome + pan
karleo = Patos + bome + pan
truwe = Peixes + bome + pan
i = i + 1
 MATERIAL = Trim(CStr(objSheet.Cells(i, 1).Value))
 

    If CStr(MATERIAL) = "" Then
    Exit For
    Else
    
    
 '   MATERIAL = Trim(CStr(objSheet.Cells(i, 1).Value))
  '  Lote = Trim(CStr(objSheet.Cells(i, 2).Value))
  '  quantidade = Trim(CStr(objSheet.Cells(i, 3).Value))
'    From_Centro = Trim(CStr(objSheet.Cells(i, 4).Value))
    To_centro = Trim(CStr(objSheet.Cells(i, 6).Value))
    
    
    
    
Session.findById(truwe).Text = Trim(CStr(objSheet.Cells(1, 5).Value))
Session.findById(karleo).Text = To_centro
Session.findById(Psan).Text = Trim(CStr(objSheet.Cells(i, 3).Value))
Session.findById(klitem).Text = Trim(CStr(objSheet.Cells(i, 2).Value))
Session.findById(finer).Text = Trim(CStr(objSheet.Cells(i, 1).Value))

    i = i - 1

End If

Next
If Session.findById(CStr(MATERIAL)).Text = "" Then Exit For
'----------mudar a pagina------------------
posicao = pagina * PagAtual
Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER").verticalScrollbar.Position = posicao
 
'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER").verticalScrollbar.Position = 29


Next
'------------------------copiar primeira quantidade



'Session.findById("wnd[0]/tbar[0]/btn[11]").press
'Session.findById("wnd[0]/tbar[0]/btn[3]").press
'session.findById("wnd[0]/tbar[0]/btn[3]").press
'Next
End Sub
