Attribute VB_Name = "criar_ped"
Sub Digitar_ped()

Dim Ref_Material_Excel_Plan As String
Dim objExcel
Dim quantidade As String
Dim objSheet, intRow, i
Dim lin_ped_ttl_excel As Long
Dim Ref_Material_ID_Tab_ME_SAP As String
Dim number As String
Dim mat_exc_dig As String
Dim Lanc_Actual_lin_ME_SAP As String
Dim re As Integer
Dim Ref_Lote_ID_Tab_ME_SAP As String
Dim Lanc_Lote_ME_SAP As String
Dim Lanc_qtd_ME_SAP As String
Dim Ref_Quant_pedida_ID_Tab_ME_SAP As String
Dim Cont_max_SAP_Lanc_item As Long
Dim page_actual_for_dig As Long
Dim Center_dest_tab As String
Dim Dep_dest_tab As String
Dim Lote As String
Dim Ref_From_Centro_Excel_Plan As String
Dim To_centro As String
Dim Lanc_CenterDest_ME_SAP As String
Dim Lanc_dep_ME_SAP As String


Set objExcel = GetObject(, "Excel.Application")

Set objSheet = objExcel.ActiveWorkbook.ActiveSheet

Call Conexao_SAP("ME21N")

'#Linhas

On Error Resume Next

'For i = 1 To objSheet.UsedRange.Rows.Count

 

 '   Ref_Material_Excel_Plan = Trim(CStr(objSheet.Cells(i, 1).Value))
  '  Lote = Trim(CStr(objSheet.Cells(i, 2).Value))
  '  quantidade = Trim(CStr(objSheet.Cells(i, 3).Value))
    Ref_From_Centro_Excel_Plan = Trim(CStr(objSheet.Cells(1, 4).Value))
    'To_centro = Trim(CStr(objSheet.Cells(i, 6).Value))


Session.findById("wnd[0]").maximize

'-----------_----__-____--______--____------___-----__-___--_----_-------_---_____----__---_---_---___--__--_--__--_-____----___--__-__-_----____-
'
'------deposito origem -------
If Ref_From_Centro_Excel_Plan = "2009" Then Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").Text = "2009"

If Ref_From_Centro_Excel_Plan = "2005" Then Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").Text = "2005"

If Ref_From_Centro_Excel_Plan = "2001" Then Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB0:SAPLMEGUI:0030/subSUB1:SAPLMEGUI:1105/ctxtMEPO_TOPLINE-SUPERFIELD").Text = "2001"


'--------------material----------
'Session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,0]").Text = Ref_Material_Excel_Plan
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
'Cont_max_SAP_Lanc_item = Session.findbyid("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB1:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1102/tabsHEADER_DETAIL/tabpTABHDT8/ssubTABSTRIPCONTROL2SUB:SAPLMEGUI:1221").verticalScrollbar.Maximum
 Cont_max_SAP_Lanc_item = Session.findById("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211").visiblerowcount

list_lins_max_ME_SAP = Cont_max_SAP_Lanc_item
lin_ped_ttl_excel = objSheet.UsedRange.Rows.Count
totalpag = lin_ped_ttl_excel / Cont_max_SAP_Lanc_item
'---------------contou as paginas e agora vai se preparar para preencher as colunas
ttlpag = objExcel.WorksheetFunction.RoundUp(totalpag, 0)
'###############revisado para entrar na tabela de digita��o do pedido
Ref_Material_ID_Tab_ME_SAP = Trim(CStr("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-EMATN[4,"))
'---------centro destino--------
Center_dest_tab = CStr("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-NAME1[7,")
'---------deposito destino------
Dep_dest_tab = CStr("wnd[0]/usr/subSUB0:SAPLMEGUI:0013/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-LGOBE[8,")


Ref_Quant_pedida_ID_Tab_ME_SAP = Trim(CStr("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/txtMEPO1211-MENGE[6,"))
Ref_Lote_ID_Tab_ME_SAP = Trim(CStr("wnd[0]/usr/subSUB0:SAPLMEGUI:0016/subSUB2:SAPLMEVIEWS:1100/subSUB2:SAPLMEVIEWS:1200/subSUB1:SAPLMEGUI:1211/tblSAPLMEGUITC_1211/ctxtMEPO1211-CHARG[5,"))
Ref_End_ID_Tab_ME_SAP = Trim(CStr("]"))
'------agora vai preencher os valores das linhas
For page_actual_for_dig = 1 To ttlpag

For i = posicao To objSheet.UsedRange.Rows.Count -1

Lanc_Actual_lin_ME_SAP = i
Lanc_qtd_ME_SAP = Ref_Quant_pedida_ID_Tab_ME_SAP + Lanc_Actual_lin_ME_SAP + Ref_End_ID_Tab_ME_SAP
Lanc_Lote_ME_SAP = Ref_Lote_ID_Tab_ME_SAP + Lanc_Actual_lin_ME_SAP + Ref_End_ID_Tab_ME_SAP
mat_exc_dig = Ref_Material_ID_Tab_ME_SAP + Lanc_Actual_lin_ME_SAP + Ref_End_ID_Tab_ME_SAP
Lanc_CenterDest_ME_SAP = Center_dest_tab + Lanc_Actual_lin_ME_SAP + Ref_End_ID_Tab_ME_SAP
Lanc_dep_ME_SAP = Dep_dest_tab + Lanc_Actual_lin_ME_SAP + Ref_End_ID_Tab_ME_SAP
i = i + 1
 Ref_Material_Excel_Plan = Trim(CStr(objSheet.Cells(i, 1).Value))
 

    If CStr(Ref_Material_Excel_Plan) = "" Then
    Exit For
    Else
    
    
 '   Ref_Material_Excel_Plan = Trim(CStr(objSheet.Cells(i, 1).Value))
  '  Lote = Trim(CStr(objSheet.Cells(i, 2).Value))
  '  quantidade = Trim(CStr(objSheet.Cells(i, 3).Value))
'    Ref_From_Centro_Excel_Plan = Trim(CStr(objSheet.Cells(i, 4).Value))
    To_centro = Trim(CStr(objSheet.Cells(i, 6).Value))
    
    
    
    
Session.findById(Lanc_dep_ME_SAP).Text = Trim(CStr(objSheet.Cells(1, 5).Value))
Session.findById(Lanc_CenterDest_ME_SAP).Text = To_centro
Session.findById(Lanc_qtd_ME_SAP).Text = Trim(CStr(objSheet.Cells(i, 3).Value))
Session.findById(Lanc_Lote_ME_SAP).Text = Trim(CStr(objSheet.Cells(i, 2).Value))
Session.findById(mat_exc_dig).Text = Trim(CStr(objSheet.Cells(i, 1).Value))

    i = i - 1

End If

Next
If Session.findById(CStr(Ref_Material_Excel_Plan)).Text = "" Then Exit For
'----------mudar a pagina------------------
posicao = list_lins_max_ME_SAP * page_actual_for_dig
Session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER").verticalScrollbar.Position = posicao
 
'session.findById("wnd[0]/usr/tabsTAXI_TABSTRIP_OVERVIEW/tabpT\01/ssubSUBSCREEN_BODY:SAPMV50A:1102/tblSAPMV50ATC_LIPS_OVER").verticalScrollbar.Position = 29


Next
'------------------------copiar primeira quantidade



'Session.findById("wnd[0]/tbar[0]/btn[11]").press
'Session.findById("wnd[0]/tbar[0]/btn[3]").press
'session.findById("wnd[0]/tbar[0]/btn[3]").press
'Next
End Sub
