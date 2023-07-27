Dim FechaPS, Ruta, valueFechaPS, valueMesFI, ValueRuta, valueAnioFI, valueMesCompararFI, mesCompararFI 
dim filename1, filename2, filename3, filename4, filename5, filename6, filename7, filename8

FechaPS = "Digite el mes y el año del reporte" + vbCrLf + vbCrLf +"Ejemplo: 005.2023"
Ruta = "Digite la ruta a guardar los archivos"
mesCompararFI = "Digite el numero del mes de comparacion"

If Not IsObject(application) Then
   Set SapGuiAuto  = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session    = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session,     "on"
   WScript.ConnectObject application, "on"
End If

'PS
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/ctxtCN_PROJN-LOW").text = "1"
session.findById("wnd[0]/usr/ctxtCN_PROJN-HIGH").text = "zzzzzzzzzzzzzzzzzzzzzzzz"
valueFechaPS = InputBox(FechaPS)
valueMesFI = Mid(valueFechaPS,2,2)
valueAnioFI = Mid(valueFechaPS,5,4)
ValueRuta = InputBox(Ruta)

filename1 = valueRuta +  "\pantallazo1PS.png" 'Ruta y nombre del archivo a guardar
filename2 = valueRuta +  "\pantallazo2PS.png" 'Ruta y nombre del archivo a guardar
filename3 = valueRuta +  "\pantallazo1-0L.png" 'Ruta y nombre del archivo a guardar
filename4 = valueRuta +  "\pantallazo2-0L.png" 'Ruta y nombre del archivo a guardar
filename5 = valueRuta +  "\pantallazo3-0L.png" 'Ruta y nombre del archivo a guardar
filename6 = valueRuta +  "\pantallazo1-3L.png" 'Ruta y nombre del archivo a guardar
filename7 = valueRuta +  "\pantallazo2-3L.png" 'Ruta y nombre del archivo a guardar
filename8 = valueRuta +  "\pantallazo3-3L.png" 'Ruta y nombre del archivo a guardar

session.findById("wnd[0]/usr/ctxtPAR_02").text = "001.1900"
session.findById("wnd[0]/usr/ctxtPAR_03").text = valueFechaPS

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename1, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/ctxtPAR_03").setFocus
session.findById("wnd[0]/usr/ctxtPAR_03").caretPosition = 8
session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell/shellcont[1]/shell").selectItem "        336","C          2"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell/shellcont[1]/shell").ensureVisibleHorizontalItem "        336","C          2"
session.findById("wnd[0]/usr/cntlCONTAINER/shellcont/shell/shellcont[1]/shell/shellcont[1]/shell/shellcont[1]/shell").topNode = "        313"
session.findById("wnd[0]/mbar/menu[2]/menu[0]").select
session.findById("wnd[0]/tbar[1]/btn[33]").press
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").currentCellRow = 8
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").firstVisibleRow = 0
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").selectedRows = "8"
session.findById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell").clickCurrentCell
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleRow = 4

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename2, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = ValueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "ps.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

'FI
session.findById("wnd[0]/tbar[0]/okcd").text = "F.01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "71265623"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 1
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").text = "0L"
valueMesCompararFI = InputBox(mesCompararFI)
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/ctxtBILAVERS").text = "NIIF"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtBILBJAHR").text = valueAnioFI
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtB-MONATE-LOW").text = "1"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtB-MONATE-HIGH").text = valueMesFI
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtBILVJAHR").text = valueAnioFI
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-LOW").text = "1"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").text = valueMesCompararFI

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename3, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").caretPosition = 2
session.findById("wnd[0]/tbar[1]/btn[8]").press

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename4, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr").verticalScrollbar.position = 1
session.findById("wnd[0]/usr").verticalScrollbar.position = 2
session.findById("wnd[0]/usr").verticalScrollbar.position = 3
session.findById("wnd[0]/usr").verticalScrollbar.position = 4
session.findById("wnd[0]/usr").verticalScrollbar.position = 5
session.findById("wnd[0]/usr").verticalScrollbar.position = 6
session.findById("wnd[0]/usr").verticalScrollbar.position = 7
session.findById("wnd[0]/usr").verticalScrollbar.position = 10

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename5, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = ValueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "0LFI.XLS"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").text = "3L"

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename6, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").caretPosition = 2
session.findById("wnd[0]/tbar[1]/btn[8]").press

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename7, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr").verticalScrollbar.position = 1
session.findById("wnd[0]/usr").verticalScrollbar.position = 2
session.findById("wnd[0]/usr").verticalScrollbar.position = 3
session.findById("wnd[0]/usr").verticalScrollbar.position = 4
session.findById("wnd[0]/usr").verticalScrollbar.position = 5
session.findById("wnd[0]/usr").verticalScrollbar.position = 6

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename8, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = ValueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "3LFI.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

