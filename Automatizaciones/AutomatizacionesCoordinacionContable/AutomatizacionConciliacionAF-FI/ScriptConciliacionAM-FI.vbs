'30.04.2023
Dim Fecha, Ruta, valueFecha, valueMes, ValueRuta
dim filename1, filename2, filename3 ,filename4, filename5, filename6,filename7, filename8, filename9, filename10, filename11, filename12
dim filename13, filename14, filename15, filename16, filename17, filename18

Fecha = "Digite la fecha deseada para la transaccion S_ALR_87011963" + vbCrLf + vbCrLf +"Ejemplo: 01.02.2023"
Ruta = "Digite la ruta a guardar los archivos"


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

'01 AF
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "S_ALR_87011963"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtBUKRS-LOW").text = "IGEN"

valueFecha = InputBox(Fecha)
valueMes = Mid(valueFecha, 4, 2)
ValueRuta = InputBox(Ruta)
filename1 = valueRuta +  "\pantallazo01AF1.png" 'Ruta y nombre del archivo a guardar
filename2 = valueRuta +  "\pantallazo01AF2-1.png" 'Ruta y nombre del archivo a guardar
filename13 = valueRuta +  "\pantallazo01AF2-2.png" 'Ruta y nombre del archivo a guardar
filename3 = valueRuta +  "\pantallazo4LFI1.png" 'Ruta y nombre del archivo a guardar
filename4 = valueRuta +  "\pantallazo4LFI2-1.png" 'Ruta y nombre del archivo a guardar
filename14 = valueRuta +  "\pantallazo4LFI2-2.png" 'Ruta y nombre del archivo a guardar
filename5 = valueRuta +  "\pantallazo0LFI1.png" 'Ruta y nombre del archivo a guardar
filename6 = valueRuta +  "\pantallazo0LFI2-1.png" 'Ruta y nombre del archivo a guardar
filename15 = valueRuta +  "\pantallazo0LFI2-2.png" 'Ruta y nombre del archivo a guardar
filename7 = valueRuta +  "\pantallazo09AF1.png" 'Ruta y nombre del archivo a guardar
filename8 = valueRuta +  "\pantallazo09AF2-1.png" 'Ruta y nombre del archivo a guardar
filename16 = valueRuta +  "\pantallazo09AF2-2.png" 'Ruta y nombre del archivo a guardar
filename9 = valueRuta +  "\pantallazo40AF1.png" 'Ruta y nombre del archivo a guardar
filename10 = valueRuta +  "\pantallazo40AF2-1.png" 'Ruta y nombre del archivo a guardar
filename17 = valueRuta +  "\pantallazo40AF2-2.png" 'Ruta y nombre del archivo a guardar
filename11 = valueRuta +  "\pantallazo3LFI1.png" 'Ruta y nombre del archivo a guardar
filename12 = valueRuta +  "\pantallazo3LFI2-1.png" 'Ruta y nombre del archivo a guardar
filename18 = valueRuta +  "\pantallazo3LFI2-2.png" 'Ruta y nombre del archivo a guardar


session.findById("wnd[0]/usr/ctxtBERDATUM").text = valueFecha
session.findById("wnd[0]/usr/ctxtBEREICH1").text = "01"
session.findById("wnd[0]/usr/ctxtSRTVR").text = "ZPUC"

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename1, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/ctxtSRTVR").setFocus
session.findById("wnd[0]/usr/ctxtSRTVR").caretPosition = 4
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"WAERS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "WAERS"
session.findById("wnd[0]/tbar[1]/btn[25]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "S1",12
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "TEXT",19
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "BTR1",21
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "BTR2",20
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "BTR3",21
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "WAERS",6
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"WAERS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "S1"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "TEXT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "BTR1"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "BTR2"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "BTR3"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "WAERS"

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename2, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleRow = 10

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename13, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = ValueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "01AF.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

'4L FI
session.findById("wnd[0]/tbar[0]/okcd").text = "F.01"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtSD_KTOPL-LOW").text = "PUC"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/ctxtBILAVERS").text = "NIIF"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/ctxtBILAVERS").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/ctxtBILAVERS").caretPosition = 4
session.findById("wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL").select
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL-ILOW_I[1,0]").text = "1600000000"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL-ILOW_I[1,1]").text = "1910000000"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL-IHIGH_I[2,0]").text = "1699999999"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL-IHIGH_I[2,1]").text = "1999999999"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL-IHIGH_I[2,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL-IHIGH_I[2,1]").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = "IGEN"
session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").text = "4L"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtBILBJAHR").text = "2023"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtB-MONATE-LOW").text = "1"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtB-MONATE-HIGH").text = valueMes
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtBILVJAHR").text = "2022"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-LOW").text = "1"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").text = valueMes


session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").caretPosition = 2
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2").select
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2/ssub%_SUBSCREEN_TABBL1:RFBILA00:0002/ctxtBILAGKON").text = "3"

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename3, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2/ssub%_SUBSCREEN_TABBL1:RFBILA00:0002/ctxtBILAGKON").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2/ssub%_SUBSCREEN_TABBL1:RFBILA00:0002/ctxtBILAGKON").caretPosition = 1
session.findById("wnd[0]/tbar[1]/btn[8]").press

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename4, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr").verticalScrollbar.position = 111

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename14, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = ValueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "4LFI.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press


'0L FI
session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").text = "0L"
session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").caretPosition = 1
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2").select

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename5, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/tbar[1]/btn[8]").press

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename6, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr").verticalScrollbar.position = 112
'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename15, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = ValueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "0LFI.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press

'09 AF
session.findById("wnd[0]/tbar[0]/okcd").text = "S_ALR_87011963"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtBUKRS-LOW").text = "IGEN"
session.findById("wnd[0]/usr/ctxtBERDATUM").text = valueFecha
session.findById("wnd[0]/usr/ctxtBEREICH1").text = "09"
session.findById("wnd[0]/usr/ctxtSRTVR").text = "ZPUC"
session.findById("wnd[0]/usr/ctxtSRTVR").setFocus
session.findById("wnd[0]/usr/ctxtSRTVR").caretPosition = 4

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename7, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"WAERS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "WAERS"
session.findById("wnd[0]/tbar[1]/btn[25]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "S1",12
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "TEXT",19
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "BTR1",21
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "BTR2",19
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "BTR3",21
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "WAERS",6
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"WAERS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "S1"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "TEXT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "BTR1"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "BTR2"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "BTR3"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "WAERS"

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename8, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleRow = 10
'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename16, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = ValueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "09AF.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press

'40 AF
session.findById("wnd[0]/tbar[0]/okcd").text = "S_ALR_87011963"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtBUKRS-LOW").text = "IGEN"
session.findById("wnd[0]/usr/ctxtBERDATUM").text = valueFecha
session.findById("wnd[0]/usr/ctxtBEREICH1").text = "40"
session.findById("wnd[0]/usr/ctxtSRTVR").text = "ZPUC"
session.findById("wnd[0]/usr/ctxtSRTVR").setFocus
session.findById("wnd[0]/usr/ctxtSRTVR").caretPosition = 4

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename9, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setCurrentCell -1,"WAERS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "WAERS"
session.findById("wnd[0]/tbar[1]/btn[25]").press
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "S1",10
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "TEXT",10
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "BTR1",10
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "BTR2",10
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "BTR3",10
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "WAERS",10
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "S1",12
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "TEXT",19
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "BTR1",21
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "BTR2",20
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "BTR3",21
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").setColumnWidth "WAERS",6
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").currentCellColumn = "WAERS"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "S1"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "TEXT"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "BTR1"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "BTR2"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "BTR3"
session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").selectColumn "WAERS"

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename10, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell").firstVisibleRow = 10
'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename17, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = ValueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "40AF.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press
session.findById("wnd[0]/tbar[0]/btn[15]").press

'3L FI
session.findById("wnd[0]/tbar[0]/okcd").text = "F.01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ctxtSD_KTOPL-LOW").text = "PUC"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/ctxtBILAVERS").text = "NIIF"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/ctxtBILAVERS").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/ctxtBILAVERS").caretPosition = 4
session.findById("wnd[0]/usr/btn%_SD_SAKNR_%_APP_%-VALU_PUSH").press
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL").select
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL-ILOW_I[1,0]").text = "1600000000"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL-ILOW_I[1,1]").text = "1910000000"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL-IHIGH_I[2,0]").text = "1699999999"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL-IHIGH_I[2,1]").text = "1999999999"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL-IHIGH_I[2,1]").setFocus
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/ctxtRSCSEL-IHIGH_I[2,1]").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = "IGEN"
session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").text = "3L"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtBILBJAHR").text = "2023"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtB-MONATE-LOW").text = "1"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtB-MONATE-HIGH").text = valueMes
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtBILVJAHR").text = "2022"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-LOW").text = "1"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").text = valueMes
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").caretPosition = 2
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2").select
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2/ssub%_SUBSCREEN_TABBL1:RFBILA00:0002/ctxtBILAGKON").text = "3"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2/ssub%_SUBSCREEN_TABBL1:RFBILA00:0002/ctxtBILAGKON").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM2/ssub%_SUBSCREEN_TABBL1:RFBILA00:0002/ctxtBILAGKON").caretPosition = 1

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename11, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/tbar[1]/btn[8]").press

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename12, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr").verticalScrollbar.position = 112
'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename18, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = ValueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "3LFI.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
