dim filename1, filename2, filename3, filename4, fechaInicial, fechaFinal, ruta, valueRuta,valueFechaInicial, valueFechaFinal, valueAnio

fechaInicial = "Digite la fecha inicial para sacar el reporte" + vbCrLf + vbCrLf +"Ejemplo: 01.02.2023"
fechaFinal = "Digite la fecha final para sacar el reporte" + vbCrLf + vbCrLf +"Ejemplo: 01.02.2023"
ruta = "Digite la ruta a guardar los archivos"

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

valueRuta = InputBox(ruta)

filename1 = valueRuta +  "\pantallazo1.png" 'Ruta y nombre del archivo a guardar
filename2 = valueRuta +  "\pantallazo2.png" 'Ruta y nombre del archivo a guardar
filename3 = valueRuta +  "\pantallazo3.png" 'Ruta y nombre del archivo a guardar
filename4 = valueRuta +  "\pantallazo4.png" 'Ruta y nombre del archivo a guardar

session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "S_ALR_87012347"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "43798063"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 8
session.findById("wnd[1]/tbar[0]/btn[8]").press
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").currentCellRow = 4
session.findById("wnd[1]/usr/cntlALV_CONTAINER_1/shellcont/shell").selectedRows = "4"
session.findById("wnd[1]/tbar[0]/btn[2]").press
session.findById("wnd[0]/tbar[1]/btn[16]").press
valueFechaInicial = InputBox(fechaInicial)
valueFechaFinal = InputBox(fechaFinal)
valueAnio = Mid(valueFechaInicial, 7, 4)
session.findById("wnd[0]/usr/txtBR_GJAHR-LOW").text = valueAnio
session.findById("wnd[0]/usr/ctxtBR_BUDAT-LOW").text = valueFechaInicial
session.findById("wnd[0]/usr/ctxtBR_BUDAT-HIGH").text = valueFechaFinal

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename1, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/ctxtBR_BUDAT-HIGH").setFocus
session.findById("wnd[0]/usr/ctxtBR_BUDAT-HIGH").caretPosition = 10
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN005_%_APP_%-VALU_PUSH").press

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename2, True 'Guardar la imagen en formato PNG

session.findById("wnd[1]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[0]/mbar/menu[3]/menu[9]").select

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename3, True 'Guardar la imagen en formato PNG

session.findById("wnd[1]/tbar[0]/btn[0]").press

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename4, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = valueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "CFC12.htm"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[0]").press



session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = valueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "CFC12.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[0]").press
