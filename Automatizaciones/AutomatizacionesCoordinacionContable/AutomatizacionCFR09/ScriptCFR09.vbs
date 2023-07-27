dim filename1, filename2, filename3, fechaActual, fechaAnterior, ruta, valueRuta,valueAnioActual, valueMesActual, valueAnioAnterior, valueMesAnterior
dim valueFechaActual, valueFechaAnterior

fechaActual = "Digite el mes y año del trimestre actual con el siguiente formato: MM.AAAA" + vbCrLf + vbCrLf +"Ejemplo: 02.2023"
fechaAnterior = "Digite el mes y año del trimestre anterior con el siguiente formato: MM.AAAA" + vbCrLf + vbCrLf +"Ejemplo: 02.2023"
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
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").text = "f.01"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/tbar[1]/btn[17]").press
session.findById("wnd[1]/usr/txtENAME-LOW").text = "1033650783"
session.findById("wnd[1]/usr/txtENAME-LOW").setFocus
session.findById("wnd[1]/usr/txtENAME-LOW").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[8]").press

valueRuta = InputBox(ruta)
filename1 = valueRuta +  "\pantallazo1.png" 'Ruta y nombre del archivo a guardar
filename2 = valueRuta +  "\pantallazo2.png" 'Ruta y nombre del archivo a guardar
filename3 = valueRuta +  "\pantallazo3.png" 'Ruta y nombre del archivo a guardar

valueFechaActual = InputBox(fechaActual)
valueFechaAnterior = InputBox(fechaAnterior)
valueAnioActual = Mid(valueFechaActual, 4,4)
valueMesActual = Mid(valueFechaActual, 1,2)
valueAnioAnterior = Mid(valueFechaAnterior, 4,4)
valueMesAnterior = Mid(valueFechaAnterior, 1,2)

session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtBILBJAHR").text = valueAnioActual
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtB-MONATE-LOW").text = "1"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtB-MONATE-HIGH").text = valueMesActual
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtBILVJAHR").text = valueAnioAnterior
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-LOW").text = "1"
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").text = valueMesAnterior

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename1, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").setFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_TABBL1/tabpUCOM1/ssub%_SUBSCREEN_TABBL1:RFBILA00:0001/txtV-MONATE-HIGH").caretPosition = 2
session.findById("wnd[0]/tbar[1]/btn[8]").press

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename2, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = valueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "CFR09.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/mbar/menu[0]/menu[1]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = ValueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "CFR09.htm"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 5
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]").maximize
session.findById("wnd[0]/usr").verticalScrollbar.position = 1179

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename3, True 'Guardar la imagen en formato PNG
