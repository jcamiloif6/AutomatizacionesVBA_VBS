dim fecha, valueFecha, filename1,filename2, filename3, filename4, ruta, valueRuta

fecha = "Digite la fecha exacta" + vbCrLf + vbCrLf +"Ejemplo: 01.02.2023"
ruta = "Digite la ruta a guardar los pantallazos"

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
session.findById("wnd[0]/mbar/menu[4]/menu[3]/menu[1]").select
session.findById("wnd[0]/usr/ctxtRS38R-QNUM").text = "1_OP_DETALLE"

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename1, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/btnP1").press
session.findById("wnd[0]/usr/ctxtSP$00004-LOW").text = "F"
session.findById("wnd[0]/usr/ctxtSP$00004-HIGH").text = "K"
session.findById("wnd[0]/usr/txt%_SP$00006_%_APP_%-TEXT").setFocus
session.findById("wnd[0]/usr/txt%_SP$00006_%_APP_%-TEXT").caretPosition = 20
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 5,"TEXT"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/txtSP$00006-LOW").text = "L"
session.findById("wnd[0]/usr/txt%_SP$00013_%_APP_%-TEXT").setFocus
session.findById("wnd[0]/usr/txt%_SP$00013_%_APP_%-TEXT").caretPosition = 10
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 1,"TEXT"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "1"
session.findById("wnd[1]/tbar[0]/btn[0]").press
valueFecha = InputBox(fecha)
session.findById("wnd[0]/usr/ctxtSP$00013-LOW").text = valueFecha
session.findById("wnd[0]/usr/txt%_SP$00016_%_APP_%-TEXT").setFocus
session.findById("wnd[0]/usr/txt%_SP$00016_%_APP_%-TEXT").caretPosition = 22
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 5,"TEXT"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/txtSP$00016-LOW").text = "L"
session.findById("wnd[0]/usr/txt%_SP$00019_%_APP_%-TEXT").setFocus
session.findById("wnd[0]/usr/txt%_SP$00019_%_APP_%-TEXT").caretPosition = 19
session.findById("wnd[0]").sendVKey 2
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").setCurrentCell 5,"TEXT"
session.findById("wnd[1]/usr/cntlOPTION_CONTAINER/shellcont/shell").selectedRows = "5"
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtSP$00019-LOW").text = "X"

'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename2, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/ctxtSP$00019-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSP$00019-LOW").caretPosition = 1
session.findById("wnd[0]/tbar[1]/btn[8]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").select
'para tomar el pantallazo
session.ActiveWindow().Hardcopy filename3, True 'Guardar la imagen en formato PNG

session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/tbar[0]/btn[0]").press
