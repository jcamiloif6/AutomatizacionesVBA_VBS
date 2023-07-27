dim year, yearValue, mes, valueMes, ruta, ValueRuta, filename1, filename3

mes = "Digite el número del mes del reporte"
year = "Digite el año del reporte"
ruta = "Ubicación en donde se guardarán los archivos"

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
session.findById("wnd[0]/tbar[0]/okcd").text = "f.08"
session.findById("wnd[0]/tbar[0]/btn[0]").press
session.findById("wnd[0]/usr/ctxtSD_KTOPL-LOW").text = "PUC"
session.findById("wnd[0]/usr/ctxtSD_BUKRS-LOW").text = "IGEN"
yearValue = InputBox(year)
valueMes = InputBox(mes)
ValueRuta = InputBox(ruta)

filename1 = valueRuta +  "\pantallazo1-3L.png" 'Ruta y nombre del archivo a guardar
filename3 = valueRuta +  "\pantallazo1-0L.png" 'Ruta y nombre del archivo a guardar

session.findById("wnd[0]/usr/txtSD_GJAHR-LOW").text = yearValue
session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").text = "3L"
session.findById("wnd[0]/usr/txtB_MONATE-LOW").text = "1"
session.findById("wnd[0]/usr/txtB_MONATE-HIGH").text = valueMes
session.findById("wnd[0]/usr/txtZSB_POS2").text = "4"
session.ActiveWindow().Hardcopy filename1, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/txtZSB_POS2").setFocus
session.findById("wnd[0]/usr/txtZSB_POS2").caretPosition = 1
session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = ValueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "3L.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = valueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "3L.htm"
session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 54
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/tbar[0]/btn[3]").press
session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").text = "0L"
session.ActiveWindow().Hardcopy filename3, True 'Guardar la imagen en formato PNG

session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").setFocus
session.findById("wnd[0]/usr/ctxtSD_RLDNR-LOW").caretPosition = 2
session.findById("wnd[0]/tbar[1]/btn[8]").press

session.findById("wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = ValueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "0L.XLS"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 2
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/mbar/menu[3]/menu[5]/menu[2]/menu[2]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]").select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[3,0]").setFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").text = valueRuta
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = "0L.htm"
session.findById("wnd[1]/usr/ctxtDY_PATH").setFocus
session.findById("wnd[1]/usr/ctxtDY_PATH").caretPosition = 54
session.findById("wnd[1]/tbar[0]/btn[0]").press

session.findById("wnd[0]/tbar[0]/btn[12]").press
session.findById("wnd[0]/tbar[0]/btn[12]").press
