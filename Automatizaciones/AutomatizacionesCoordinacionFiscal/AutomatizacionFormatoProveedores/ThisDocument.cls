VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisDocument"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'Dim fr As FormField
Dim fr, fa, ica, icaMunicipio1, icaMunicipio2 As ContentControl
Dim fr2, fa2, ica2, municipio12, municipio22, icaMunicipio12, icaMunicipio22 As ContentControlListEntry
Dim botonISLV, botonNRDI, botonGC, boton22, botonOtros As Boolean
Dim botonBien, botonServicio As Boolean
Dim botonCR, botonRTE, botonNCDR, botonAutorretenedor, botonRST, botonDeclarante, botonND As Boolean
Dim contentFR, contentFA, contentICA, contentMunicipio1, contentICA1, contentICA2 As Integer
Dim contentMunicipio2 As Integer

Sub ListaDesplegable2()
    
contentICA1 = 1
contentICA2 = 6
contentFR = 3
contentFA = 4
contentICA = 5
Set fr = ActiveDocument.ContentControls(contentFR)
Set fa = ActiveDocument.ContentControls(contentFA)
Set ica = ActiveDocument.ContentControls(contentICA)
Set icaMunicipio1 = ActiveDocument.ContentControls(contentICA1)
Set icaMunicipio2 = ActiveDocument.ContentControls(contentICA2)
    
With fr.DropdownListEntries
    If Not IsNull(ListBox1m) Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add("NA")
        fr2.Select
    ElseIf ListBoxRenta.Value = "04 Régimen Tributario Especial" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add("EX")
        fr2.Select
   ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Honorarios" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(3)
        fr2.Select
   ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Comisiones  en General (Jurídicas)" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(5)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Servicios De Transporte De Carga" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(8)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Servicios Temporales (Personal)" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(9)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Transporte Aereo de Personas" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(10)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Servicios de Aseo y Vigilancia" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(11)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Servicios Integrales de Salud IPS" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(12)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Contratos De Construcción" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(13)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Servicios Generales - Declarantes" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(14)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Compra de Combustibles" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(17)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Compras y Suministros - Declarantes" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(26)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Intereses o Rendimientos Financieros" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(18)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Adquisición Bienes Raíces" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(21)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Adquisición de Vehículos" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(22)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Otros Ingresos Tributarios - Declarantes" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(24)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Arrendamientos Bienes Muebles" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(28)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Arrendamientos Bienes Raíces" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(42)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Servicios Transporte De Pasajeros" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(43)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Restaurante, Hotel y Hospedaje" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(44)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Indemnizaciones" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(33)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Software Bienes y Servicios" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(34)
        fr2.Select
    ElseIf ListBoxRenta.Value = "05 Contribuyente Renta" And ListBoxRentaDescripcion.Value = "Servicios Públicos" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(35)
        fr2.Select
    ElseIf ListBoxRenta.Value = "15 Autorretenedor" Or ListBoxRenta.Value = "47 Régimen Simple Tributación" _
        Or ListBoxRenta.Value = "06 No Contribuyente de Renta" Or ListBoxRenta.Value = "04 Régimen Tributario Especial" _
        Or ListBoxRenta.Value = "No Declarante" Then
        .Clear
        Set fr2 = fr.DropdownListEntries.Add("EX")
        fr2.Select
    Else
        .Clear
        Set fr2 = fr.DropdownListEntries.Add(" ")
        fr2.Select
    End If
End With

With fa.DropdownListEntries
    If Not IsNull(ListBox1m) Then
        .Clear
        Set fa2 = fa.DropdownListEntries.Add("NA")
        fa2.Select
    ElseIf ListBoxIVA.Value = "13 Gran Contribuyente" Or ListBoxIVA.Value = "23 Agentes de Retención en el IVA" Then
        .Clear
        Set fa2 = fa.DropdownListEntries.Add("EX")
        fa2.Select
    ElseIf ListBoxIVA.Value = "48 Impuesto sobre las ventas" And ListBoxIVADescripcion.Value = "Servicios Técnicos" Then
        .Clear
        Set fa2 = fa.DropdownListEntries.Add(77)
        fa2.Select
    ElseIf ListBoxIVA.Value = "48 Impuesto sobre las ventas" And ListBoxIVADescripcion.Value = "Honorarios" Then
        .Clear
        Set fa2 = fa.DropdownListEntries.Add(83)
        fa2.Select
    ElseIf ListBoxIVA.Value = "48 Impuesto sobre las ventas" And ListBoxIVADescripcion.Value = "Comisiones" Then
        .Clear
        Set fa2 = fa.DropdownListEntries.Add(73)
        fa2.Select
    ElseIf ListBoxIVA.Value = "48 Impuesto sobre las ventas" And ListBoxIVADescripcion.Value = "Bienes Gravados 19%" Then
        .Clear
        Set fa2 = fa.DropdownListEntries.Add(76)
        fa2.Select
    ElseIf ListBoxIVA.Value = "48 Impuesto sobre las ventas" And ListBoxIVADescripcion.Value = "Servicios 19%" Then
        .Clear
        Set fa2 = fa.DropdownListEntries.Add(77)
        fa2.Select
    ElseIf ListBoxIVA.Value = "48 Impuesto sobre las ventas" And ListBoxIVADescripcion.Value = "Vehículos" Then
        .Clear
        Set fa2 = fa.DropdownListEntries.Add(68)
        fa2.Select
    ElseIf ListBoxIVA.Value = "48 Impuesto sobre las ventas" And ListBoxIVADescripcion.Value = "Telefonía Móvil" Then
        .Clear
        Set fa2 = fa.DropdownListEntries.Add(72)
        fa2.Select
    ElseIf ListBoxIVA.Value = "48 Impuesto sobre las ventas" And ListBoxIVADescripcion.Value = "Bienes Gravados 5%" Then
        .Clear
        Set fa2 = fa.DropdownListEntries.Add(88)
        fa2.Select
    ElseIf ListBoxIVA.Value = "48 Impuesto sobre las ventas" And ListBoxIVADescripcion.Value = "Servicios 5%" Then
        .Clear
        Set fa2 = fa.DropdownListEntries.Add(89)
        fa2.Select
    ElseIf ListBoxIVA.Value = "No Aplica" Then
        .Clear
        Set fa2 = fa.DropdownListEntries.Add("EX")
        fa2.Select
    ElseIf ListBoxIVA.Value = "49 No responsable de IVA" Then
        .Clear
        Set fa2 = fa.DropdownListEntries.Add("EX")
        fa2.Select
    Else
        .Clear
        Set fa2 = fa.DropdownListEntries.Add(" ")
        fa2.Select
    End If
End With


With ica.DropdownListEntries
    If Not IsNull(ListBox1m) Then
        .Clear
        Set ica2 = ica.DropdownListEntries.Add("NA")
        ica2.Select
    ElseIf ListBoxRenta.Value = "47 Régimen Simple Tributación" Then
        .Clear
        Set ica2 = ica.DropdownListEntries.Add("EX")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Bogotá" And ListBoxIVA.Value = "13 Gran Contribuyente" Then
        .Clear
        Set ica2 = .Add("01-EX")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Cali" And ListBoxICA1.Value = "Sin Reteica" Then
        .Clear
        Set ica2 = .Add("02-EX")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Barranquilla" And ListBoxIVA.Value = "13 Gran Contribuyente" Then
        .Clear
        Set ica2 = .Add("03-EX")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "San Rafael" And (ListBoxIVA.Value = "13 Gran Contribuyente" _
        Or ListBoxICA1.Value = "Sin Reteica") Then
        .Clear
        Set ica2 = .Add("07-EX")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Giron" And (ListBoxIVA.Value = "13 Gran Contribuyente" _
        Or ListBoxICA1.Value = "Sin Reteica") Then
        .Clear
        Set ica2 = .Add("13-EX")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Bucaramanga" And ListBoxIVA.Value = "13 Gran Contribuyente" Then
        .Clear
        Set ica2 = .Add("14-EX")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Lebrija" And ListBoxICA1.Value = "Sin Reteica" Then
        .Clear
        Set ica2 = .Add("15-EX")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Betulia" And ListBoxICA1.Value = "Sin Reteica" Then
        .Clear
        Set ica2 = .Add("16-EX")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Medellín" And (ListBoxIVA.Value = "13 Gran Contribuyente" _
        Or ListBoxICA1.Value = "Sin Reteica") Then
        .Clear
        Set ica2 = .Add("18-EX")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Norcasia" And ListBoxIVA.Value = "13 Gran Contribuyente" Then
        .Clear
        Set ica2 = .Add("20-EX")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Barrancabermeja" And (ListBoxIVA.Value = "13 Gran Contribuyente" _
        Or ListBoxRenta.Value = "15 Autorretenedor") Then
        .Clear
        Set ica2 = .Add("17-EX")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Bogotá" And (ListBoxICA1.Value = "Servicios de Restaurante" _
        Or ListBoxICA1.Value = "Servicios de Hotel" Or ListBoxICA1.Value = "Servicios de Vigilancia") Then
        .Clear
        Set ica2 = .Add("01-02")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Bogotá" And ListBoxICA1.Value = "Compras y Otras Actividades Comerciales" Then
        .Clear
        Set ica2 = .Add("01-03")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Bogotá" And (ListBoxICA1.Value = "Transporte de pasajeros" _
     Or ListBoxICA1.Value = "Edición de periódicos" Or ListBoxICA1.Value = "Edición de revistas y/o otras publicaciones") Then
        .Clear
        Set ica2 = .Add("01-04")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Bogotá" And ListBoxICA1.Value = "Otros Servicios" Then
        .Clear
        Set ica2 = .Add("01-05")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Bogotá" And (ListBoxICA1.Value = "Servcicio de Cosultoría" _
        Or ListBoxICA1.Value = "Servcicio de Interventoría") Then
        .Clear
        Set ica2 = .Add("01-06")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Bogotá" And (ListBoxICA1.Value = "Seguros Generales de Vida" _
        Or ListBoxICA1.Value = "Servicios Financieros") Then
        .Clear
        Set ica2 = .Add("01-08")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Cali" And ListBoxICA1.Value = "Compras y Otras Actividades Comerciales" Then
        .Clear
        Set ica2 = .Add("02-02")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Cali" And (ListBoxICA1.Value = "Servicios de Hotel" _
        Or ListBoxICA1.Value = "Servicios de Restaurante y Cafeteria") Then
        .Clear
        Set ica2 = .Add("02-03")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Cali" And ListBoxICA1.Value = "Otros Servicios" Then
        .Clear
        Set ica2 = .Add("02-04")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Barranquilla" And ListBoxICA1.Value = "Servicios de Restaurante" Then
        .Clear
        Set ica2 = .Add("03-03")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Barranquilla" And ListBoxICA1.Value = "Servicios de aseo y limpieza" Then
        .Clear
        Set ica2 = .Add("03-06")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Barranquilla" And ListBoxICA1.Value = "Otros Servicios" Then
        .Clear
        Set ica2 = .Add("03-05")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Barranquilla" And ListBoxICA1.Value = "Otras Actividades Comerciales" Then
        .Clear
        Set ica2 = .Add("03-02")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Cimitarra" And (ListBoxICA1.Value = "Actividades de Construcción" _
        Or ListBoxICA1.Value = "Transporte" Or ListBoxICA1.Value = "Servicios Informáticos" _
        Or ListBoxICA1.Value = "Actividades Profesionales" Or ListBoxICA1.Value = "Actividades Científicas" _
        Or ListBoxICA1.Value = "Actividades Técnicas" Or ListBoxICA1.Value = "Actividades de Seguridad Privada" _
        Or ListBoxICA1.Value = "Actividades de Servicios Administrativos" Or ListBoxICA1.Value = "Actividades de Servicios Comerciales" _
        Or ListBoxICA1.Value = "Actividades de Servicios Industriales" Or ListBoxICA1.Value = "Otras Actividades") Then
        .Clear
        Set ica2 = .Add("04-06")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Cimitarra" And (ListBoxICA1.Value = "Actividades Finaniceras" _
        Or ListBoxICA1.Value = "Actividades de Alquiler y Arrendamiento") Then
        .Clear
        Set ica2 = .Add("04-07")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Cimitarra" And ListBoxICA1.Value = "Compra Combustible" Then
        .Clear
        Set ica2 = .Add("04-09")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Cimitarra" And ListBoxICA1.Value = "Otras Compras" Then
        .Clear
        Set ica2 = .Add("04-08")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "San Carlos" And ListBoxICA1.Value = "Servicios" Then
        .Clear
        Set ica2 = .Add("05-01")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "San Carlos" And ListBoxICA1.Value = "Compras" Then
        .Clear
        Set ica2 = .Add("05-02")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "San Rafael" And ListBoxICA1.Value = "Compras" Then
        .Clear
        Set ica2 = .Add("07-05")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "San Rafael" And ListBoxICA1.Value = "Servicios" Then
        .Clear
        Set ica2 = .Add("07-06")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Victoria" And ListBoxICA1.Value = "Servicios" Then
        .Clear
        Set ica2 = .Add("10-01")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Victoria" And ListBoxICA1.Value = "Compras" Then
        .Clear
        Set ica2 = .Add("10-02")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Giron" And ListBoxICA1.Value = "Compras" Then
        .Clear
        Set ica2 = .Add("13-07")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Giron" And ListBoxICA1.Value = "Servicios" Then
        .Clear
        Set ica2 = .Add("13-02")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Bucaramanga" And ListBoxICA1.Value = "Servicios" Then
        .Clear
        Set ica2 = .Add("14-04")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Bucaramanga" And ListBoxICA1.Value = "Compras" Then
        .Clear
        Set ica2 = .Add("14-03")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Lebrija" And ListBoxICA1.Value = "Compras" Then
        .Clear
        Set ica2 = .Add("15-01")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Lebrija" And ListBoxICA1.Value = "Servicios" Then
        .Clear
        Set ica2 = .Add("15-02")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Betulia" And ListBoxICA1.Value = "Servicios" Then
        .Clear
        Set ica2 = .Add("16-10")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Betulia" And ListBoxICA1.Value = "Compras" Then
        .Clear
        Set ica2 = .Add("16-09")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Betulia" And ListBoxICA1.Value = "Servicios de Construcción" Then
        .Clear
        Set ica2 = .Add("16-11")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Betulia" And ListBoxICA1.Value = "Restaurante y/o cafetería" Then
        .Clear
        Set ica2 = .Add("16-12")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Betulia" And ListBoxICA1.Value = "Administración de Instalaciones" Then
        .Clear
        Set ica2 = .Add("B1-13")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Betulia" And ListBoxICA1.Value = "Servicios de Seguridad Privada" Then
        .Clear
        Set ica2 = .Add("16-14")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Medellín" And ListBoxICA1.Value = "Compra de Bienes y Servicios" Then
        .Clear
        Set ica2 = .Add("18-01")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Zapatoca" And ListBoxICA1.Value = "Compra de Bienes y Servicios" Then
        .Clear
        Set ica2 = .Add("23-01")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Uribia" And ListBoxICA1.Value = "Compra de Bienes y Servicios" Then
        .Clear
        Set ica2 = .Add("25-01")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Salgar" And ListBoxICA1.Value = "Compra de Bienes y Servicios" Then
        .Clear
        Set ica2 = .Add("27-01")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "San Luis" And ListBoxICA1.Value = "Compra de Bienes y Servicios" Then
        .Clear
        Set ica2 = .Add("29-01")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Gómez Plata" And ListBoxICA1.Value = "Compra de Bienes y Servicios" Then
        .Clear
        Set ica2 = .Add("30-01")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Santa Rosa de Osos" And ListBoxICA1.Value = "Compra de Bienes y Servicios" Then
        .Clear
        Set ica2 = .Add("31-01")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "S.V Chucurí" And (ListBoxICA1.Value = "Servicios" _
        Or ListBoxICA1.Value = "Compras") Then
        .Clear
        Set ica2 = .Add("19-04")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Norcasia" And ListBoxICA1.Value = "Servicios" Then
        .Clear
        Set ica2 = .Add("20-04")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Norcasia" And ListBoxICA1.Value = "Compras" Then
        .Clear
        Set ica2 = .Add("20-03")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Chaparral" And (ListBoxICA1.Value = "Obras Civiles" _
        Or ListBoxICA1.Value = "Actividades Comerciales al Por Mayor" Or ListBoxICA1.Value = "Actividades Comerciales al Por Menor" _
        Or ListBoxICA1.Value = "Transporte" Or ListBoxICA1.Value = "Actividades de  Edición e Impresión y Publicaciones" _
        Or ListBoxICA1.Value = "Otras Industrias Manufactureras" Or ListBoxICA1.Value = "Desarrollo de Sistemas de Informatica" _
        Or ListBoxICA1.Value = "Comercio de Vehículos") Then
        .Clear
        Set ica2 = .Add("21-02")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Chaparral" And (ListBoxICA1.Value = "Publicidad" _
        Or ListBoxICA1.Value = " Alquiler Bien Mueble" Or ListBoxICA1.Value = "Actividades de agencias de viaje" _
        Or ListBoxICA1.Value = "Servicios de Seguridad e Investigación Privada" Or ListBoxICA1.Value = "Educación") Then
        .Clear
        Set ica2 = .Add("21-03")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Chaparral" And (ListBoxICA1.Value = "Compra de Combustible y/o Lubricantes" _
        Or ListBoxICA1.Value = "Alojamientos" Or ListBoxICA1.Value = "Alimentación y Actividades de Catering") Then
        .Clear
        Set ica2 = .Add("21-04")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Chaparral" And (ListBoxICA1.Value = "Mantenimiento y Reparación de Maquinaria y Equipos" _
        Or ListBoxICA1.Value = "Generación y Transmisión Eléctrica" Or ListBoxICA1.Value = "Distribución de Combustible" _
        Or ListBoxICA1.Value = "Actividades de Programación y Transmisión de Televisión" Or ListBoxICA1.Value = "Alimentación de Saneamiento Ambiental" _
        Or ListBoxICA1.Value = "Transporte por Tuberías") Then
        .Clear
        Set ica2 = .Add("21-05")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Granada" And ListBoxICA1.Value = "Compras" Then
        .Clear
        Set ica2 = .Add("22-01")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Granada" And ListBoxICA1.Value = "Servicios" Then
        .Clear
        Set ica2 = .Add("22-02")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Puerto Wilches" And ListBoxICA1.Value = "Actividades Industriales" Then
        .Clear
        Set ica2 = .Add("24-01")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Puerto Wilches" And ListBoxICA1.Value = "Actividades de Servicios y Comerciales" Then
        .Clear
        Set ica2 = .Add("24-02")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Amalfi" And ListBoxICA1.Value = "Compras" Then
        .Clear
        Set ica2 = .Add("26-01")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Amalfi" And ListBoxICA1.Value = "Servicios" Then
        .Clear
        Set ica2 = .Add("26-02")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Cocorná" And ListBoxICA1.Value = "Bienes y Servicios" Then
        .Clear
        Set ica2 = .Add("28-03")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Cocorná" And (ListBoxICA1.Value = "Mantenimiento y Reparación de Maquinaria y Equipo" _
        Or ListBoxICA1.Value = "Actividades de Limpieza" Or ListBoxICA1.Value = "Servicios Administrativos de Apoyo a las Empresas" _
        Or ListBoxICA1.Value = "Suministro de Vapor y Aire Acondicionado" Or ListBoxICA1.Value = "Actividades de Saneamiento Ambiental" _
        Or ListBoxICA1.Value = "Construcción de Edificio y Canteras" Or ListBoxICA1.Value = "Obras de Ingeniería Civil" _
        Or ListBoxICA1.Value = "Alquiler y Arrendamiento de Vehículos y Maquinaria") Then
        .Clear
        Set ica2 = .Add("28-04")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Barrancabermeja" And ListBoxICA1.Value = "Servicios" Then
        .Clear
        Set ica2 = .Add("17-07")
        ica2.Select
    ElseIf ListBoxMunicipio1.Value = "Barrancabermeja" And ListBoxICA1.Value = "Compras" Then
        .Clear
        Set ica2 = .Add("17-06")
        ica2.Select
    Else
        .Clear
        Set ica2 = ica.DropdownListEntries.Add(" ")
        ica2.Select
    End If
End With

With icaMunicipio2.DropdownListEntries
    If Not IsNull(ListBoxMunicipio2) Then
        If Not IsNull(ListBox1m) Then
            .Clear
            Set icaMunicipio22 = .Add("NA")
            icaMunicipio22.Select
        ElseIf ListBoxRenta.Value = "47 Régimen Simple Tributación" Then
            .Clear
            Set icaMunicipio22 = .Add("EX")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Bogotá" And ListBoxIVA.Value = "13 Gran Contribuyente" Then
            .Clear
            Set icaMunicipio22 = .Add("01-EX")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Cali" And ListBoxICA2.Value = "Sin Reteica" Then
            .Clear
            Set icaMunicipio22 = .Add("02-EX")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Barranquilla" And ListBoxIVA.Value = "13 Gran Contribuyente" Then
            .Clear
            Set icaMunicipio22 = .Add("03-EX")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "San Rafael" And (ListBoxIVA.Value = "13 Gran Contribuyente" _
            Or ListBoxICA2.Value = "Sin Reteica") Then
            .Clear
            Set icaMunicipio22 = .Add("07-EX")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Giron" And (ListBoxIVA.Value = "13 Gran Contribuyente" _
            Or ListBoxICA2.Value = "Sin Reteica") Then
            .Clear
            Set icaMunicipio22 = .Add("13-EX")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Bucaramanga" And ListBoxIVA.Value = "13 Gran Contribuyente" Then
            .Clear
            Set icaMunicipio22 = .Add("14-EX")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Lebrija" And ListBoxICA2.Value = "Sin Reteica" Then
            .Clear
            Set icaMunicipio22 = .Add("15-EX")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Betulia" And ListBoxICA2.Value = "Sin Reteica" Then
            .Clear
            Set icaMunicipio22 = .Add("16-EX")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Medellín" And (ListBoxIVA.Value = "13 Gran Contribuyente" _
            Or ListBoxICA2.Value = "Sin Reteica") Then
            .Clear
            Set icaMunicipio22 = .Add("18-EX")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Norcasia" And ListBoxIVA.Value = "13 Gran Contribuyente" Then
            .Clear
            Set icaMunicipio22 = .Add("20-EX")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Barrancabermeja" And (ListBoxIVA.Value = "13 Gran Contribuyente" _
            Or ListBoxRenta.Value = "15 Autorretenedor") Then
            .Clear
            Set icaMunicipio22 = .Add("17-EX")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Bogotá" And (ListBoxICA2.Value = "Servicios de Restaurante" _
            Or ListBoxICA2.Value = "Servicios de Hotel" Or ListBoxICA2.Value = "Servicios de Vigilancia") Then
            .Clear
            Set icaMunicipio22 = .Add("01-02")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Bogotá" And ListBoxICA2.Value = "Compras y Otras Actividades Comerciales" Then
            .Clear
            Set icaMunicipio22 = .Add("01-03")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Bogotá" And (ListBoxICA2.Value = "Transporte de pasajeros" _
            Or ListBoxICA2.Value = "Edición de periódicos" Or ListBoxICA2.Value = "Edición de revistas y/o otras publicaciones") Then
            .Clear
            Set icaMunicipio22 = .Add("01-04")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Bogotá" And ListBoxICA2.Value = "Otros Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("01-05")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Bogotá" And (ListBoxICA2.Value = "Servcicio de Cosultoría" _
            Or ListBoxICA2.Value = "Servcicio de Interventoría") Then
            .Clear
            Set icaMunicipio22 = .Add("01-06")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Bogotá" And (ListBoxICA2.Value = "Seguros Generales de Vida" _
            Or ListBoxICA2.Value = "Servicios Financieros") Then
            .Clear
            Set icaMunicipio22 = .Add("01-08")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Cali" And ListBoxICA2.Value = "Compras y Otras Actividades Comerciales" Then
            .Clear
            Set icaMunicipio22 = .Add("02-02")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Cali" And (ListBoxICA2.Value = "Servicios de Hotel" _
            Or ListBoxICA2.Value = "Servicios de Restaurante y Cafeteria") Then
            .Clear
            Set icaMunicipio22 = .Add("02-03")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Cali" And ListBoxICA2.Value = "Otros Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("02-04")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Barranquilla" And ListBoxICA2.Value = "Servicios de Restaurante" Then
            .Clear
            Set icaMunicipio22 = .Add("03-03")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Barranquilla" And ListBoxICA2.Value = "Servicios de aseo y limpieza" Then
            .Clear
            Set icaMunicipio22 = .Add("03-06")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Barranquilla" And ListBoxICA2.Value = "Otros Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("03-05")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Barranquilla" And ListBoxICA2.Value = "Otras Actividades Comerciales" Then
            .Clear
            Set icaMunicipio22 = .Add("03-02")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Cimitarra" And (ListBoxICA2.Value = "Actividades de Construcción" _
            Or ListBoxICA2.Value = "Transporte" Or ListBoxICA2.Value = "Servicios Informáticos" _
            Or ListBoxICA2.Value = "Actividades Profesionales" Or ListBoxICA2.Value = "Actividades Científicas" _
            Or ListBoxICA2.Value = "Actividades Técnicas" Or ListBoxICA2.Value = "Actividades de Seguridad Privada" _
            Or ListBoxICA2.Value = "Actividades de Servicios Administrativos" Or ListBoxICA2.Value = "Actividades de Servicios Comerciales" _
            Or ListBoxICA2.Value = "Actividades de Servicios Industriales" Or ListBoxICA2.Value = "Otras Actividades") Then
            .Clear
            Set icaMunicipio22 = .Add("04-06")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Cimitarra" And (ListBoxICA2.Value = "Actividades Finaniceras" _
            Or ListBoxICA2.Value = "Actividades de Alquiler y Arrendamiento") Then
            .Clear
            Set icaMunicipio22 = .Add("04-07")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Cimitarra" And ListBoxICA2.Value = "Compra Combustible" Then
            .Clear
            Set icaMunicipio22 = .Add("04-09")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Cimitarra" And ListBoxICA2.Value = "Otras Compras" Then
            .Clear
            Set icaMunicipio22 = .Add("04-08")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "San Carlos" And ListBoxICA2.Value = "Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("05-01")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "San Carlos" And ListBoxICA2.Value = "Compras" Then
            .Clear
            Set icaMunicipio22 = .Add("05-02")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "San Rafael" And ListBoxICA2.Value = "Compras" Then
            .Clear
            Set icaMunicipio22 = .Add("07-05")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "San Rafael" And ListBoxICA2.Value = "Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("07-06")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Victoria" And ListBoxICA2.Value = "Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("10-01")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Victoria" And ListBoxICA2.Value = "Compras" Then
            .Clear
            Set icaMunicipio22 = .Add("10-02")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Giron" And ListBoxICA2.Value = "Compras" Then
            .Clear
            Set icaMunicipio22 = .Add("13-07")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Giron" And ListBoxICA2.Value = "Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("13-02")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Bucaramanga" And ListBoxICA2.Value = "Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("14-04")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Bucaramanga" And ListBoxICA2.Value = "Compras" Then
            .Clear
            Set icaMunicipio22 = .Add("14-03")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Lebrija" And ListBoxICA2.Value = "Compras" Then
            .Clear
            Set icaMunicipio22 = .Add("15-01")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Lebrija" And ListBoxICA2.Value = "Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("15-02")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Betulia" And ListBoxICA2.Value = "Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("16-10")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Betulia" And ListBoxICA2.Value = "Compras" Then
            .Clear
            Set icaMunicipio22 = .Add("16-09")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Betulia" And ListBoxICA2.Value = "Servicios de Construcción" Then
            .Clear
            Set icaMunicipio22 = .Add("16-11")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Betulia" And ListBoxICA2.Value = "Restaurante y/o cafetería" Then
            .Clear
            Set icaMunicipio22 = .Add("16-12")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Betulia" And ListBoxICA2.Value = "Administración de Instalaciones" Then
            .Clear
            Set icaMunicipio22 = .Add("B1-13")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Betulia" And ListBoxICA2.Value = "Servicios de Seguridad Privada" Then
            .Clear
            Set icaMunicipio22 = .Add("16-14")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Medellín" And ListBoxICA2.Value = "Compra de Bienes y Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("18-01")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Zapatoca" And ListBoxICA2.Value = "Compra de Bienes y Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("23-01")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Uribia" And ListBoxICA2.Value = "Compra de Bienes y Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("25-01")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Salgar" And ListBoxICA2.Value = "Compra de Bienes y Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("27-01")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "San Luis" And ListBoxICA2.Value = "Compra de Bienes y Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("29-01")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Gómez Plata" And ListBoxICA2.Value = "Compra de Bienes y Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("30-01")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Santa Rosa de Osos" And ListBoxICA2.Value = "Compra de Bienes y Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("31-01")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "S.V Chucurí" And (ListBoxICA2.Value = "Servicios" _
            Or ListBoxICA2.Value = "Compras") Then
            .Clear
            Set icaMunicipio22 = .Add("19-04")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Norcasia" And ListBoxICA2.Value = "Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("20-04")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Norcasia" And ListBoxICA2.Value = "Compras" Then
            .Clear
            Set icaMunicipio22 = .Add("20-03")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Chaparral" And (ListBoxICA2.Value = "Obras Civiles" _
            Or ListBoxICA2.Value = "Actividades Comerciales al Por Mayor" Or ListBoxICA2.Value = "Actividades Comerciales al Por Menor" _
            Or ListBoxICA2.Value = "Transporte" Or ListBoxICA2.Value = "Actividades de  Edición e Impresión y Publicaciones" _
            Or ListBoxICA2.Value = "Otras Industrias Manufactureras" Or ListBoxICA2.Value = "Desarrollo de Sistemas de Informatica" _
            Or ListBoxICA2.Value = "Comercio de Vehículos") Then
            .Clear
            Set icaMunicipio22 = .Add("21-02")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Chaparral" And (ListBoxICA2.Value = "Publicidad" _
            Or ListBoxICA2.Value = " Alquiler Bien Mueble" Or ListBoxICA2.Value = "Actividades de agencias de viaje" _
            Or ListBoxICA2.Value = "Servicios de Seguridad e Investigación Privada" Or ListBoxICA2.Value = "Educación") Then
            .Clear
            Set icaMunicipio22 = .Add("21-03")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Chaparral" And (ListBoxICA2.Value = "Compra de Combustible y/o Lubricantes" _
            Or ListBoxICA2.Value = "Alojamientos" Or ListBoxICA2.Value = "Alimentación y Actividades de Catering") Then
            .Clear
            Set icaMunicipio22 = .Add("21-04")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Chaparral" And (ListBoxICA2.Value = "Mantenimiento y Reparación de Maquinaria y Equipos" _
            Or ListBoxICA2.Value = "Generación y Transmisión Eléctrica" Or ListBoxICA2.Value = "Distribución de Combustible" _
            Or ListBoxICA2.Value = "Actividades de Programación y Transmisión de Televisión" Or ListBoxICA2.Value = "Alimentación de Saneamiento Ambiental" _
            Or ListBoxICA2.Value = "Transporte por Tuberías") Then
            .Clear
            Set icaMunicipio22 = .Add("21-05")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Granada" And ListBoxICA2.Value = "Compras" Then
            .Clear
            Set icaMunicipio22 = .Add("22-01")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Granada" And ListBoxICA2.Value = "Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("22-02")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Puerto Wilches" And ListBoxICA2.Value = "Actividades Industriales" Then
            .Clear
            Set icaMunicipio22 = .Add("24-01")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Puerto Wilches" And ListBoxICA2.Value = "Actividades de Servicios y Comerciales" Then
            .Clear
            Set icaMunicipio22 = .Add("24-02")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Amalfi" And ListBoxICA2.Value = "Compras" Then
            .Clear
            Set icaMunicipio22 = .Add("26-01")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Amalfi" And ListBoxICA2.Value = "Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("26-02")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Cocorná" And ListBoxICA2.Value = "Bienes y Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("28-03")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Cocorná" And (ListBoxICA2.Value = "Mantenimiento y Reparación de Maquinaria y Equipo" _
            Or ListBoxICA2.Value = "Actividades de Limpieza" Or ListBoxICA2.Value = "Servicios Administrativos de Apoyo a las Empresas" _
            Or ListBoxICA2.Value = "Suministro de Vapor y Aire Acondicionado" Or ListBoxICA2.Value = "Actividades de Saneamiento Ambiental" _
            Or ListBoxICA2.Value = "Construcción de Edificio y Canteras" Or ListBoxICA2.Value = "Obras de Ingeniería Civil" _
            Or ListBoxICA2.Value = "Alquiler y Arrendamiento de Vehículos y Maquinaria") Then
            .Clear
            Set icaMunicipio22 = .Add("28-04")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Barrancabermeja" And ListBoxICA2.Value = "Servicios" Then
            .Clear
            Set icaMunicipio22 = .Add("17-07")
            icaMunicipio22.Select
        ElseIf ListBoxMunicipio2.Value = "Barrancabermeja" And ListBoxICA2.Value = "Compras" Then
            .Clear
            Set icaMunicipio22 = .Add("17-06")
            icaMunicipio22.Select
        Else
            .Clear
            Set icaMunicipio22 = .Add(" ")
            icaMunicipio22.Select
        End If
    Else
        .Clear
        Set icaMunicipio22 = .Add(" ")
        icaMunicipio22.Select
    End If
End With



End Sub
Sub miListaDesplegable()

Set xConcepto = ActiveDocument.FormFields("Listadesplegable1")

valorConcepto = xConcepto.Result
Call ListaDesplegable2

End Sub
Sub esconder()

ThisDocument.she

End Sub


Sub botonValidar_Click()

Call ListaDesplegable2
MsgBox "OK"

End Sub

Sub botonValidarICA1_Click()

Call ListaDesplegable2

End Sub

Private Sub botonValidarMunicipio1_Click()

End Sub

Sub botonValidarICA2_Click()

Call ListaDesplegable2

End Sub

Sub BotonValidarIVA_Click()



End Sub

Private Sub Document_BuildingBlockInsert(ByVal Range As Range, ByVal Name As String, ByVal Category As String, ByVal BlockType As String, ByVal Template As String)

End Sub

Sub Document_Close()

'MsgBox "¿Deseas guardar el proceso?", vbYesNo + vbExclamation

End Sub

Private Sub Document_ContentControlAfterAdd(ByVal NewContentControl As ContentControl, ByVal InUndoRedo As Boolean)

End Sub

Private Sub Document_ContentControlNonContentChange(ByVal ContentControl As ContentControl)

End Sub

Private Sub Document_New()

End Sub

Sub Document_Open()

ListBoxIVA.Height = 25.05
ListBoxRenta.Height = 25.05
ListBox1m.Height = 25.05
ListBoxIVADescripcion.Height = 25.05
ListBoxRentaDescripcion.Height = 25.05
ListBoxMunicipio1.Height = 25.05
ListBoxMunicipio2.Height = 25.05

ListBoxMunicipio1.Clear
ListBoxMunicipio1.Height = 36.7
ListBoxMunicipio1.AddItem "Bogotá"
ListBoxMunicipio1.AddItem "Cali"
ListBoxMunicipio1.AddItem "Barranquilla"
ListBoxMunicipio1.AddItem "Cimitarra"
ListBoxMunicipio1.AddItem "San Carlos"
ListBoxMunicipio1.AddItem "San Rafael"
ListBoxMunicipio1.AddItem "Victoria"
ListBoxMunicipio1.AddItem "Giron"
ListBoxMunicipio1.AddItem "Bucaramanga"
ListBoxMunicipio1.AddItem "Lebrija"
ListBoxMunicipio1.AddItem "Betulia"
ListBoxMunicipio1.AddItem "Barrancabermeja"
ListBoxMunicipio1.AddItem "Medellín"
ListBoxMunicipio1.AddItem "S.V Chucurí"
ListBoxMunicipio1.AddItem "Norcasia"
ListBoxMunicipio1.AddItem "Chaparral"
ListBoxMunicipio1.AddItem "Granada"
ListBoxMunicipio1.AddItem "Zapatoca"
ListBoxMunicipio1.AddItem "Puerto Wilches"
ListBoxMunicipio1.AddItem "Uribia"
ListBoxMunicipio1.AddItem "Amalfi"
ListBoxMunicipio1.AddItem "Salgar"
ListBoxMunicipio1.AddItem "Cocorná"
ListBoxMunicipio1.AddItem "San Luis"
ListBoxMunicipio1.AddItem "Gómez Plata"
ListBoxMunicipio1.AddItem "Santa Rosa de Osos"

ListBoxICA1.Clear

ListBoxMunicipio2.Clear
ListBoxMunicipio2.Height = 36.7
ListBoxMunicipio2.AddItem "Bogotá"
ListBoxMunicipio2.AddItem "Cali"
ListBoxMunicipio2.AddItem "Barranquilla"
ListBoxMunicipio2.AddItem "Cimitarra"
ListBoxMunicipio2.AddItem "San Carlos"
ListBoxMunicipio2.AddItem "San Rafael"
ListBoxMunicipio2.AddItem "Victoria"
ListBoxMunicipio2.AddItem "Giron"
ListBoxMunicipio2.AddItem "Bucaramanga"
ListBoxMunicipio2.AddItem "Lebrija"
ListBoxMunicipio2.AddItem "Betulia"
ListBoxMunicipio2.AddItem "Barrancabermeja"
ListBoxMunicipio2.AddItem "Medellín"
ListBoxMunicipio2.AddItem "S.V Chucurí"
ListBoxMunicipio2.AddItem "Norcasia"
ListBoxMunicipio2.AddItem "Chaparral"
ListBoxMunicipio2.AddItem "Granada"
ListBoxMunicipio2.AddItem "Zapatoca"
ListBoxMunicipio2.AddItem "Puerto Wilches"
ListBoxMunicipio2.AddItem "Uribia"
ListBoxMunicipio2.AddItem "Amalfi"
ListBoxMunicipio2.AddItem "Salgar"
ListBoxMunicipio2.AddItem "Cocorná"
ListBoxMunicipio2.AddItem "San Luis"
ListBoxMunicipio2.AddItem "Gómez Plata"
ListBoxMunicipio2.AddItem "Santa Rosa de Osos"

ListBoxICA2.Clear

ListBoxICA1.Height = 25.05
ListBoxICA2.Height = 25.05

OptionButton13.Value = False
OptionButton19.Value = False

OptionButton16.Value = False
OptionButton17.Value = False
OptionButton18.Value = False

'OptionButton13.Value = False
'OptionButton19.Value = False
'OptionButton20.Value = False
'OptionButton21.Value = False
'OptionButton101.Value = False
'OptionButton10.Value = False
'OptionButton16.Value = False
'OptionButton17.Value = False
'OptionButton18.Value = False
'OptionButton3.Value = False
'OptionButton22.Value = False
'OptionButton1.Value = False
'OptionButton2.Value = False
'OptionButton5.Value = False
'OptionButton4.Value = False
'OptionButton6.Value = False
'OptionButton7.Value = False
'OptionButton71.Value = False
'OptionButton29.Value = False
'OptionButton30.Value = False
'OptionButton32.Value = False

End Sub
Sub Municipios()

ListBox1m.Height = 25.05

ListBoxMunicipio1.Clear
ListBoxMunicipio1.Height = 25.05
ListBoxMunicipio1.AddItem "Bogotá"
ListBoxMunicipio1.AddItem "Cali"
ListBoxMunicipio1.AddItem "Barranquilla"
ListBoxMunicipio1.AddItem "Cimitarra"
ListBoxMunicipio1.AddItem "San Carlos"
ListBoxMunicipio1.AddItem "San Rafael"
ListBoxMunicipio1.AddItem "Victoria"
ListBoxMunicipio1.AddItem "Giron"
ListBoxMunicipio1.AddItem "Bucaramanga"
ListBoxMunicipio1.AddItem "Lebrija"
ListBoxMunicipio1.AddItem "Betulia"
ListBoxMunicipio1.AddItem "Barrancabermeja"
ListBoxMunicipio1.AddItem "Medellín"
ListBoxMunicipio1.AddItem "S.V Chucurí"
ListBoxMunicipio1.AddItem "Norcasia"
ListBoxMunicipio1.AddItem "Chaparral"
ListBoxMunicipio1.AddItem "Granada"
ListBoxMunicipio1.AddItem "Zapatoca"
ListBoxMunicipio1.AddItem "Puerto Wilches"
ListBoxMunicipio1.AddItem "Uribia"
ListBoxMunicipio1.AddItem "Amalfi"
ListBoxMunicipio1.AddItem "Salgar"
ListBoxMunicipio1.AddItem "Cocorná"
ListBoxMunicipio1.AddItem "San Luis"
ListBoxMunicipio1.AddItem "Gómez Plata"
ListBoxMunicipio1.AddItem "Santa Rosa de Osos"

ListBoxICA1.Clear

ListBoxMunicipio2.Clear
ListBoxMunicipio2.Height = 25.05
ListBoxMunicipio2.AddItem "Bogotá"
ListBoxMunicipio2.AddItem "Cali"
ListBoxMunicipio2.AddItem "Barranquilla"
ListBoxMunicipio2.AddItem "Cimitarra"
ListBoxMunicipio2.AddItem "San Carlos"
ListBoxMunicipio2.AddItem "San Rafael"
ListBoxMunicipio2.AddItem "Victoria"
ListBoxMunicipio2.AddItem "Giron"
ListBoxMunicipio2.AddItem "Bucaramanga"
ListBoxMunicipio2.AddItem "Lebrija"
ListBoxMunicipio2.AddItem "Betulia"
ListBoxMunicipio2.AddItem "Barrancabermeja"
ListBoxMunicipio2.AddItem "Medellín"
ListBoxMunicipio2.AddItem "S.V Chucurí"
ListBoxMunicipio2.AddItem "Norcasia"
ListBoxMunicipio2.AddItem "Chaparral"
ListBoxMunicipio2.AddItem "Granada"
ListBoxMunicipio2.AddItem "Zapatoca"
ListBoxMunicipio2.AddItem "Puerto Wilches"
ListBoxMunicipio2.AddItem "Uribia"
ListBoxMunicipio2.AddItem "Amalfi"
ListBoxMunicipio2.AddItem "Salgar"
ListBoxMunicipio2.AddItem "Cocorná"
ListBoxMunicipio2.AddItem "San Luis"
ListBoxMunicipio2.AddItem "Gómez Plata"
ListBoxMunicipio2.AddItem "Santa Rosa de Osos"

ListBoxICA2.Clear

ListBoxICA1.Height = 25.05
ListBoxICA2.Height = 25.05

Call ListaDesplegable2

End Sub

Private Sub Document_XMLAfterInsert(ByVal NewXMLNode As XMLNode, ByVal InUndoRedo As Boolean)

End Sub

Sub Lista1_Click()

'If OptionButton18.Value = True Then
'    ListBox1m.Clear
'    ListBox1m.AddItem "Donaciones"
'    ListBox1m.AddItem "Pagadores Alternos"
'    ListBox1m.AddItem "Embargos"
'    ListBox1m.AddItem "Convenios"
'    ListBox1m.AddItem "Alumbrado Público"
'    ListBox1m.AddItem "Nómina"
'    ListBox1m.AddItem "Gastos Bancarios"
'    ListBox1m.AddItem "Contribuciones"
'    ListBox1m.AddItem "Regalías"
'    ListBox1m.AddItem "Costas judiciales y gastos legales (reembolso)"
'    ListBox1m.AddItem "Deducciones de nómina"
'    ListBox1m.AddItem "Devoluciones de saldos"
'    ListBox1m.AddItem "Impuestos"
'    ListBox1m.AddItem "Pagos Tarjeta de crédito empresarial (Gerentes)"
'    ListBox1m.AddItem "Servicios Públicos"
'    ListBox1m.AddItem "Compra de predios (Personas naturales)"
'    ListBox1m.AddItem "Compensación por daño emergente"
'    ListBox1m.Height = 25.25
'End If

End Sub

Sub ListBox1_Click()

End Sub

Sub ListBox1m_Click()

Call ListaDesplegable2

End Sub

Sub ListBoxICA1_Click()

Call ListaDesplegable2

End Sub

Sub ListBoxICA2_Click()

Call ListaDesplegable2

End Sub

Sub ListBoxIVA_Click()

ListBoxIVADescripcion.Clear
ListBoxIVADescripcion.Height = 25.05

If ListBoxIVA.Value = "48 Impuesto sobre las ventas" Then
    ListBoxIVADescripcion.Clear
    ListBoxIVADescripcion.AddItem "Servicios Técnicos"
    ListBoxIVADescripcion.AddItem "Honorarios"
    ListBoxIVADescripcion.AddItem "Comisiones"
    ListBoxIVADescripcion.AddItem "Bienes Gravados 19%"
    ListBoxIVADescripcion.AddItem "Servicios 19%"
    ListBoxIVADescripcion.AddItem "Vehículos"
    ListBoxIVADescripcion.AddItem "Telefonía Móvil"
    ListBoxIVADescripcion.AddItem "Bienes Gravados 5%"
    ListBoxIVADescripcion.AddItem "Servicios 5%"
End If

Call ListaDesplegable2

End Sub

Sub ListBoxIVADescripcion_Click()

Call ListaDesplegable2

End Sub

Sub ListBoxMunicipio1_Click()

ListBoxICA1.Height = 25.05
ListBoxICA1.Clear

If ListBoxMunicipio1.Value = "Bogotá" Then
    ListBoxICA1.AddItem "Servicios de Restaurante"
    ListBoxICA1.AddItem "Servicios de Hotel"
    ListBoxICA1.AddItem "Servicios de Vigilancia"
    ListBoxICA1.AddItem "Compras y Otras Actividades Comerciales"
    ListBoxICA1.AddItem "Transporte de pasajeros"
    ListBoxICA1.AddItem "Edición de periódicos"
    ListBoxICA1.AddItem "Edición de revistas y/o otras publicaciones"
    ListBoxICA1.AddItem "Otros Servicios"
    ListBoxICA1.AddItem "Servcicio de Cosultoría"
    ListBoxICA1.AddItem "Servcicio de Interventoría"
    ListBoxICA1.AddItem "Seguros Generales de Vida"
    ListBoxICA1.AddItem "Servicios Financieros"
ElseIf ListBoxMunicipio1.Value = "Cali" Then
    ListBoxICA1.AddItem "Compras y Otras Actividades Comerciales"
    ListBoxICA1.AddItem "Servicios de Hotel"
    ListBoxICA1.AddItem "Servicios de Restaurante y Cafeteria"
    ListBoxICA1.AddItem "Sin Reteica"
    ListBoxICA1.AddItem "Otros Servicios"
ElseIf ListBoxMunicipio1.Value = "Barranquilla" Then
    ListBoxICA1.AddItem "Servicios de Restaurante"
    ListBoxICA1.AddItem "Servicios de aseo y limpieza"
    ListBoxICA1.AddItem "Otros Servicios"
    ListBoxICA1.AddItem "Otras Actividades Comerciales"
ElseIf ListBoxMunicipio1.Value = "Cimitarra" Then
    ListBoxICA1.AddItem "Actividades de Construcción"
    ListBoxICA1.AddItem "Transporte"
    ListBoxICA1.AddItem "Servicios Informáticos"
    ListBoxICA1.AddItem "Actividades Profesionales"
    ListBoxICA1.AddItem "Actividades Científicas"
    ListBoxICA1.AddItem "Actividades Técnicas"
    ListBoxICA1.AddItem "Actividades de Seguridad Privada"
    ListBoxICA1.AddItem "Actividades de Servicios Administrativos"
    ListBoxICA1.AddItem "Actividades de Servicios Comerciales"
    ListBoxICA1.AddItem "Actividades de Servicios Industriales"
    ListBoxICA1.AddItem "Actividades Finaniceras"
    ListBoxICA1.AddItem "Actividades de Alquiler y Arrendamiento"
    ListBoxICA1.AddItem "Compra Combustible"
    ListBoxICA1.AddItem "Otras Compras"
    ListBoxICA1.AddItem "Otras Actividades"
ElseIf ListBoxMunicipio1.Value = "San Rafael" Or ListBoxMunicipio1.Value = "Giron" _
    Or ListBoxMunicipio1.Value = "Lebrija" Then
    ListBoxICA1.AddItem "Servicios"
    ListBoxICA1.AddItem "Compras"
    ListBoxICA1.AddItem "Sin Reteica"
ElseIf ListBoxMunicipio1.Value = "San Carlos" _
    Or ListBoxMunicipio1.Value = "Victoria" _
    Or ListBoxMunicipio1.Value = "Bucaramanga" _
    Or ListBoxMunicipio1.Value = "Barrancabermeja" Or ListBoxMunicipio1.Value = "Norcasia" _
    Or ListBoxMunicipio1.Value = "Granada" Or ListBoxMunicipio1.Value = "Amalfi" _
    Or ListBoxMunicipio1.Value = "S.V Chucurí" Then
    ListBoxICA1.AddItem "Servicios"
    ListBoxICA1.AddItem "Compras"
ElseIf ListBoxMunicipio1.Value = "Betulia" Then
    ListBoxICA1.AddItem "Compras"
    ListBoxICA1.AddItem "Servicios"
    ListBoxICA1.AddItem "Servicios de Construcción"
    ListBoxICA1.AddItem "Restaurante y/o cafetería"
    ListBoxICA1.AddItem "Administración de Instalaciones"
    ListBoxICA1.AddItem "Servicios de Seguridad Privada"
    ListBoxICA1.AddItem "Sin Reteica"
ElseIf ListBoxMunicipio1.Value = "Medellín" Then
    ListBoxICA1.AddItem "Compra de Bienes y Servicios"
    ListBoxICA1.AddItem "Sin Reteica"
ElseIf ListBoxMunicipio1.Value = "Zapatoca" _
    Or ListBoxMunicipio1.Value = "Uribia" Or ListBoxMunicipio1.Value = "Salgar" _
    Or ListBoxMunicipio1.Value = "San Luis" Or ListBoxMunicipio1.Value = "Gómez Plata" _
    Or ListBoxMunicipio1.Value = "Santa Rosa de Osos" Then
    ListBoxICA1.AddItem "Compra de Bienes y Servicios"
ElseIf ListBoxMunicipio1.Value = "Chaparral" Then
    ListBoxICA1.AddItem "Obras Civiles"
    ListBoxICA1.AddItem "Actividades Comerciales al Por Mayor"
    ListBoxICA1.AddItem "Actividades Comerciales al Por Menor"
    ListBoxICA1.AddItem "Transporte"
    ListBoxICA1.AddItem "Actividades de  Edición e Impresión y Publicaciones"
    ListBoxICA1.AddItem "Otras Industrias Manufactureras"
    ListBoxICA1.AddItem "Desarrollo de Sistemas de Informatica"
    ListBoxICA1.AddItem "Comercio de Vehículos"
    ListBoxICA1.AddItem "Publicidad"
    ListBoxICA1.AddItem " Alquiler Bien Mueble"
    ListBoxICA1.AddItem "Actividades de agencias de viaje"
    ListBoxICA1.AddItem "Servicios de Seguridad e Investigación Privada"
    ListBoxICA1.AddItem "Educación"
    ListBoxICA1.AddItem "Compra de Combustible y/o Lubricantes"
    ListBoxICA1.AddItem "Alojamientos"
    ListBoxICA1.AddItem "Alimentación y Actividades de Catering"
    ListBoxICA1.AddItem "Mantenimiento y Reparación de Maquinaria y Equipos"
    ListBoxICA1.AddItem "Generación y Transmisión Eléctrica"
    ListBoxICA1.AddItem "Distribución de Combustible"
    ListBoxICA1.AddItem "Actividades de Programación y Transmisión de Televisión"
    ListBoxICA1.AddItem "Alimentación de Saneamiento Ambiental"
    ListBoxICA1.AddItem "Transporte por Tuberías"
ElseIf ListBoxMunicipio1.Value = "Puerto Wilches" Then
    ListBoxICA1.AddItem "Actividades Industriales"
    ListBoxICA1.AddItem "Actividades de Servicios y Comerciales"
ElseIf ListBoxMunicipio1.Value = "Cocorná" Then
    ListBoxICA1.AddItem "Bienes y Servicios"
    ListBoxICA1.AddItem "Mantenimiento y Reparación de Maquinaria y Equipo"
    ListBoxICA1.AddItem "Suministro de Vapor y Aire Acondicionado"
    ListBoxICA1.AddItem "Actividades de Saneamiento Ambiental"
    ListBoxICA1.AddItem "Construcción de Edificio y Canteras"
    ListBoxICA1.AddItem "Obras de Ingeniería Civil"
    ListBoxICA1.AddItem "Alquiler y Arrendamiento de Vehículos y Maquinaria"
    ListBoxICA1.AddItem "Actividades de Limpieza"
    ListBoxICA1.AddItem "Servicios Administrativos de Apoyo a las Empresas"
End If

Call ListaDesplegable2

End Sub

Sub ListBoxMunicipio2_Click()

ListBoxICA2.Height = 25.05
ListBoxICA2.Clear

If ListBoxMunicipio2.Value = "Bogotá" Then
    ListBoxICA2.AddItem "Servicios de Restaurante"
    ListBoxICA2.AddItem "Servicios de Hotel"
    ListBoxICA2.AddItem "Servicios de Vigilancia"
    ListBoxICA2.AddItem "Compras y Otras Actividades Comerciales"
    ListBoxICA2.AddItem "Transporte de pasajeros"
    ListBoxICA2.AddItem "Edición de periódicos"
    ListBoxICA2.AddItem "Edición de revistas y/o otras publicaciones"
    ListBoxICA2.AddItem "Otros Servicios"
    ListBoxICA2.AddItem "Servcicio de Cosultoría"
    ListBoxICA2.AddItem "Servcicio de Interventoría"
    ListBoxICA2.AddItem "Seguros Generales de Vida"
    ListBoxICA2.AddItem "Servicios Financieros"
ElseIf ListBoxMunicipio2.Value = "Cali" Then
    ListBoxICA2.AddItem "Compras y Otras Actividades Comerciales"
    ListBoxICA2.AddItem "Servicios de Hotel"
    ListBoxICA2.AddItem "Servicios de Restaurante y Cafeteria"
    ListBoxICA2.AddItem "Sin Reteica"
    ListBoxICA2.AddItem "Otros Servicios"
ElseIf ListBoxMunicipio2.Value = "Barranquilla" Then
    ListBoxICA2.AddItem "Servicios de Restaurante"
    ListBoxICA2.AddItem "Servicios de aseo y limpieza"
    ListBoxICA2.AddItem "Otros Servicios"
    ListBoxICA2.AddItem "Otras Actividades Comerciales"
ElseIf ListBoxMunicipio2.Value = "Cimitarra" Then
    ListBoxICA2.AddItem "Actividades de Construcción"
    ListBoxICA2.AddItem "Transporte"
    ListBoxICA2.AddItem "Servicios Informáticos"
    ListBoxICA2.AddItem "Actividades Profesionales"
    ListBoxICA2.AddItem "Actividades Científicas"
    ListBoxICA2.AddItem "Actividades Técnicas"
    ListBoxICA2.AddItem "Actividades de Seguridad Privada"
    ListBoxICA2.AddItem "Actividades de Servicios Administrativos"
    ListBoxICA2.AddItem "Actividades de Servicios Comerciales"
    ListBoxICA2.AddItem "Actividades de Servicios Industriales"
    ListBoxICA2.AddItem "Actividades Finaniceras"
    ListBoxICA2.AddItem "Actividades de Alquiler y Arrendamiento"
    ListBoxICA2.AddItem "Compra Combustible"
    ListBoxICA2.AddItem "Otras Compras"
    ListBoxICA2.AddItem "Otras Actividades"
    ElseIf ListBoxMunicipio2.Value = "San Rafael" Or ListBoxMunicipio2.Value = "Giron" _
        Or ListBoxMunicipio2.Value = "Lebrija" Then
        ListBoxICA2.AddItem "Servicios"
        ListBoxICA2.AddItem "Compras"
        ListBoxICA2.AddItem "Sin Reteica"
ElseIf ListBoxMunicipio2.Value = "San Carlos" _
    Or ListBoxMunicipio2.Value = "Victoria" _
    Or ListBoxMunicipio2.Value = "Bucaramanga" _
    Or ListBoxMunicipio2.Value = "Barrancabermeja" Or ListBoxMunicipio2.Value = "Norcasia" _
    Or ListBoxMunicipio2.Value = "Granada" Or ListBoxMunicipio2.Value = "Amalfi" _
    Or ListBoxMunicipio2.Value = "S.V Chucurí" Then
    ListBoxICA2.AddItem "Servicios"
    ListBoxICA2.AddItem "Compras"
ElseIf ListBoxMunicipio2.Value = "Betulia" Then
    ListBoxICA2.AddItem "Compras"
    ListBoxICA2.AddItem "Servicios"
    ListBoxICA2.AddItem "Servicios de Construcción"
    ListBoxICA2.AddItem "Restaurante y/o cafetería"
    ListBoxICA2.AddItem "Administración de Instalaciones"
    ListBoxICA2.AddItem "Servicios de Seguridad Privada"
    ListBoxICA2.AddItem "Sin Reteica"
ElseIf ListBoxMunicipio2.Value = "Medellín" Then
    ListBoxICA2.AddItem "Compra de Bienes y Servicios"
    ListBoxICA2.AddItem "Sin Reteica"
ElseIf ListBoxMunicipio2.Value = "Zapatoca" _
    Or ListBoxMunicipio2.Value = "Uribia" Or ListBoxMunicipio2.Value = "Salgar" _
    Or ListBoxMunicipio2.Value = "San Luis" Or ListBoxMunicipio2.Value = "Gómez Plata" _
    Or ListBoxMunicipio2.Value = "Santa Rosa de Osos" Then
    ListBoxICA2.AddItem "Compra de Bienes y Servicios"
ElseIf ListBoxMunicipio2.Value = "Chaparral" Then
    ListBoxICA2.AddItem "Obras Civiles"
    ListBoxICA2.AddItem "Actividades Comerciales al Por Mayor"
    ListBoxICA2.AddItem "Actividades Comerciales al Por Menor"
    ListBoxICA2.AddItem "Transporte"
    ListBoxICA2.AddItem "Actividades de  Edición e Impresión y Publicaciones"
    ListBoxICA2.AddItem "Otras Industrias Manufactureras"
    ListBoxICA2.AddItem "Desarrollo de Sistemas de Informatica"
    ListBoxICA2.AddItem "Comercio de Vehículos"
    ListBoxICA2.AddItem "Publicidad"
    ListBoxICA2.AddItem " Alquiler Bien Mueble"
    ListBoxICA2.AddItem "Actividades de agencias de viaje"
    ListBoxICA2.AddItem "Servicios de Seguridad e Investigación Privada"
    ListBoxICA2.AddItem "Educación"
    ListBoxICA2.AddItem "Compra de Combustible y/o Lubricantes"
    ListBoxICA2.AddItem "Alojamientos"
    ListBoxICA2.AddItem "Alimentación y Actividades de Catering"
    ListBoxICA2.AddItem "Mantenimiento y Reparación de Maquinaria y Equipos"
    ListBoxICA2.AddItem "Generación y Transmisión Eléctrica"
    ListBoxICA2.AddItem "Distribución de Combustible"
    ListBoxICA2.AddItem "Actividades de Programación y Transmisión de Televisión"
    ListBoxICA2.AddItem "Alimentación de Saneamiento Ambiental"
    ListBoxICA2.AddItem "Transporte por Tuberías"
ElseIf ListBoxMunicipio2.Value = "Puerto Wilches" Then
    ListBoxICA2.AddItem "Actividades Industriales"
    ListBoxICA2.AddItem "Actividades de Servicios y Comerciales"
ElseIf ListBoxMunicipio2.Value = "Cocorná" Then
    ListBoxICA2.AddItem "Bienes y Servicios"
    ListBoxICA2.AddItem "Mantenimiento y Reparación de Maquinaria y Equipo"
    ListBoxICA2.AddItem "Suministro de Vapor y Aire Acondicionado"
    ListBoxICA2.AddItem "Actividades de Saneamiento Ambiental"
    ListBoxICA2.AddItem "Construcción de Edificio y Canteras"
    ListBoxICA2.AddItem "Obras de Ingeniería Civil"
    ListBoxICA2.AddItem "Alquiler y Arrendamiento de Vehículos y Maquinaria"
    ListBoxICA2.AddItem "Actividades de Limpieza"
    ListBoxICA2.AddItem "Servicios Administrativos de Apoyo a las Empresas"
End If

Call ListaDesplegable2

End Sub

'Botón Impuesto Sobre las Ventas
Sub OptionButton1_Click()

botonISLV = True
botonNRDI = False
botonGC = False
boton22 = False

Call ListaDesplegable2

End Sub

Sub ListBoxRenta_Click()

ListBoxRentaDescripcion.Clear
ListBoxRentaDescripcion.Height = 25.05

If ListBoxRenta.Value = "05 Contribuyente Renta" Then
    ListBoxRentaDescripcion.Clear
    ListBoxRentaDescripcion.AddItem "Honorarios"
    ListBoxRentaDescripcion.AddItem "Comisiones  en General (Jurídicas)"
    ListBoxRentaDescripcion.AddItem "Servicios De Transporte De Carga"
    ListBoxRentaDescripcion.AddItem "Servicios Temporales (Personal)"
    ListBoxRentaDescripcion.AddItem "Transporte Aereo de Personas"
    ListBoxRentaDescripcion.AddItem "Servicios de Aseo y Vigilancia"
    ListBoxRentaDescripcion.AddItem "Servicios Integrales de Salud IPS"
    ListBoxRentaDescripcion.AddItem "Contratos De Construcción"
    ListBoxRentaDescripcion.AddItem "Servicios Generales - Declarantes"
    ListBoxRentaDescripcion.AddItem "Compra de Combustibles"
    ListBoxRentaDescripcion.AddItem "Compras y Suministros - Declarantes"
    ListBoxRentaDescripcion.AddItem "Intereses o Rendimientos Financieros"
    ListBoxRentaDescripcion.AddItem "Adquisición Bienes Raíces"
    ListBoxRentaDescripcion.AddItem "Adquisición de Vehículos"
    ListBoxRentaDescripcion.AddItem "Otros Ingresos Tributarios - Declarantes"
    ListBoxRentaDescripcion.AddItem "Arrendamientos Bienes Muebles"
    ListBoxRentaDescripcion.AddItem "Arrendamientos Bienes Raíces"
    ListBoxRentaDescripcion.AddItem "Servicios Transporte De Pasajeros"
    ListBoxRentaDescripcion.AddItem "Restaurante, Hotel y Hospedaje"
    ListBoxRentaDescripcion.AddItem "Indemnizaciones"
    ListBoxRentaDescripcion.AddItem "Software Bienes y Servicios"
    ListBoxRentaDescripcion.AddItem "Servicios Públicos"
End If

Call ListaDesplegable2

End Sub

Sub ListBoxRentaDescripcion_Click()

Call ListaDesplegable2

End Sub

Sub OptionButton10_Click()

End Sub
Sub OptionButton101_Click()


End Sub

Sub OptionButton11_Click()

End Sub
'Boton Persona Juridica
Sub OptionButton13_Click()

ListBox1m.Height = 25.05
ListBoxMunicipio1.Height = 25.05
ListBoxMunicipio2.Height = 25.05

contentFR = 3
contentFA = 4
contentICA = 5
Set fr = ActiveDocument.ContentControls(contentFR)
Set fa = ActiveDocument.ContentControls(contentFA)
Set ica = ActiveDocument.ContentControls(contentICA)

With fa.DropdownListEntries
    .Clear
    Set fa2 = fa.DropdownListEntries.Add(" ")
    fa2.Select
End With

With fr.DropdownListEntries
    .Clear
    Set fr2 = fr.DropdownListEntries.Add(" ")
    fr2.Select
End With

With ica.DropdownListEntries
    .Clear
    Set ica2 = ica.DropdownListEntries.Add(" ")
    ica2.Select
End With

ListBoxRentaDescripcion.Clear
ListBoxIVADescripcion.Clear

ListBoxIVA.Height = 25.05
ListBoxIVA.Clear
ListBoxIVA.AddItem "13 Gran Contribuyente"
ListBoxIVA.AddItem "23 Agentes de Retención en el IVA"
ListBoxIVA.AddItem "48 Impuesto sobre las ventas"
ListBoxIVA.AddItem "49 No responsable de IVA"

ListBoxRenta.Height = 25.05
ListBoxRenta.Clear
ListBoxRenta.AddItem "04 Régimen Tributario Especial"
ListBoxRenta.AddItem "05 Contribuyente Renta"
ListBoxRenta.AddItem "06 No Contribuyente de Renta"
ListBoxRenta.AddItem "15 Autorretenedor"
ListBoxRenta.AddItem "47 Régimen Simple Tributación"

Call Municipios

End Sub

Sub OptionButton14_Click()
End Sub
'Boton Bien
Sub OptionButton16_Click()

ListBox1m.Clear

botonBien = True
botonServicio = False
botonOtros = False

Call ListaDesplegable2

End Sub
'Boton Servicio
Sub OptionButton17_Click()

ListBox1m.Clear

botonBien = False
botonServicio = True
botonOtros = False

Call ListaDesplegable2

End Sub
'Boton Otros
Sub OptionButton18_Click()

botonBien = False
botonServicio = False
botonOtros = True

ListBox1m.Height = 25.05

    ListBox1m.Clear
    ListBox1m.AddItem "Donaciones"
    ListBox1m.AddItem "Pagadores Alternos"
    ListBox1m.AddItem "Embargos"
    ListBox1m.AddItem "Convenios"
    ListBox1m.AddItem "Alumbrado Público"
    ListBox1m.AddItem "Nómina"
    ListBox1m.AddItem "Gastos Bancarios"
    ListBox1m.AddItem "Contribuciones"
    ListBox1m.AddItem "Regalías"
    ListBox1m.AddItem "Costas judiciales y gastos legales (reembolso)"
    ListBox1m.AddItem "Deducciones de nómina"
    ListBox1m.AddItem "Devoluciones de saldos"
    ListBox1m.AddItem "Impuestos"
    ListBox1m.AddItem "Pagos Tarjeta de crédito empresarial (Gerentes)"
    ListBox1m.AddItem "Servicios Públicos"
    ListBox1m.AddItem "Compra de predios (Personas naturales)"
    ListBox1m.AddItem "Compensación por daño emergente"
    ListBox1m.Height = 25.05

Call ListaDesplegable2

End Sub
'Boton Si Facturador Electronico
Sub OptionButton20_Click()


End Sub
Sub OptionButton23_Click()

End Sub

Sub OptionButton26_Click()

End Sub
'Boton persona Natural
Sub OptionButton19_Click()

ListBox1m.Height = 25.05
ListBoxMunicipio1.Height = 25.05
ListBoxMunicipio2.Height = 25.05

contentFR = 3
contentFA = 4
contentICA = 5
Set fr = ActiveDocument.ContentControls(contentFR)
Set fa = ActiveDocument.ContentControls(contentFA)
Set ica = ActiveDocument.ContentControls(contentICA)

With fa.DropdownListEntries
    .Clear
    Set fa2 = fa.DropdownListEntries.Add(" ")
    fa2.Select
End With

With fr.DropdownListEntries
    .Clear
    Set fr2 = fr.DropdownListEntries.Add(" ")
    fr2.Select
End With

With ica.DropdownListEntries
    .Clear
    Set ica2 = ica.DropdownListEntries.Add(" ")
    ica2.Select
End With

ListBoxIVADescripcion.Clear
ListBoxRentaDescripcion.Clear

ListBoxIVA.Height = 25.05
ListBoxIVA.Clear
ListBoxIVA.AddItem "No Aplica"
ListBoxIVA.AddItem "48 Impuesto sobre las ventas"
ListBoxIVA.AddItem "49 No responsable de IVA"

ListBoxRenta.Height = 25.05
ListBoxRenta.Clear
ListBoxRenta.AddItem "No Declarante"
ListBoxRenta.AddItem "05 Contribuyente Renta"
ListBoxRenta.AddItem "47 Régimen Simple Tributación"

Call Municipios

End Sub


'Boton No Responsable de IVA
Sub OptionButton2_Click()

botonISLV = False
botonNRDI = True
botonGC = False
boton22 = False

Call ListaDesplegable2

End Sub
'Boton No Facturador Electronico
Sub OptionButton21_Click()


End Sub
'Boton Agentes de Retencion en el IVA
Sub OptionButton22_Click()

botonISLV = False
botonNRDI = False
botonGC = False
boton22 = True

Call ListaDesplegable2

End Sub

Sub OptionButton29_Click()

End Sub
'Boton Gran Contribuyente
Sub OptionButton3_Click()

botonISLV = False
botonNRDI = False
botonGC = True
boton22 = False

Call ListaDesplegable2

End Sub
Sub OptionButton30_Click()

End Sub

Sub OptionButton32_Click()

End Sub
'Boton Contributente Renta
Sub OptionButton4_Click()

botonCR = True
botonRTE = False
botonNCDR = False
botonAutorretenedor = False
botonRST = False

Call ListaDesplegable2

End Sub
'Boton Regimen Tributario Especial
Sub OptionButton5_Click()

botonCR = False
botonRTE = True
botonNCDR = False
botonAutorretenedor = False
botonRST = False

Call ListaDesplegable2

End Sub
'Boton No Contribuyente de Renta
Sub OptionButton6_Click()

botonCR = False
botonRTE = False
botonNCDR = True
botonAutorretenedor = False
botonRST = False

Call ListaDesplegable2

End Sub

Sub CheckBox1_Click()

End Sub

Sub CheckBox12_Click()

End Sub

Private Sub CheckBox14_Click()

End Sub

Sub CheckBox18_Click()

End Sub

Private Sub CheckBox5_Click()

End Sub
'Boton Autorretenedor
Sub OptionButton7_Click()

botonCR = False
botonRTE = False
botonNCDR = False
botonAutorretenedor = True
botonRST = False

Call ListaDesplegable2

End Sub
'Boton Regimen Simple Tributacion
Sub OptionButton71_Click()

botonCR = False
botonRTE = False
botonNCDR = False
botonAutorretenedor = False
botonRST = True

Call ListaDesplegable2

End Sub
Sub OptionButton8_Click()




End Sub
Sub OptionButton9_Click()



End Sub
