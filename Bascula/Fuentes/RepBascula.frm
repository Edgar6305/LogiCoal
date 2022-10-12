VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RepBascula 
   Caption         =   "Reporte de Recepciones"
   ClientHeight    =   4020
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5175
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3075
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4935
      Begin MSComCtl2.DTPicker xFecFin 
         Height          =   315
         Left            =   2760
         TabIndex        =   1
         Top             =   900
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy hh:mm tt"
         Format          =   116064259
         CurrentDate     =   39503.75
      End
      Begin MSComCtl2.DTPicker xFecIni 
         Height          =   315
         Left            =   2760
         TabIndex        =   2
         Top             =   420
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "dd/MM/yyyy hh:mm tt"
         Format          =   116064259
         CurrentDate     =   39503.25
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   1095
         Left            =   180
         TabIndex        =   3
         Top             =   1620
         Width           =   4635
         _Version        =   65536
         _ExtentX        =   8176
         _ExtentY        =   1931
         _StockProps     =   14
         Caption         =   "Informes "
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI Emoji"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "RepBascula.frx":0000
            Left            =   1020
            List            =   "RepBascula.frx":0019
            TabIndex        =   4
            Text            =   "1- RECEPCION"
            Top             =   420
            Width           =   2415
         End
         Begin VB.Image Image1 
            Height          =   375
            Left            =   3840
            Picture         =   "RepBascula.frx":00A9
            Stretch         =   -1  'True
            Top             =   360
            Width           =   405
         End
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   180
         TabIndex        =   6
         Top             =   960
         Width           =   2235
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   420
         Width           =   2235
      End
   End
   Begin KewlButtonz.KewlButtons Command1 
      Height          =   435
      Left            =   180
      TabIndex        =   7
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Ejecutar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "RepBascula.frx":160F
      PICN            =   "RepBascula.frx":162B
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons Command2 
      Height          =   435
      Left            =   3300
      TabIndex        =   8
      Top             =   3360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "RepBascula.frx":1BC5
      PICN            =   "RepBascula.frx":1BE1
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "RepBascula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xFec As Date, xTur As String, xSql As String, xMaq As String
Dim oT As New ADODB.Recordset
Dim FlagExcel As Boolean

Private Sub Command1_Click()
Dim xCount, i As Long
Dim xEx As String
Dim xR As New ADODB.Recordset

On Error GoTo Recover

Select Case Mid(Combo1, 1, 1)
Case 1
    xSql = "Set DateFormat DMY "
    xSql = xSql + " SELECT   0 AS No, Bascula.IdTiquete, Bascula.FechaTurno,'FRONTIERNEXT-PL' AS Origen_Excel, Minas.Descripcion AS DesMIna, Lotes.Bloque + ' ' + OperadoresMineros.Descripcion AS Pit_Excel,"
    xSql = xSql + "          Lotes.Panel, Lotes.Manto, Bascula.Documentoasociado, Transportador.Descripcion AS DesTransportador,Bascula.Placas, 'ROM' As TProd_Excel, Bascula.PesoLleno, Bascula.PesoVacio, Bascula.PesoLleno -  Bascula.PesoVacio AS PesoNeto,"
    xSql = xSql + "          Year(Bascula.FechaTurno) as Anio_Excel, DATENAME(Month,Bascula.FechaTurno) As Mes, 'Produccion' as Prod_Excel, PilasFisicas.TipoCarbon AS DesTipoPila, DATEPART(isoww,Bascula.FechaTurno) as Semana_Excel, DATEPART(isoww,Bascula.FechaTurno) as Semana_Excel1,"
    xSql = xSql + "          Minas.Descripcion AS DesMIna_Excel, PilasFisicas.Descripcion AS DesPila, Bascula.FechaTurno AS Fec_Excel, OperadoresMineros.Descripcion AS DesOperador, Lotes.Bloque,"
    xSql = xSql + "          Bascula.TransaccionOrigen, Bascula.NumeroTransaccion, Materiales.Descripcion AS DesMaterial,Bascula.Conductor, Bascula.FechaLleno, Bascula.FechaVacio, Bascula.UsoTara, Bascula.Observaciones,"
    xSql = xSql + "          Conductores.Nombre, Bascula.Estado, Lotes.Tajo"
    xSql = xSql + "  INTO    RepRecepciones"
    xSql = xSql + "  FROM    Minas INNER JOIN"
    xSql = xSql + "          Lotes ON Minas.IdMina = Lotes.IdMina INNER JOIN"
    xSql = xSql + "          Pilas ON Lotes.Pila = Pilas.IdPila INNER JOIN"
    xSql = xSql + "          PilasFisicas ON Pilas.IdPilaFisica = PilasFisicas.IdPilaFisica INNER JOIN"
    xSql = xSql + "          Bascula ON Lotes.IdLote = Bascula.NumeroTransaccion INNER JOIN"
    xSql = xSql + "          Acopios ON Pilas.IdAcopio = Acopios.IdAcopio INNER JOIN"
    xSql = xSql + "          Transportador ON Bascula.IdTransportador = Transportador.IdTransportador INNER JOIN"
    xSql = xSql + "          OperadoresMineros ON Lotes.Operador = OperadoresMineros.IdOperador INNER JOIN"
    xSql = xSql + "          Materiales ON Bascula.IdMaterial = Materiales.IdMaterial LEFT OUTER JOIN"
    xSql = xSql + "          Conductores ON Bascula.Conductor = Conductores.Cedula"
    
    xSql = xSql + " WHERE  (Bascula.TransaccionOrigen = 'LT') AND (Bascula.Estado='AC') AND (Bascula.Fechalleno >= '" & Format(xFecIni, "dd/MM/yyyy hh:mm") & "') AND (Bascula.Fechalleno <= '" & Format(xFecFin, "dd/MM/yyyy hh:mm") & "')"
    
    If file("RepRecepciones") Then Conn.Execute ("DROP TABLE RepRecepciones")
    Conn.Execute (xSql)
    'Conn.Execute ("Delete RepRecepciones")
    'Conn.Execute ("Set DateFormat DMY Insert INTO RepRecepciones " & xSql)
    
    If FlagExcel = True Then
        ExportaExcel ("RepRecepciones")
    Else
        Menu.oCr.Formulas(0) = "FI='" & Format(xFecIni.Value, "dd mmmm yyyy hh:mm") & "'"
        Menu.oCr.Formulas(1) = "FF='" & Format(xFecFin.Value, "dd mmmm yyyy hh:mm") & "'"
        Menu.oCr.ReportFileName = sDataReportPath + "RepRecepcion.Rpt"
    End If
    
Case 2
    xSql = "Set DateFormat DMY"
    xSql = xSql + " SELECT 0 AS No,Bascula.IdTiquete, Bascula.FechaTurno, 'FRONTIERNEXT-PL' AS Origen_Excel, 'MINA LA ESTANCIA' AS Mina_Excell, dbo.FS_NombreAcopio(Bascula.NumeroTransaccion) AS DesAcopio, "
    xSql = xSql + "        Terceros.Descripcion, 'ND' As Manto_excel, Bascula.Documentoasociado , Transportador.Descripcion AS DesTransportador, Bascula.Placas, TiposCarbon.Descripcion AS DesTipoCarbon,"
    
    xSql = xSql + "        Bascula.PesoLleno, Bascula.PesoVacio, Bascula.PesoLleno -  Bascula.PesoVacio AS PesoNeto, Year(Bascula.FechaTurno) as Anio_Excel, DATENAME(Month,Bascula.FechaTurno) As Mes, Terceros.Descripcion AS Ter_Excel1, Terceros.Descripcion AS Ter_Excel2,"
    xSql = xSql + "        Ventas.CantidadPedida, Ventas.CantidadDespachada, Conductores.Nombre,  Bascula.FechaLleno, Bascula.FechaVacio, Bascula.Estado,"
    xSql = xSql + "        Bascula.TransaccionOrigen, Bascula.NumeroTransaccion, Bascula.Carpado,  Bascula.IdMaterial, Bascula.IdTransportador, Bascula.Conductor,Bascula.Observaciones"
    xSql = xSql + " INTO   RepDespachos"
    xSql = xSql + " FROM   Ventas INNER JOIN"
    xSql = xSql + "        Bascula ON Ventas.IdVentas = Bascula.NumeroTransaccion INNER JOIN"
    xSql = xSql + "        Terceros ON Ventas.IdCliente = Terceros.IdCliente INNER JOIN"
    xSql = xSql + "        Transportador ON Bascula.IdTransportador = Transportador.IdTransportador LEFT OUTER JOIN"
    xSql = xSql + "        TiposCarbon ON Bascula.IdTipoCarbon = TiposCarbon.IdTipoCarbon LEFT OUTER JOIN"
    xSql = xSql + "        Conductores ON Bascula.Conductor = Conductores.Cedula"
    xSql = xSql + " WHERE  (Bascula.TransaccionOrigen = 'DS') AND (Bascula.Estado='AC') AND (Bascula.FechaVacio >= '" & Format(xFecIni, "dd/MM/yyyy hh:mm") & "') AND (Bascula.FechaVacio <= '" & Format(xFecFin, "dd/MM/yyyy hh:mm") & "')"
    
    If file("RepDespachos") Then Conn.Execute ("DROP TABLE RepDespachos")
    Conn.Execute (xSql)
    
    If FlagExcel = True Then
        ExportaExcel ("RepDespachos")
    End If
    
    Menu.oCr.Formulas(0) = "FI='" & Format(xFecIni.Value, "dd mmmm yyyy hh:mm") & "'"
    Menu.oCr.Formulas(1) = "FF='" & Format(xFecFin.Value, "dd mmmm yyyy hh:mm") & "'"
    Menu.oCr.ReportFileName = sDataReportPath + "RepDespachos.Rpt"
    
Case 3
        xSql = "Set DateFormat DMY "
        xSql = xSql + " SELECT  0 AS No,Bascula.IdTiquete, Bascula.FechaTurno, 'FRONTIERNEXT-PL' AS Origen_Excel, 'MINA LA ESTANCIA' AS Mina_Excell, dbo.FS_NombreAcopio(Bascula.NumeroTransaccion) AS DesAcopio_Excell, "
        xSql = xSql + "         Acopios.Descripcion AS DesAcopio, 'ND' as Manto, Bascula.Documentoasociado , Transportador.Descripcion, TiposCarbon.Descripcion AS DesTipoCarbon, "
        xSql = xSql + "         Bascula.Placas,  Bascula.PesoLleno, Bascula.PesoVacio,Bascula.PesoLleno -  Bascula.PesoVacio AS PesoNeto,"
        xSql = xSql + "         Year(Bascula.FechaTurno) as Anio_Excel, DATENAME(Month,Bascula.FechaTurno) As Mes, Acopios.Descripcion AS Cliente_Excel, Acopios.Descripcion AS Cliente2_Excel,"
        xSql = xSql + "         Bascula.FechaVacio,Bascula.FechaLleno, Bascula.Observaciones, Traslados.Cantidad, Traslados.CantidadDespachada, PilasFisicas.Descripcion AS DesPila,"
        xSql = xSql + "         Acopios.Ubicacion, Traslados.Fecha AS FechaOrdenTraslado, Bascula.Usuario, Usuarios_T.Descripcion AS DesUsuario, Usuarios_T.Cargo, Bascula.Estado,"
        xSql = xSql + "         Bascula.Carpado, Bascula.TransaccionOrigen, Bascula.NumeroTransaccion, Bascula.Conductor"
        xSql = xSql + "  INTO   RepTraslados"
        xSql = xSql + "  FROM   PilasFisicas INNER JOIN"
        xSql = xSql + "         Pilas ON PilasFisicas.IdPilaFisica = Pilas.IdPilaFisica INNER JOIN"
        xSql = xSql + "         Bascula INNER JOIN"
        xSql = xSql + "         Transportador ON Bascula.IdTransportador = Transportador.IdTransportador INNER JOIN"
        xSql = xSql + "         Traslados ON Bascula.TransaccionOrigen = 'TR' AND Bascula.NumeroTransaccion = Traslados.IdTraslado ON Pilas.IdPila = Traslados.PilaDestino INNER JOIN"
        xSql = xSql + "         Acopios ON Pilas.IdAcopio = Acopios.IdAcopio LEFT OUTER JOIN"
        xSql = xSql + "         TiposCarbon ON Bascula.IdTipoCarbon = TiposCarbon.IdTipoCarbon LEFT OUTER JOIN"
        xSql = xSql + "         Usuarios_T ON Bascula.Usuario = Usuarios_T.Login"
        xSql = xSql + "  WHERE  (Bascula.TransaccionOrigen = 'TR') AND (Bascula.Estado='AC') AND (Bascula.FechaVacio >= '" & Format(xFecIni, "dd/MM/yyyy hh:mm") & "') AND (Bascula.FechaVacio <= '" & Format(xFecFin, "dd/MM/yyyy hh:mm") & "')"
        
        If file("RepTraslados") Then Conn.Execute ("DROP TABLE RepTraslados")
        Conn.Execute (xSql)
'            Conn.Execute ("Delete RepTraslados")
'            Conn.Execute ("Set Dateformat dmy Insert INTO RepTraslados" & xSql)
        
        If FlagExcel = True Then
            ExportaExcel ("RepTraslados")
        End If
        
        Menu.oCr.ReportFileName = sDataReportPath + "RepTraslados.Rpt"

Case 4
        xSql = "Set DateFormat DMY "
        xSql = xSql + " SELECT Bascula.IdTiquete, Bascula.Documentoasociado, Materiales.Descripcion, Transportador.Descripcion AS DesTransportador, Bascula.Placas, Bascula.Conductor, Bascula.PesoLleno, Bascula.PesoVacio, Bascula.FechaLleno, Bascula.FechaVacio,"
        xSql = xSql + "        Bascula.PesoLleno -  Bascula.PesoVacio AS PesoNeto, Bascula.FechaLlegada , Bascula.UsoTara, Bascula.Observaciones, Bascula.Usuario,Usuarios_T.Descripcion AS DesUsuario, Usuarios_T.Cargo, Bascula.Estado"
        xSql = xSql + " INTO   RepRecepcionOtros"
        xSql = xSql + " FROM   Bascula INNER JOIN"
        xSql = xSql + "        Materiales ON Bascula.IdMaterial = Materiales.IdMaterial INNER JOIN"
        xSql = xSql + "        Transportador ON Bascula.IdTransportador = Transportador.IdTransportador INNER JOIN"
        xSql = xSql + "        Usuarios_T ON Bascula.Usuario = Usuarios_T.Login"
        xSql = xSql + " WHERE  (Bascula.TransaccionOrigen = 'RO') AND (Bascula.Estado='AC') AND (Bascula.FechaVacio >= '" & Format(xFecIni, "dd/MM/yyyy hh:mm") & "') AND (Bascula.FechaVacio <= '" & Format(xFecFin, "dd/MM/yyyy hh:mm") & "')"
        
        If file("RepRecepcionOtros") Then Conn.Execute ("DROP TABLE RepRecepcionOtros")
        Conn.Execute (xSql)
'        Conn.Execute ("Delete RepRecepcionOtros")
'        Conn.Execute ("Set Dateformat dmy Insert INTO RepRecepcionOtros" & xSql)
        
        If FlagExcel = True Then
            ExportaExcel ("RepRecepcionOtros")
        End If
       
        Menu.oCr.ReportFileName = sDataReportPath + "RepRecepcionOtros.Rpt"
Case 5

        xSql = " Set DateFormat DMY "
        xSql = xSql + " SELECT  Lotes.IdLote, PilasFisicas.Descripcion, Lotes.Cantidad, Lotes.FechaApertura, Lotes.FechaCierre, OperadoresMineros.Descripcion AS DesOperador, Minas.Descripcion AS DesMina, Lotes.Nivel, Lotes.Panel, Lotes.Manto,"
        xSql = xSql + "         Lotes.Tajo , Lotes.Bloque, Lotes.Estado"
        xSql = xSql + " INTO    RepLotes"
        xSql = xSql + " FROM    Lotes INNER JOIN"
        xSql = xSql + "         Minas ON Lotes.IdMina = Minas.IdMina INNER JOIN"
        xSql = xSql + "         Pilas ON Lotes.Pila = Pilas.IdPila INNER JOIN"
        xSql = xSql + "         PilasFisicas ON Pilas.IdPilaFisica = PilasFisicas.IdPilaFisica INNER JOIN"
        xSql = xSql + "         OperadoresMineros ON Lotes.Operador = OperadoresMineros.IdOperador"
        xSql = xSql + " WHERE   (Lotes.FechaApertura >= '" & Format(xFecIni, "dd/MM/yyyy hh:mm") & "') "
           
        If file("RepLotes") Then Conn.Execute ("DROP TABLE RepLotes")
        Conn.Execute (xSql)
       
        If FlagExcel = True Then
            ExportaExcel ("RepLotes")
        Else
            Menu.oCr.Formulas(0) = "FI='" & Format(xFecIni.Value, "dd mmmm yyyy hh:mm") & "'"
            Menu.oCr.Formulas(1) = "FF='" & Format(xFecFin.Value, "dd mmmm yyyy hh:mm") & "'"
            Menu.oCr.ReportFileName = sDataReportPath + "RepLotes.Rpt"
        End If
Case 6

        xSql = " Set DateFormat DMY "
        xSql = xSql + " SELECT  Bascula.IdTiquete, Bascula.TransaccionOrigen, Bascula.NumeroTransaccion, Bascula.FechaLleno, Transportador.Descripcion, Bascula.Placas, Placas.TipoVehiculo"
        xSql = xSql + " INTO    RepCarpado"
        xSql = xSql + " FROM    Bascula INNER JOIN"
        xSql = xSql + "         Placas ON Bascula.Placas = Placas.Placas INNER JOIN"
        xSql = xSql + "         Transportador ON Bascula.IdTransportador = Transportador.IdTransportador"
        xSql = xSql + " WHERE  (Bascula.Carpado = 1) AND (Bascula.FechaTurno >= '" & Format(xFecIni, "dd/MM/yyyy hh:mm") & "') AND (Bascula.FechaTurno <= '" & Format(xFecFin, "dd/MM/yyyy hh:mm") & "')"

        If file("RepCarpado") Then Conn.Execute ("DROP TABLE RepCarpado")
        Conn.Execute (xSql)
       
        If FlagExcel = True Then
            ExportaExcel ("RepCarpado")
        Else
            Menu.oCr.Formulas(0) = "FI='" & Format(xFecIni.Value, "dd mmmm yyyy hh:mm") & "'"
            Menu.oCr.Formulas(1) = "FF='" & Format(xFecFin.Value, "dd mmmm yyyy hh:mm") & "'"
            Menu.oCr.ReportFileName = sDataReportPath + "RepCarpado.Rpt"
        End If
        
        xSql = " Set DateFormat DMY "
        xSql = xSql + " SELECT    Bascula.Placas, COUNT(Bascula.Placas)  AS Cantidad, MIN(Bascula.Fechalleno) AS FechaInicial, MAX(Bascula.Fechalleno) AS FechaFinal, Placas.TipoVehiculo , Transportador.Descripcion"
        xSql = xSql + " INTO      RepCarpadoRes"
        xSql = xSql + " FROM      Bascula INNER JOIN         Placas ON Bascula.Placas = Placas.Placas INNER JOIN         Transportador ON Bascula.IdTransportador = Transportador.IdTransportador"
        xSql = xSql + " WHERE     (Bascula.Carpado = 1) AND (Bascula.FechaTurno >= '" & Format(xFecIni, "dd/MM/yyyy hh:mm") & "') AND (Bascula.FechaTurno <= '" & Format(xFecFin, "dd/MM/yyyy hh:mm") & "')"
        xSql = xSql + " GROUP BY  Bascula.Placas, Placas.TipoVehiculo , Transportador.Descripcion"
        
        If file("RepCarpadoRes") Then Conn.Execute ("DROP TABLE RepCarpadoRes")
        Conn.Execute (xSql)
        
        If FlagExcel = True Then
            ExportaExcel ("RepCarpadoRes")
        Else
            Menu.oCr.Formulas(0) = "FI='" & Format(xFecIni.Value, "dd mmmm yyyy hh:mm") & "'"
            Menu.oCr.Formulas(1) = "FF='" & Format(xFecFin.Value, "dd mmmm yyyy hh:mm") & "'"
            Menu.oCr.ReportFileName = sDataReportPath + "RepCarpadoRes.Rpt"
        End If
Case 7

        xSql = " Set DateFormat DMY "
        xSql = xSql + " SELECT    Trituracion.IdTrituracion, Trituradoras.Descripcion, Trituracion.FechaInicio, Trituracion.FechaCierre, PilasFisicas.Descripcion AS DesPilaOrigen,"
        xSql = xSql + "           DATEDIFF(HOUR, Trituracion.FechaInicio, Trituracion.FechaCierre) AS HorasOrden, ROUND(Trituracion.HorasEfectivas,0) HorasEfectivas,"
        xSql = xSql + "           (SELECT SUM(DATEDIFF(HOUR,FechaInicio, FechaFin)) FROM   TrituradoraParos WHERE IdTrituracion=Trituracion.IdTrituracion) AS HorasParos,Trituracion.Cantidad AS CantidadTotal,"
        xSql = xSql + "           (SELECT  Top 1 TD.Porcentaje  FROM  TrituracionDetalle AS TD WHERE TD.IdTrituracion =Trituracion.IdTrituracion) AS Porcentaje1,"
        xSql = xSql + "           (SELECT  Top 1 TD.Porcentaje  FROM  TrituracionDetalle AS TD WHERE TD.IdTrituracion =Trituracion.IdTrituracion)*Trituracion.Cantidad/100 AS CantProcesada1,"
        xSql = xSql + "           (SELECT  Top 1 PF.Descripcion FROM PilasFisicas AS PF INNER JOIN  Pilas ON PF.IdPilaFisica = Pilas.IdPilaFisica INNER JOIN  TrituracionDetalle AS TD ON Pilas.IdPila = TD.PilaDestino WHERE   TD.IdTrituracion =Trituracion.IdTrituracion) AS TipoDes1,"
        xSql = xSql + "           (SELECT  Top 1 TD.Porcentaje  FROM  TrituracionDetalle AS TD WHERE TD.IdTrituracion =Trituracion.IdTrituracion Order By TD.IdTrituracionDetalle DESC) AS Porcentaje2,"
        xSql = xSql + "           (SELECT  Top 1 TD.Porcentaje  FROM  TrituracionDetalle AS TD WHERE TD.IdTrituracion =Trituracion.IdTrituracion Order By TD.IdTrituracionDetalle DESC)*Trituracion.Cantidad/100 AS CantProcesada2,"
        xSql = xSql + "           (SELECT  Top 1 PF.Descripcion FROM PilasFisicas AS PF INNER JOIN  Pilas ON PF.IdPilaFisica = Pilas.IdPilaFisica INNER JOIN  TrituracionDetalle AS TD ON Pilas.IdPila = TD.PilaDestino"
        xSql = xSql + "            WHERE   TD.IdTrituracion =Trituracion.IdTrituracion ORDER BY TD.IdTrituracionDetalle DESC) AS TipoDes2"
        xSql = xSql + " INTO      RepTrituracionRES"
        xSql = xSql + " FROM      PilasFisicas INNER JOIN"
        xSql = xSql + "           Pilas ON PilasFisicas.IdPilaFisica = Pilas.IdPilaFisica INNER JOIN"
        xSql = xSql + "           Trituracion INNER JOIN"
        xSql = xSql + "           Trituradoras ON Trituracion.IdTrituradora = Trituradoras.IdTrituradora ON Pilas.IdPila = Trituracion.PilaOrigen"
        xSql = xSql + " WHERE    (Trituradoras.ProduccionHora > 0) AND (Trituracion.FechaInicio >='" & Format(xFecIni, "dd/MM/yyyy hh:mm") & "') AND (Trituracion.FechaCierre <='" & Format(xFecFin, "dd/MM/yyyy hh:mm") & "')"
        xSql = xSql + " ORDER BY Trituracion.IdTrituracion"
        
        If file("RepTrituracionRES") Then Conn.Execute ("DROP TABLE RepTrituracionRES")
        Conn.Execute (xSql)
        
        If FlagExcel = True Then
            ExportaExcel ("RepTrituracionRES")
        Else
            Menu.oCr.Formulas(0) = "FI='" & Format(xFecIni.Value, "dd mmmm yyyy hh:mm") & "'"
            Menu.oCr.Formulas(1) = "FF='" & Format(xFecFin.Value, "dd mmmm yyyy hh:mm") & "'"
            Menu.oCr.ReportFileName = sDataReportPath + "RepTrituracionRES.Rpt"
        End If

End Select

If FlagExcel = False Then
    Menu.oCr.Action = 1
    Call BorraRpt(Menu.oCr, 1)
End If

FlagExcel = False
    
Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error en Generación de Reporte," & vbCrLf & Err.Description
    MsgBox MSG, , "Reportes"
    Err.Clear   ' Borra campos del objeto Err
    Exit Sub
End If
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim xFecIni As Date
Dim xFecFin As Date

'Combo1.Clear
'Combo1.Text = oT!CodigoActividad & " " & oT!Descripcion
'Do While Not oT.EOF
'    Combo1.AddItem oT!CodigoActividad & " " & oT!Descripcion
'oT.MoveNext
'Loop
'oT.Close

End Sub

Private Sub Form_Activate()
    Me.SetFocus
    xFec = Now()
    xFecIni.Value = Format(xFec, "dd/MM/yyyy 06:00:00")
    xFecFin.Value = Format(xFec, "dd/MM/yyyy 18:00:00")
End Sub

Private Sub Image1_Click()
    FlagExcel = True
    Call Command1_Click
End Sub
