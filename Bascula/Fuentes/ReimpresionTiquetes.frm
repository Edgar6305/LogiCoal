VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Begin VB.Form ReimpresionTiquetes 
   Caption         =   "Reimpresion de Tiquetes"
   ClientHeight    =   2415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5175
   LinkTopic       =   "Form1"
   ScaleHeight     =   2415
   ScaleWidth      =   5175
   StartUpPosition =   1  'CenterOwner
   Begin Threed.SSFrame SSFrame4 
      Height          =   1095
      Left            =   120
      TabIndex        =   0
      Top             =   300
      Width           =   4815
      _Version        =   65536
      _ExtentX        =   8493
      _ExtentY        =   1931
      _StockProps     =   14
      Caption         =   "Reimpresión de Tiquetes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Emoji"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.TextBox NumeroTiquete 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   3060
         TabIndex        =   4
         Text            =   "0"
         Top             =   540
         Width           =   1215
      End
      Begin VB.Line Line1 
         X1              =   1200
         X2              =   3950
         Y1              =   860
         Y2              =   860
      End
      Begin VB.Label Label1 
         Caption         =   "Numero de Tiquete"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   3
         Top             =   540
         Width           =   1635
      End
   End
   Begin KewlButtonz.KewlButtons Command1 
      Height          =   435
      Left            =   480
      TabIndex        =   1
      Top             =   1620
      Visible         =   0   'False
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
      MICON           =   "ReimpresionTiquetes.frx":0000
      PICN            =   "ReimpresionTiquetes.frx":001C
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
      Left            =   3060
      TabIndex        =   2
      Top             =   1620
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
      MICON           =   "ReimpresionTiquetes.frx":05B6
      PICN            =   "ReimpresionTiquetes.frx":05D2
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
Attribute VB_Name = "ReimpresionTiquetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xR As New ADODB.Recordset
Dim xNumero As Long

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Activate()
    Me.SetFocus
End Sub

Private Sub Command1_Click()
Dim xSql As String
Dim xTrTipo As String


On Error GoTo Recover

    Select Case xR!IdTransaccion
    Case 1
    
        If xR!TransaccionOrigen = "LT" Then
            xSql = "           SELECT  Bascula.IdTiquete, Bascula.Documentoasociado As Remision, Lotes.IdLote, Lotes.Panel,lotes.Nivel, lotes.Bloque, Lotes.Manto, Bascula.Placas, Placas.Conductor, Placas.TipoVehiculo, Transportador.Descripcion AS DesTransportador,"
            xSql = xSql + "                 Minas.Descripcion AS DesMinas, OperadoresMineros.Descripcion AS DesOperador, PilasFisicas.Descripcion, PilasFisicas.TipoCarbon, Bascula.PesoLleno, Bascula.PesoVacio, Bascula.FechaLleno,"
            xSql = xSql + "                 Bascula.FechaVacio, Bascula.Usuario, Usuarios_T.Descripcion AS DesUsuarioNombre, Usuarios_T.Cargo, Bascula.Observaciones, Bascula.UsoTara"
            xSql = xSql + " FROM     Lotes INNER JOIN"
            xSql = xSql + "                 Bascula INNER JOIN"
            xSql = xSql + "                 Placas ON Bascula.Placas = Placas.Placas INNER JOIN"
            xSql = xSql + "                 Transportador ON Placas.IdTransportador = Transportador.IdTransportador ON Lotes.IdLote = Bascula.NumeroTransaccion AND 'LT' = Bascula.TransaccionOrigen INNER JOIN"
            xSql = xSql + "                 Minas ON Lotes.IdMina = Minas.IdMina INNER JOIN"
            xSql = xSql + "                 OperadoresMineros ON Lotes.Operador = OperadoresMineros.IdOperador INNER JOIN"
            xSql = xSql + "                 TiposCarbon ON Lotes.IdTipoCarbon = TiposCarbon.IdTipoCarbon INNER JOIN"
            xSql = xSql + "                 Pilas ON Lotes.Pila = Pilas.IdPila INNER JOIN"
            xSql = xSql + "                 PilasFisicas ON Pilas.IdPilaFisica = PilasFisicas.IdPilaFisica INNER JOIN"
            xSql = xSql + "                 Usuarios_T ON Bascula.Usuario = Usuarios_T.Login"
            xSql = xSql + " Where    Bascula.IdTiquete = " & xNumero
        
            Conn.Execute ("Delete RepTiqueteRecepcion")
            Conn.Execute ("Set Dateformat dmy Insert INTO RepTiqueteRecepcion " & xSql)
            Menu.oCr.ReportFileName = sDataReportPath + "RepTiqueteRecepcion.Rpt"
            Menu.oCr.Action = 1
            
        Else
            xSql = " SELECT Bascula.IdTiquete, Bascula.Documentoasociado, Materiales.Descripcion, Transportador.Descripcion AS DesTransportador, Bascula.Placas, Bascula.Conductor, Bascula.PesoLleno, Bascula.PesoVacio, Bascula.FechaLleno, Bascula.FechaVacio,"
            xSql = xSql + "                  Bascula.FechaLlegada , Bascula.UsoTara, Bascula.Observaciones, Bascula.Usuario,Usuarios_T.Descripcion AS DesUsuario, Usuarios_T.Cargo"
            xSql = xSql + "   FROM  Bascula INNER JOIN"
            xSql = xSql + "                 Materiales ON Bascula.IdMaterial = Materiales.IdMaterial INNER JOIN"
            xSql = xSql + "                 Transportador ON Bascula.IdTransportador = Transportador.IdTransportador INNER JOIN"
            xSql = xSql + "                 Usuarios_T ON Bascula.Usuario = Usuarios_T.Login"
            xSql = xSql + " Where    Bascula.IdTiquete = " & xNumero
            
            Conn.Execute ("Delete RepTiqueteRecepcionOtros")
            Conn.Execute ("Set Dateformat dmy Insert INTO RepTiqueteRecepcionOtros " & xSql)
            Menu.oCr.ReportFileName = sDataReportPath + "RepTiqueteRecepcionOtros.Rpt"
            Menu.oCr.Action = 1
        End If
               
    Case 2
            
        If xR!TransaccionOrigen = "DS" Then ' <== VENTAS
            xSql = "        SELECT Bascula.IdTiquete, Bascula.Documentoasociado AS Remision, Bascula.Placas, Bascula.Conductor, Bascula.PesoLleno, Bascula.PesoVacio, "
            xSql = xSql + "        Bascula.FechaLleno, Bascula.FechaVacio, Bascula.FechaLlegada, Bascula.Usuario AS DesUsuario, Bascula.Observaciones,"
            xSql = xSql + "        Materiales.Descripcion AS Desmaterial, Transportador.Descripcion AS DesTransportador, PilasFisicas.Descripcion, Terceros.Identificacion, Terceros.Descripcion AS DesTercero, "
            xSql = xSql + "        Ventas.OrdenCompraCliente , Acopios.Descripcion AS DesAcopio, Acopios.Ubicacion, Ventas.CantidadPedida, Ventas.CantidadDespachada , Usuarios_T.Descripcion AS DesUsuario_T, "
            xSql = xSql + "        Usuarios_T.Cargo, Conductores.Nombre, Bascula.IdTipoCarbon"
            xSql = xSql + " FROM   Bascula INNER JOIN"
            xSql = xSql + "        Ventas ON Bascula.TransaccionOrigen = 'DS' AND Bascula.NumeroTransaccion = Ventas.IdVentas INNER JOIN"
            xSql = xSql + "        VentasDetalle ON Ventas.IdVentas = VentasDetalle.IdVenta INNER JOIN"
            xSql = xSql + "        Pilas ON VentasDetalle.IdPila = Pilas.IdPila INNER JOIN"
            xSql = xSql + "        PilasFisicas ON Pilas.IdPilaFisica = PilasFisicas.IdPilaFisica INNER JOIN"
            xSql = xSql + "        Materiales ON Bascula.IdMaterial = Materiales.IdMaterial INNER JOIN"
            xSql = xSql + "        Transportador ON Bascula.IdTransportador = Transportador.IdTransportador INNER JOIN"
            xSql = xSql + "        Terceros ON Ventas.IdCliente = Terceros.IdCliente INNER JOIN"
            xSql = xSql + "        Acopios ON Pilas.IdAcopio = Acopios.IdAcopio INNER JOIN"
            xSql = xSql + "        Usuarios_T ON Bascula.Usuario = Usuarios_T.Login INNER JOIN"
            xSql = xSql + "        Conductores ON Bascula.Conductor = Conductores.Cedula"
            xSql = xSql + " WHERE  Bascula.IdTiquete = " & xNumero
        
            Conn.Execute ("Delete RepTiqueteDespacho")
            Conn.Execute ("Set Dateformat dmy Insert INTO RepTiqueteDespacho" & xSql)
            Menu.oCr.ReportFileName = sDataReportPath + "RepTiqueteDespacho.Rpt"
            Menu.oCr.Action = 1
        Else
            xSql = ""
            xSql = xSql + " SELECT  Bascula.IdTiquete, Bascula.Documentoasociado, Transportador.Descripcion, Bascula.Placas, Bascula.Conductor, Bascula.PesoLleno, Bascula.PesoVacio,"
            xSql = xSql + "         Bascula.FechaLleno, Bascula.FechaVacio, Bascula.Observaciones, Traslados.Cantidad, Traslados.CantidadDespachada, PilasFisicas.Descripcion AS DesPila,"
            xSql = xSql + "         Acopios.Descripcion AS DesAcopio, Acopios.Ubicacion, Traslados.Fecha AS FechaOrdenTraslado, Bascula.Usuario, Usuarios_T.Descripcion AS DesUsuario, "
            xSql = xSql + "         Usuarios_T.Cargo, Conductores.Nombre, Bascula.IdTipoCarbon"
            xSql = xSql + " FROM    PilasFisicas INNER JOIN"
            xSql = xSql + "         Pilas ON PilasFisicas.IdPilaFisica = Pilas.IdPilaFisica INNER JOIN"
            xSql = xSql + "         Bascula INNER JOIN"
            xSql = xSql + "         Transportador ON Bascula.IdTransportador = Transportador.IdTransportador INNER JOIN"
            xSql = xSql + "         Traslados ON Bascula.TransaccionOrigen = 'TR' AND Bascula.NumeroTransaccion = Traslados.IdTraslado ON Pilas.IdPila = Traslados.PilaDestino INNER JOIN"
            xSql = xSql + "         Acopios ON Pilas.IdAcopio = Acopios.IdAcopio INNER JOIN"
            xSql = xSql + "         Usuarios_T ON Bascula.Usuario = Usuarios_T.Login INNER JOIN"
            xSql = xSql + "         Conductores ON Bascula.Conductor = Conductores.Cedula"
            xSql = xSql + " WHERE   Bascula.IdTiquete = " & xNumero
            
            Conn.Execute ("Delete RepTiqueteTraslados")
            Conn.Execute ("Set Dateformat dmy Insert INTO RepTiqueteTraslados" & xSql)
            Menu.oCr.ReportFileName = sDataReportPath + "RepTiqueteTraslado.Rpt"
            Menu.oCr.Action = 1
        End If
    End Select
    
    Call BorraRpt(Menu.oCr, 1)

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al Imprimir," & vbCrLf & Err.Description
    MsgBox MSG, , "ImprimeTiquete_Click()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub

Private Sub NumeroTiquete_GotFocus()
Call Mark(NumeroTiquete)
End Sub

Private Sub NumeroTiquete_KeyDown(KeyCode As Integer, Shift As Integer)
Dim xSql As String

Select Case KeyCode
Case vbKeyDown, vbKeyReturn, vbKeyTab
    Set xR = Conn.Execute("Select * From Bascula Where IdTiquete=" & NumeroTiquete)
    If xR.EOF Then
        MsgBox "El Tiquete NO se Localiza, Verifique"
        NumeroTiquete.SetFocus
    End If
    xNumero = NumeroTiquete
    Command1.Visible = Not xR.EOF
End Select

End Sub

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
                dControl.SelStart = 0
                dControl.SelLength = Len(dControl.text)
End If
End Sub
