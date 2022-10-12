VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RepSeguimiento 
   Caption         =   "Reporte de Seguimiento de Mulas"
   ClientHeight    =   4860
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5220
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
   ScaleHeight     =   4860
   ScaleWidth      =   5220
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
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4935
      Begin VB.Frame Frame2 
         Caption         =   "Tipo de Informe"
         Height          =   1455
         Left            =   420
         TabIndex        =   9
         Top             =   2220
         Width           =   4275
         Begin Threed.SSOption Tipo 
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Mulas en transito  "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin Threed.SSOption Tipo 
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   11
            Top             =   420
            Width           =   1755
            _Version        =   65536
            _ExtentX        =   3096
            _ExtentY        =   344
            _StockProps     =   78
            Caption         =   "Diferencia               "
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Alignment       =   1
         End
         Begin VB.Image Image1 
            Height          =   375
            Left            =   3240
            Picture         =   "RepSeguimiento.frx":0000
            Stretch         =   -1  'True
            Top             =   480
            Width           =   405
         End
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         Left            =   2220
         TabIndex        =   8
         Top             =   1560
         Width           =   2475
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "RepSeguimiento.frx":1566
         Left            =   420
         List            =   "RepSeguimiento.frx":1570
         TabIndex        =   7
         Text            =   "TRASLADOS"
         Top             =   1560
         Width           =   1575
      End
      Begin MSComCtl2.DTPicker xFecFin 
         Height          =   315
         Left            =   1800
         TabIndex        =   1
         Top             =   900
         Width           =   1335
         _ExtentX        =   2355
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59375619
         CurrentDate     =   39503.75
      End
      Begin MSComCtl2.DTPicker xFecIni 
         Height          =   315
         Left            =   1800
         TabIndex        =   2
         Top             =   420
         Width           =   1335
         _ExtentX        =   2355
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
         CustomFormat    =   "dd/MM/yyyy"
         Format          =   59375619
         CurrentDate     =   39503.25
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Inicial"
         Height          =   315
         Left            =   420
         TabIndex        =   4
         Top             =   420
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   420
         TabIndex        =   3
         Top             =   960
         Width           =   1095
      End
   End
   Begin KewlButtonz.KewlButtons Command1 
      Height          =   435
      Left            =   60
      TabIndex        =   5
      Top             =   4140
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
      MICON           =   "RepSeguimiento.frx":1587
      PICN            =   "RepSeguimiento.frx":15A3
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
      Left            =   3240
      TabIndex        =   6
      Top             =   4140
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
      MICON           =   "RepSeguimiento.frx":1B3D
      PICN            =   "RepSeguimiento.frx":1B59
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
Attribute VB_Name = "RepSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FlagExcel As Boolean
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
    xFecFin = Now
    xFecIni = Now
End Sub

Private Sub Combo1_Click()
Dim xSql As String
Dim xR As New ADODB.Recordset

On Error GoTo Recover

If Combo1 = "TRASLADOS" Then
    xSql = "Set DateFormat DMY"
    xSql = xSql + " SELECT   DISTINCT Acopios.Descripcion AS DesAcopios"
    xSql = xSql + " FROM     Bascula INNER JOIN"
    xSql = xSql + "          Transportador ON Bascula.IdTransportador = Transportador.IdTransportador INNER JOIN"
    xSql = xSql + "          Traslados ON Bascula.NumeroTransaccion = Traslados.IdTraslado INNER JOIN"
    xSql = xSql + "          Pilas ON Traslados.PilaDestino = Pilas.IdPila INNER JOIN"
    xSql = xSql + "          Acopios ON Pilas.IdAcopio = Acopios.IdAcopio LEFT OUTER JOIN"
    xSql = xSql + "          MovimientosDetalle ON Bascula.IdTiquete = MovimientosDetalle.Tiquete"
    xSql = xSql + " WHERE    (Bascula.FechaTurno >='" & Format(xFecIni, "dd/MM/yyyy") & "' AND Bascula.FechaTurno <='" & Format(xFecFin, "dd/MM/yyyy") & "') AND (Bascula.Estado = 'AC') AND (Bascula.TransaccionOrigen = 'TR')"
    xSql = xSql + " ORDER BY Acopios.Descripcion"
    Set xR = Conn.Execute(xSql)
    If Not xR.EOF Then Combo2 = xR!DesAcopios
    Do While Not xR.EOF
        Combo2.AddItem xR!DesAcopios
    xR.MoveNext
    Loop
    xR.Close
Else
    xSql = "Set DateFormat DMY"
    xSql = xSql + " SELECT   DISTINCT Terceros.Descripcion AS DesTercero"
    xSql = xSql + " FROM     MovimientosDetalle RIGHT OUTER JOIN"
    xSql = xSql + "          Terceros INNER JOIN"
    xSql = xSql + "          Ventas ON Terceros.IdCliente = Ventas.IdCliente INNER JOIN"
    xSql = xSql + "          Bascula INNER JOIN"
    xSql = xSql + "          Transportador ON Bascula.IdTransportador = Transportador.IdTransportador ON Ventas.IdVentas = Bascula.NumeroTransaccion ON MovimientosDetalle.Tiquete = Bascula.IdTiquete"
    xSql = xSql + " WHERE    (Bascula.FechaTurno >='" & Format(xFecIni, "dd/MM/yyyy") & "' AND Bascula.FechaTurno <='" & Format(xFecFin, "dd/MM/yyyy") & "') AND (Bascula.Estado = 'AC') AND (Bascula.TransaccionOrigen = 'DS')"
    xSql = xSql + " ORDER BY Terceros.Descripcion"
    Set xR = Conn.Execute(xSql)
    If Not xR.EOF Then Combo2 = xR!DesTercero
    Do While Not xR.EOF
        Combo2.AddItem xR!DesTercero
    xR.MoveNext
    Loop
    xR.Close

End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error Al Elegir Opcion Ventas o Traslados," & vbCrLf & Err.Description
    MsgBox MSG, , "Combo1_Click()"
    Err.Clear   ' Borra campos del objeto Err
    Exit Sub
End If
End Sub

Private Sub Command1_Click()
Dim xSql As String
Dim xR As New ADODB.Recordset

On Error GoTo Recover
Select Case Combo1

Case "TRASLADOS"
    xSql = "Set DateFormat DMY"
    xSql = xSql + " SELECT    Bascula.IdTiquete, Bascula.Placas, Bascula.IdTransportador, Transportador.Descripcion, Bascula.PesoLleno - Bascula.PesoVacio AS Neto, Bascula.FechaLleno, MovimientosDetalle.Tiquete, MovimientosDetalle.Fecha,"
    xSql = xSql + "           MovimientosDetalle.PesoNeto, "
    xSql = xSql + "           CASE WHEN  MovimientosDetalle.Fecha=NULL THEN DATEDIFF(Hour, Bascula.FechaLleno, MovimientosDetalle.Fecha) ELSE DATEDIFF(Hour, Bascula.FechaLleno,'" & Format(Now, "dd/MM/yyyy hh:mm") & "') END AS HorasDiferencia,"
    xSql = xSql + "           ROUND(ABS((Bascula.PesoLleno - Bascula.PesoVacio - MovimientosDetalle.PesoNeto) / (Bascula.PesoLleno - Bascula.PesoVacio)), 3)*100 AS Porcentaje, Bascula.TransaccionOrigen, Bascula.NumeroTransaccion, Pilas.IdAcopio,"
    xSql = xSql + "           Acopios.Descripcion AS DesAcopio"
    xSql = xSql + " INTO      RepMulasTR"
    xSql = xSql + " FROM      Bascula INNER JOIN"
    xSql = xSql + "           Transportador ON Bascula.IdTransportador = Transportador.IdTransportador INNER JOIN"
    xSql = xSql + "           Traslados ON Bascula.NumeroTransaccion = Traslados.IdTraslado INNER JOIN"
    xSql = xSql + "           Pilas ON Traslados.PilaDestino = Pilas.IdPila INNER JOIN"
    xSql = xSql + "           Acopios ON Pilas.IdAcopio = Acopios.IdAcopio LEFT OUTER JOIN"
    xSql = xSql + "           MovimientosDetalle ON Bascula.IdTiquete = MovimientosDetalle.Tiquete"
    
    If Tipo(1) Then
        xSql = xSql + " WHERE    (Bascula.FechaTurno >='" & Format(xFecIni, "dd/MM/yyyy") & "' AND Bascula.FechaTurno <='" & Format(xFecFin, "dd/MM/yyyy") & "') AND (Bascula.Estado = 'AC') AND (Bascula.TransaccionOrigen = 'TR') AND MovimientosDetalle.Tiquete IS NOT NULL AND Acopios.Descripcion='" & Combo2 & "'"
    Else
        xSql = xSql + " WHERE    (Bascula.FechaTurno >='" & Format(xFecIni, "dd/MM/yyyy") & "' AND Bascula.FechaTurno <='" & Format(xFecFin, "dd/MM/yyyy") & "') AND (Bascula.Estado = 'AC') AND (Bascula.TransaccionOrigen = 'TR') AND MovimientosDetalle.Tiquete IS  NULL AND Acopios.Descripcion='" & Combo2 & "'"
    End If
    xSql = xSql + " ORDER BY Bascula.IdTiquete"

    If file("RepMulasTR") Then Conn.Execute ("DROP TABLE RepMulasTR")
    Conn.Execute (xSql)
    If FlagExcel = True Then
        ExportaExcel ("RepMulasTR")
    Else
    End If

Case "VENTAS"

    xSql = "Set DateFormat DMY"
    xSql = xSql + " SELECT   Bascula.IdTiquete, Bascula.Placas, Bascula.IdTransportador, Transportador.Descripcion, Bascula.PesoLleno - Bascula.PesoVacio AS Neto, Bascula.FechaLleno, MovimientosDetalle.Tiquete, MovimientosDetalle.Fecha,"
    xSql = xSql + "          MovimientosDetalle.PesoNeto, CASE WHEN MovimientosDetalle.Fecha = NULL THEN DATEDIFF(Hour, Bascula.FechaLleno, MovimientosDetalle.Fecha) ELSE DATEDIFF(Hour, Bascula.FechaLleno, GETDATE()) END AS HorasDiferencia,"
    xSql = xSql + "          Bascula.PesoLleno - Bascula.PesoVacio - MovimientosDetalle.PesoNeto AS PesoDif, ROUND(ABS((Bascula.PesoLleno - Bascula.PesoVacio - MovimientosDetalle.PesoNeto) / (Bascula.PesoLleno - Bascula.PesoVacio)), 3) * 100 AS Porcentaje,"
    xSql = xSql + "          Bascula.TransaccionOrigen, Bascula.NumeroTransaccion, Ventas.IdVentas, Terceros.Descripcion AS DesTercero"
    xSql = xSql + " INTO      RepMulasDS"
    xSql = xSql + " FROM     MovimientosDetalle RIGHT OUTER JOIN"
    xSql = xSql + "          Terceros INNER JOIN"
    xSql = xSql + "          Ventas ON Terceros.IdCliente = Ventas.IdCliente INNER JOIN"
    xSql = xSql + "          Bascula INNER JOIN"
    xSql = xSql + "          Transportador ON Bascula.IdTransportador = Transportador.IdTransportador ON Ventas.IdVentas = Bascula.NumeroTransaccion ON MovimientosDetalle.Tiquete = Bascula.IdTiquete"
    If Tipo(1) Then
        xSql = xSql + " WHERE    (Bascula.FechaTurno >='" & Format(xFecIni, "dd/MM/yyyy") & "' AND Bascula.FechaTurno <='" & Format(xFecFin, "dd/MM/yyyy") & "') AND (Bascula.Estado = 'AC') AND (Bascula.TransaccionOrigen = 'DS') AND MovimientosDetalle.Tiquete IS NOT NULL AND Terceros.Descripcion='" & Combo2 & "'"
    Else
        xSql = xSql + " WHERE    (Bascula.FechaTurno >='" & Format(xFecIni, "dd/MM/yyyy") & "' AND Bascula.FechaTurno <='" & Format(xFecFin, "dd/MM/yyyy") & "') AND (Bascula.Estado = 'AC') AND (Bascula.TransaccionOrigen = 'DS') AND MovimientosDetalle.Tiquete IS  NULL AND Terceros.Descripcion='" & Combo2 & "'"
    End If

    xSql = xSql + " ORDER BY Bascula.IdTiquete"
    If file("RepMulasDS") Then Conn.Execute ("DROP TABLE RepMulasDS")
    Conn.Execute (xSql)
    
    If FlagExcel = True Then
        ExportaExcel ("RepMulasDS")
    Else
    End If
    
End Select

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

Private Sub Image1_Click()
    FlagExcel = True
    Call Command1_Click
End Sub
