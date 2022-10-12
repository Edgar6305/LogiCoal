VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form RepMovimientos 
   Caption         =   "Reporte de Movimientos "
   ClientHeight    =   4245
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5190
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
   ScaleHeight     =   4245
   ScaleWidth      =   5190
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
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   4935
      Begin Threed.SSOption SSOption2 
         Height          =   255
         Left            =   180
         TabIndex        =   10
         Top             =   1740
         Width           =   1395
         _Version        =   65536
         _ExtentX        =   2461
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Traslados"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSOption SSOption1 
         Height          =   255
         Left            =   180
         TabIndex        =   9
         Top             =   1440
         Width           =   1635
         _Version        =   65536
         _ExtentX        =   2884
         _ExtentY        =   450
         _StockProps     =   78
         Caption         =   "Ventas"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComCtl2.DTPicker xFecFin 
         Height          =   315
         Left            =   2760
         TabIndex        =   1
         Top             =   900
         Visible         =   0   'False
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
         Format          =   115933187
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
         Format          =   115933187
         CurrentDate     =   39503.25
      End
      Begin Threed.SSFrame SSFrame4 
         Height          =   1095
         Left            =   120
         TabIndex        =   3
         Top             =   2100
         Width           =   4635
         _Version        =   65536
         _ExtentX        =   8176
         _ExtentY        =   1931
         _StockProps     =   14
         Caption         =   "Clientes / Acopios"
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
            ItemData        =   "RepMovimientos.frx":0000
            Left            =   1020
            List            =   "RepMovimientos.frx":0010
            TabIndex        =   4
            Top             =   420
            Width           =   2415
         End
         Begin VB.Image Image1 
            Height          =   375
            Left            =   3840
            Picture         =   "RepMovimientos.frx":0067
            Stretch         =   -1  'True
            Top             =   360
            Width           =   405
         End
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha Desde"
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   420
         Width           =   2235
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha Final"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   960
         Visible         =   0   'False
         Width           =   2235
      End
   End
   Begin KewlButtonz.KewlButtons Command1 
      Height          =   435
      Left            =   120
      TabIndex        =   7
      Top             =   3600
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
      MICON           =   "RepMovimientos.frx":15CD
      PICN            =   "RepMovimientos.frx":15E9
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
      Left            =   3360
      TabIndex        =   8
      Top             =   3600
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
      MICON           =   "RepMovimientos.frx":1B83
      PICN            =   "RepMovimientos.frx":1B9F
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
Attribute VB_Name = "RepMovimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xFec As Date, xTur As String, xSql As String, xMaq As String
Dim oT As New ADODB.Recordset
Dim FlagExcel As Boolean

Private Sub Command1_Click()
Dim xCount, i As Long
Dim xSql As String
Dim xR As New ADODB.Recordset

On Error GoTo Recover

If SSOption1 Then
    
    xSql = "SET DATEFORMAT DMY"
    xSql = xSql + " SELECT vMovimientosDS.IdTiquete, vMovimientosDS.Tercero, vMovimientosDS.FechaLleno, vMovimientosDS.PesoNeto, MovimientosDetalle.Cliente, MovimientosDetalle.FechaLlegada,"
    xSql = xSql + "        MovimientosDetalle.PesoLleno - MovimientosDetalle.PesoVacio AS NetoLlegada, "
    xSql = xSql + "        CASE WHEN MovimientosDetalle.Cliente IS NULL THEN DATEDIFF(hh, vMovimientosDS.FechaLleno, Getdate()) / 24 ELSE DATEDIFF(hh,vMovimientosDS.FechaLleno, MovimientosDetalle.FechaLlegada) / 24 END AS DiasRecorrido, "
    xSql = xSql + "        CASE WHEN MovimientosDetalle.PesoLleno IS NULL THEN vMovimientosDS.PesoNeto ELSE vMovimientosDS.PesoNeto - (MovimientosDetalle.PesoLleno - MovimientosDetalle.PesoVacio) END AS Diferencia, Terceros.Descripcion"
    xSql = xSql + " INTO   RepMovVentas"
    xSql = xSql + " FROM   Terceros INNER JOIN"
    xSql = xSql + "        vMovimientosDS ON Terceros.Identificacion = vMovimientosDS.Tercero LEFT OUTER JOIN"
    xSql = xSql + "        MovimientosDetalle ON vMovimientosDS.IdTiquete = MovimientosDetalle.Tiquete"
    xSql = xSql + " WHERE  (vMovimientosDS.FechaLleno >= '" & Format(xFecIni, "dd/MM/yyyy hh:mm") & "') AND (vMovimientosDS.Tercero = '" & Mid(Combo1, 1, 3) & "')"
    
    
    If file("RepMovVentas") Then Conn.Execute ("DROP TABLE RepMovVentas")
    Conn.Execute (xSql)
    
    If FlagExcel = True Then
        ExportaExcel ("RepMovVentas")
    Else
        Menu.oCr.Formulas(0) = "FI='" & Format(xFecIni.Value, "dd mmmm yyyy hh:mm") & "'"
        Menu.oCr.Formulas(1) = "FF='" & Format(xFecFin.Value, "dd mmmm yyyy hh:mm") & "'"
        Menu.oCr.ReportFileName = sDataReportPath + "RepMovVentas.Rpt"
    End If

ElseIf SSOption2 Then
    
    xSql = "SET DATEFORMAT DMY"
    xSql = xSql + " SELECT vMovimientosTR.IdTiquete, vMovimientosTR.Pesoneto, vMovimientosTR.Tercero, vMovimientosTR.FechaLleno, MovimientosDetalle.PesoLleno, MovimientosDetalle.PesoVacio, MovimientosDetalle.FechaLlegada,"
    xSql = xSql + "        CASE WHEN MovimientosDetalle.Cliente IS NULL THEN DATEDIFF(hh, vMovimientosTR.FechaLleno, Getdate()) / 24 ELSE DATEDIFF(hh,vMovimientosTR.FechaLleno, MovimientosDetalle.FechaLlegada) / 24 END AS DiasRecorrido,"
    xSql = xSql + "        CASE WHEN MovimientosDetalle.PesoLleno IS NULL  THEN vMovimientosTR.PesoNeto ELSE vMovimientosTR.PesoNeto - (MovimientosDetalle.PesoLleno - MovimientosDetalle.PesoVacio) END AS Diferencia, Terceros.Descripcion"
    xSql = xSql + " INTO   RepMovTraslados"
    xSql = xSql + " FROM   vMovimientosTR INNER JOIN"
    xSql = xSql + "        Terceros ON vMovimientosTR.Tercero = Terceros.Identificacion LEFT OUTER JOIN"
    xSql = xSql + "        MovimientosDetalle ON vMovimientosTR.IdTiquete = MovimientosDetalle.Tiquete"
    xSql = xSql + " WHERE  (vMovimientosTR.FechaLleno >= '23/05/2022') AND (vMovimientosTR.Tercero = '409')"
    
    If file("RepMovTraslados") Then Conn.Execute ("DROP TABLE RepMovTraslados")
    Conn.Execute (xSql)
    
    If FlagExcel = True Then
        ExportaExcel ("RepMovTraslados")
    Else
        Menu.oCr.Formulas(0) = "FI='" & Format(xFecIni.Value, "dd mmmm yyyy hh:mm") & "'"
        Menu.oCr.Formulas(1) = "FF='" & Format(xFecFin.Value, "dd mmmm yyyy hh:mm") & "'"
        Menu.oCr.ReportFileName = sDataReportPath + "RepMovTraslados.Rpt"
    End If
End If

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

End Sub

Private Sub Form_Activate()
    Me.SetFocus
    xFec = Now()
    xFecIni.Value = Format(xFec, "dd/MM/yyyy 00:00:00")
    xFecFin.Value = Format(xFec, "dd/MM/yyyy 23:59:59")
End Sub

Private Sub Image1_Click()
    FlagExcel = True
    Call Command1_Click
End Sub

Private Sub SSOption1_Click(Value As Integer)
Dim xR As New ADODB.Recordset
Dim xSql As String

Combo1.Clear
If SSOption1.Value Then
    xSql = "SET DATEFORMAT DMY"
    xSql = xSql + " SELECT DISTINCT vMovimientosDS.Tercero,Terceros.Descripcion"
    xSql = xSql + " FROM  Terceros INNER JOIN"
    xSql = xSql + "       vMovimientosDS ON Terceros.Identificacion = vMovimientosDS.Tercero LEFT OUTER JOIN"
    xSql = xSql + "       MovimientosDetalle ON vMovimientosDS.IdTiquete = MovimientosDetalle.Tiquete"
    xSql = xSql + " WHERE (vMovimientosDS.FechaLleno >='" & Format(xFecIni, "dd/MM/yyyy hh:mm") & "')"
    
    Set xR = Conn.Execute(xSql)
    If Not xR.EOF Then Combo1.text = xR!Tercero & " " & xR!Descripcion
    Do While Not xR.EOF
        Combo1.AddItem xR!Tercero & " " & xR!Descripcion
    xR.MoveNext
    Loop
    xR.Close
End If
End Sub

Private Sub SSOption2_Click(Value As Integer)
Dim xR As New ADODB.Recordset
Dim xSql As String

Combo1.Clear
If SSOption2.Value Then
    xSql = "SET DATEFORMAT DMY"
    xSql = xSql + " SELECT DISTINCT vMovimientosTR.Tercero, Terceros.Descripcion"
    xSql = xSql + " FROM    vMovimientosTR INNER JOIN"
    xSql = xSql + "         Terceros ON vMovimientosTR.Tercero = Terceros.Identificacion LEFT OUTER JOIN"
    xSql = xSql + "         MovimientosDetalle ON vMovimientosTR.IdTiquete = MovimientosDetalle.Tiquete"
    xSql = xSql + " WHERE   (vMovimientosTR.FechaLleno >='" & Format(xFecIni, "dd/MM/yyyy hh:mm") & "')"
    
    Set xR = Conn.Execute(xSql)
    If Not xR.EOF Then Combo1.text = xR!Tercero & " " & xR!Descripcion
    Do While Not xR.EOF
        Combo1.AddItem xR!Tercero & " " & xR!Descripcion
    xR.MoveNext
    Loop
    xR.Close
End If
End Sub
