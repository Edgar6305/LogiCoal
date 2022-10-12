VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Comprobantes 
   Caption         =   "Comprobantes SieSa"
   ClientHeight    =   3120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6105
   Icon            =   "Comprobantes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3120
   ScaleWidth      =   6105
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   180
      ScaleHeight     =   1965
      ScaleWidth      =   5685
      TabIndex        =   0
      Top             =   180
      Width           =   5715
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "Comprobantes.frx":000C
         Left            =   1680
         List            =   "Comprobantes.frx":000E
         TabIndex        =   1
         Top             =   960
         Width           =   3795
      End
      Begin MSComCtl2.DTPicker xFecFin 
         Height          =   315
         Left            =   4260
         TabIndex        =   2
         Top             =   360
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
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
         Format          =   51642369
         CurrentDate     =   44770
      End
      Begin MSComCtl2.DTPicker xFecIni 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
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
         Format          =   51642369
         CurrentDate     =   44770
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "Comprobante"
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
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "Fecha Final"
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
         Index           =   1
         Left            =   3300
         TabIndex        =   5
         ToolTipText     =   "Se refiere a la Fecha Turno Final"
         Top             =   420
         Visible         =   0   'False
         Width           =   915
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "Fecha Corte"
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
         Index           =   0
         Left            =   180
         TabIndex        =   4
         ToolTipText     =   "Se refiere a la Fecha Turno Inicial"
         Top             =   420
         Width           =   1215
      End
   End
   Begin KewlButtonz.KewlButtons Command2 
      Height          =   435
      Left            =   4140
      TabIndex        =   7
      Top             =   2460
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
      BCOLO           =   16777152
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Comprobantes.frx":0010
      PICN            =   "Comprobantes.frx":002C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons Command1 
      Height          =   435
      Left            =   180
      TabIndex        =   8
      Top             =   2460
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
      MICON           =   "Comprobantes.frx":05C6
      PICN            =   "Comprobantes.frx":05E2
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
Attribute VB_Name = "Comprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xFile As String
Dim xR As New ADODB.Recordset
Dim xSql As String

Private Sub Command1_Click()
Dim xC As New ADODB.Recordset
Dim Data As String, xDes As String, xNotaE As String, xNotaD As String, xDesItem As String
Dim xNumero As Long
Dim xTipo As String
Dim i As Integer

On Error GoTo Recover

xTipo = Mid(Combo1.text, 1, 3)
Select Case xTipo

Case "EIN"
    xSql = "SET DATEFORMAT DMY"
    xSql = xSql + " SELECT IdTiquete, PesoLleno - PesoVacio AS PesoNeto, FechaTurno"
    xSql = xSql + " From Bascula"
    xSql = xSql + " WHERE  Estado='AC' AND IdMaterial=1 AND TransaccionOrigen='LT' AND Procesado=0 AND FechaTurno<='" & xFecIni & "'"
    
    Set xR = Conn.Execute(xSql)
    Set xC = Conn.Execute("Select * From NumerosComprobantes Where Tipo='" & xTipo & "'")
    xNumero = xC!Numero
    i = 1
    xNotaE = "Esta es la NOTA del Encabezado"                           '==> Notas
    xNotaD = "Esta es la NOTA del Detalle"                              '==> Notas
    xDes = "Esta es la Descripcion del movimiento"                      '==> Descripcion
    xDesItem = "Esta es la Descripcion del Item"                        '==> Descripcion
        
    If Not xR.EOF Then
        xFile = sPathComprobantes & "INVENTARIOS.TXT"
        Open xFile For Output As #1
            ' Conector de Inicio
            Data = Format(i, "0000000")                                 '==> Numero
            Data = Data + "000000"                                      '==> Valor fijo
            Data = Data + "01"                                          '==> Version
            Data = Data + "002"                                         '==> Compañia
        Print #1, Data
        
        Do While Not xR.EOF
            ' Encabezado V2
            i = i + 1
            Data = Format(i, "0000000")                                 '==> Numero
            Data = Data + "0450"                                        '==> Valor fijo = 450
            Data = Data + "00"                                          '==> Valor fijo = 00
            Data = Data + "02"                                          '==> Version = 02
            Data = Data + "002"                                         '==> Compañia
            Data = Data + "1"                                           '==> Registro Automatico=1
            Data = Data + "029"                                         '==> Centro de operación
            Data = Data + "EIN"                                         '==> Tipo de documento          ==========================
            Data = Data + Format(i, "########")                         '==> Consecutivo de documento   ==========================
            Data = Data + Format(xR!FechaTurno, "yyyyMMdd")             '==> Numero Fecha YYYYMMDD
            Data = Data + Completa("802022622", 15, " ")                '==> Tercero
            
            Data = Data + "061"                                         '==> Clase Documento
            Data = Data + "1"                                           '==> Estado
            Data = Data + "0"                                           '==> Estado Impresion
            Data = Data + Completa(xNotaE, 255, " ")                    '==> Notas
            Data = Data + "601"                                         '==> Concepto
            Data = Data + Completa("", 5, " ")                          '==> Bodega de salida
            Data = Data + "PL090"                                       '==> Bodega de entrada
            Data = Data + Completa("", 15, " ")                         '==> Documento alterno
            Data = Data + "029"                                         '==> Centro de operación movimiento
            
            Data = Data + Completa("", 3, " ")                          '==> Tipo de documento
            Data = Data + "00000000"                                    '==> Consecutivo de documento
            
            Data = Data + Completa("", 138, " ")                        '==> Vehiculo, Transp, Sucur, Coductor, Nombre, Identif, No guia
            
            Data = Data + "0000000000.0000"                             '==> Cajas
            Data = Data + "000000000000000.0000"                             '==> Peso
            Data = Data + "000000000000000.0000"                             '==> Volumen
            Data = Data + "000000000000000.0000"                             '==> Seguro
            Data = Data + Completa(xNota, 255, " ")                     '==> Notas
            Print #1, Data
            
            ' Movimiento V12
            i = i + 1
            Data = Format(i, "0000000")                                 '==> Numero
            Data = Data + "0470"                                        '==> Valor fijo = 470
            Data = Data + "00"                                          '==> Valor fijo = 00
            Data = Data + "12"                                          '==> Version = 12
            Data = Data + "002"                                         '==> Compañia
            Data = Data + " 33"                                         '==> Centro de Costo
            Data = Data + "EIN"                                         '==> Tipo de Documento
            Data = Data + Format(i - 1, "00000000")                     '==> Numero de Documento
            Data = Data + Format(1, "0000000000")                       '==> Numero Registro del Movimiento ===========================
            Data = Data + Completa("", 55, " ")                         '==> Llenar de espacios
            Data = Data + "PL090"                                       '==> Bodega
            Data = Data + Completa("", 10, " ")                         '==> Ubicaciones
            Data = Data + Completa("", 15, " ")                         '==> Lotes en Bodega
            Data = Data + "601"                                         '==> 601=Entrada; 602=Salida; 603=Ajustes; 607=Transferencias; 610=Procesos; 699=Saldos Iniciales; 605=Transferencia en transito
            Data = Data + "01"                                          '==> código de motivo
            Data = Data + "029"                                         '==> Valida en maestro, código de centro de operación del movimiento
            Data = Data + "  "                                          '==> Llenar de espacios
            Data = Data + Completa("", 15, " ")                         '==> Obligatorio si la  cuenta contable exige ccosto. Valida en maestro, código de centro de costo del movimiento.
            Data = Data + Completa("", 15, " ")                         '==> Valida en maestro, código de proyecto del movimiento
            Data = Data + "KG  "                                        '==> Unidad de medida del movimiento
            Data = Data + Format(xR!PesoNeto, "000000000000000.0000")   '==> Cantidad
            Data = Data + Format(0, "000000000000000.0000")             '==> Cantidad adicional
            Data = Data + Format(0, "000000000000000.0000")             '==> Cantidad adicional
            xNotaD = "Numero Tiquete : " & Format(xR!IdTiquete, "########")
            Data = Data + Completa(xNotaD, 255, " ")                    '==> Notas
            Data = Data + Completa("Descripcion Detalle", 2000, " ")    '==> Descripcion
            Data = Data + Completa("", 40, " ")                         '==> Descripcion del item
            Data = Data + Completa("", 4, " ")                          '==> unidad de inventario del item sea idem al de la base de datos.
            Data = Data + Completa("", 10, " ")                         '==> Ubicacion de entrada maneja ubicaciones
            Data = Data + Completa("", 15, " ")                         '==> Lotes en Bodega de entrada                                                                        '==> Codigo, es obligatorio si no va referencia ni codigo de barras
            Data = Data + "0000049"                                     '==> Codigo del ITEM
            Data = Data + Completa("", 50, " ")                         '==> Codigo, es obligatorio si no va codigo de item ni codigo de barras
            Data = Data + Completa("", 20, " ")                         '==> Codigo, es obligatorio si no va codigo de item ni referencia
            Data = Data + Completa("", 20, " ")                         '==> Es obligatorio si el ítem maneja extensión 1
            Data = Data + Completa("", 20, " ")                         '==> Es obligatorio si el ítem maneja extensión 2
            Data = Data + Completa("21", 20, " ")                       '==> Valida en maestro, código de unidad de negocio del movimiento.
            Data = Data + Format(0, "0000000000")                       '==> Rowid del movimiento base.
            Print #1, Data
            xR.MoveNext
            xNumero = xNumero + 1
        Loop
        i = i + 1
        ' Conector de Final
        Data = Format(i, "0000000")                                 '==> Numero
        Data = Data + "9999"                                        '==> Valor fijo
        Data = Data + "00"                                          '==> Sub Tipo
        Data = Data + "01"                                          '==> Version
        Data = Data + "002"                                         '==> Compañia
        Print #1, Data
        
        Close #1
        xSql = "UPDATE NumerosComprobantes SET Numero=" & xNumero & " WHERE Tipo='" & xTipo & "'"
        Conn.Execute (xSql)
        xSql = "UPDATE Bascula SET Procesado = 1 "
        xSql = xSql + " WHERE  Estado='AC' AND IdMaterial=1 AND TransaccionOrigen='LT' AND Procesado=0 AND FechaTurno<='" & xFecIni & "'"
        
        Conn.Execute (xSql)
        MsgBox "Comprobante Finalizado"
    Else
        MsgBox "NO hay Datos Para Mostrar"
    End If
    xR.Close
    xC.Close
End Select

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error en Creación de Comprobante," & vbCrLf & Err.Description
    MsgBox MSG, , "Comprobantes"
    Err.Clear
    Resume Next
End If

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    xSql = "Select * From NumerosComprobantes"
    Set xR = Conn.Execute(xSql)
    
    If Not xR.EOF Then Combo1.text = xR!Tipo & " " & xR!Descripcion
    
    Do While Not xR.EOF
        Combo1.AddItem xR!Tipo & " " & xR!Descripcion
        xR.MoveNext
    Loop
    xR.Close
    Combo1.ListIndex = 0
    xFecIni = CDate("01/" + Format(Month(Now), "00") + "/" + Format(Year(Now)))
    xFecFin = Now
End Sub
