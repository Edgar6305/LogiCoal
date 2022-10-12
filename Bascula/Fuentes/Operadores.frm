VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Operadores 
   Caption         =   "Operadores - Proveedores"
   ClientHeight    =   6300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12855
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   12855
   StartUpPosition =   1  'CenterOwner
   Begin MBSplit.Splitter VSplitter 
      Height          =   4935
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   8705
      SplitterPos     =   40
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3300
         Top             =   3720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   23
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":0000
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":015A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":02B4
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":040E
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":0568
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":06C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":081C
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":0976
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":0AD0
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":0C2A
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":0D84
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":0EDE
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":18F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":1E8A
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":2424
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":29BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":3508
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":45DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":5730
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":5D46
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":60A4
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":66C7
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Operadores.frx":6CDD
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.TreeView oTree 
         Height          =   3735
         Left            =   0
         TabIndex        =   14
         Top             =   0
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   6588
         _Version        =   393217
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
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
      Begin VB.Frame Frame1 
         Height          =   4635
         Left            =   5640
         TabIndex        =   1
         Top             =   120
         Width           =   6735
         Begin VB.TextBox MemVar_1 
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
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   7
            Top             =   480
            Width           =   1080
         End
         Begin VB.TextBox MemVar_2 
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
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   6
            Top             =   1260
            Width           =   4275
         End
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
            ItemData        =   "Operadores.frx":75CF
            Left            =   1800
            List            =   "Operadores.frx":75D9
            TabIndex        =   5
            Text            =   "OPERADOR"
            Top             =   1980
            Width           =   1575
         End
         Begin VB.TextBox MemVar_4 
            BackColor       =   &H80000018&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1515
            Left            =   1800
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   4
            Text            =   "Operadores.frx":75F2
            Top             =   2400
            Width           =   4215
         End
         Begin MSComCtl2.DTPicker Fecha 
            Height          =   315
            Left            =   1800
            TabIndex        =   2
            Top             =   840
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy hh:mm"
            Format          =   115867651
            CurrentDate     =   44769
         End
         Begin MSMask.MaskEdBox MemVar_3 
            Height          =   315
            Left            =   1800
            TabIndex        =   3
            Top             =   1620
            Width           =   1575
            _ExtentX        =   2778
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
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Tercero"
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
            Left            =   420
            TabIndex        =   13
            Top             =   2040
            Width           =   960
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ID Operador"
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
            Left            =   420
            TabIndex        =   12
            Top             =   540
            Width           =   960
         End
         Begin VB.Label Label11 
            Caption         =   "Descripción"
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
            Left            =   420
            TabIndex        =   11
            Top             =   1320
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "NIT"
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
            Left            =   420
            TabIndex        =   10
            Top             =   1680
            Width           =   255
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
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
            Left            =   420
            TabIndex        =   9
            Top             =   2460
            Width           =   1125
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha "
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
            Left            =   420
            TabIndex        =   8
            Top             =   900
            Width           =   495
         End
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   750
      Left            =   0
      TabIndex        =   15
      Top             =   0
      Width           =   19815
      _ExtentX        =   34951
      _ExtentY        =   1323
      BandCount       =   5
      _CBWidth        =   19815
      _CBHeight       =   750
      _Version        =   "6.0.8169"
      Child1          =   "Tbar"
      MinHeight1      =   330
      Width1          =   2160
      NewRow1         =   0   'False
      Caption2        =   "Estados"
      Child2          =   "Combo3"
      MinHeight2      =   330
      Width2          =   3210
      NewRow2         =   -1  'True
      Child3          =   "oNuevo"
      MinHeight3      =   315
      Width3          =   1500
      NewRow3         =   0   'False
      Child4          =   "Actualizar"
      MinHeight4      =   315
      Width4          =   1500
      NewRow4         =   0   'False
      MinHeight5      =   315
      Width5          =   1305
      NewRow5         =   0   'False
      Begin VB.ComboBox Combo3 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "Operadores.frx":75F7
         Left            =   825
         List            =   "Operadores.frx":7607
         TabIndex        =   19
         Text            =   "TODAS"
         Top             =   390
         Width           =   2355
      End
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   330
         Left            =   165
         TabIndex        =   18
         Top             =   30
         Width           =   19560
         _ExtentX        =   34502
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Grabar"
               Object.ToolTipText     =   "Grabar Registro Actual"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Browse"
               Object.ToolTipText     =   "Explorar Tabla de Datos"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Borrar"
               Object.ToolTipText     =   "Borrar Registro Actual"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Imprimir"
               Object.ToolTipText     =   "Imprimir"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Top"
               Object.ToolTipText     =   "Primer Registro"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Previo"
               Object.ToolTipText     =   "Registro Anterior"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Proximo"
               Object.ToolTipText     =   "Próximo Registro"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bottom"
               Object.ToolTipText     =   "Ultmo Registro"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Foto"
               Object.ToolTipText     =   "Insetar Imagen"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Estado"
               Object.ToolTipText     =   "Salir"
               ImageIndex      =   32
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Salida"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
         EndProperty
      End
      Begin KewlButtonz.KewlButtons oNuevo 
         Height          =   315
         Left            =   3405
         TabIndex        =   17
         Top             =   390
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Nuevo"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Operadores.frx":763C
         PICN            =   "Operadores.frx":7658
         PICH            =   "Operadores.frx":7BF2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons Actualizar 
         Height          =   315
         Left            =   4935
         TabIndex        =   16
         Top             =   390
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Actualizar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Operadores.frx":818C
         PICN            =   "Operadores.frx":81A8
         PICH            =   "Operadores.frx":8742
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
   Begin ComctlLib.StatusBar oBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   20
      Top             =   5985
      Width           =   12855
      _ExtentX        =   22675
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "USUARIO"
            TextSave        =   "USUARIO"
            Key             =   "Usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "FECHA CREACION"
            TextSave        =   "FECHA CREACION"
            Key             =   "Ot"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "ESTADO"
            TextSave        =   "ESTADO"
            Key             =   "Estado"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Operadores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim okUnload As Boolean
Dim IsNewRecord As Boolean
Dim OkOpen As Boolean
Dim oT As New ADODB.Recordset

Private Sub Actualizar_Click()
If MsgBox("Esta Seguro de Actualizar El Operador ", vbYesNo, "Operadores/Proveedores") = vbYes Then
    xSql = "UPDATE OperadoresMineros Set Estado = 'AC' Where IDOperador=" & MemVar_1
    Conn.Execute (xSql)
    Call MemVar_1_LostFocus
End If
End Sub

Private Sub Form_Activate()
If Not OkOpen Then
     Call LoadData
     OkOpen = True
    MemVar_1.SetFocus
End If
Me.SetFocus
Call MuestraArbol
End Sub

Private Sub Form_Resize()
    Dim lSplitHeight As Long
    If Me.WindowState = vbMinimized Then Exit Sub
    'set the position/size of the command buttons and the logo picture
    'set the position/size of the splitter
    lSplitHeight = Me.ScaleHeight - CoolBar1.Height - oBar.Height - 100
    VSplitter.Height = IIf(lSplitHeight < 0, 0, lSplitHeight)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oT.Close
    OkOpen = False
    Unload Me
End Sub

Private Sub Form_Load()

xSql = "Select Top 1 * From OperadoresMineros " & xFilter
oT.Open xSql, Conn, 2, 3, 1

Set VSplitter.LeftOrTopCtl = oTree
Set VSplitter.RightOrBottomCtl = Frame1

End Sub

Private Sub MemVar_2_GotFocus()
    Call Mark(MemVar_2)
End Sub

Private Sub MemVar_2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
                Call omenu("Browse")
        Case vbKeyUp
                MemVar_1.SetFocus
        Case vbKeyDown, vbKeyReturn
                If MemVar_1 = "" And Not IsNewRecord Then
                    MemVar_1.SetFocus
                Else
                    MemVar_3.SetFocus
                End If
End Select
End Sub

Private Sub MemVar_3_GotFocus()
        Call Mark(MemVar_3)
End Sub

Private Sub MemVar_3_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
                Call omenu("Browse")
        Case vbKeyUp
                MemVar_2.SetFocus
        Case vbKeyDown, vbKeyReturn
                If MemVar_2 = "" Then
                    MsgBox "Debe escribir el nombre de Operador/Proveedor, Verifique", vbInformation
                    MemVar_2.SetFocus
                Else
                    MemVar_4.SetFocus
                End If
End Select
End Sub

Private Sub MemVar_3_LostFocus()
Dim xC As New ADODB.Recordset

If MemVar_3 = "" Then
    MsgBox "El NIT del Operador/Proveedor NO puede ser Blancos, Verifique", vbInformation
    MemVar_3.SetFocus
    Exit Sub
End If

If Not Conn.Execute("Select Nit From OperadoresMineros WHERE Nit='" & MemVar_3 & "'").EOF Then
    MsgBox "El NIT ya se encuentra Registrado, Verifique", vbInformation
    MemVar_3.SetFocus
End If

End Sub

Private Sub oNuevo_Click()
If oNuevo.Caption = "Nuevo" Then
    oNuevo.Caption = "Cancelar"
    IsNewRecord = True
    MemVar_1.text = ""
    MemVar_2.text = ""
    MemVar_3.text = ""
    Combo1.ListIndex = 0
    MemVar_7 = "IN"

    oBar.Panels("Usuario").text = "USUARIO: " & Susuario
    oBar.Panels("Ot").text = "FECHA CREACION: " & Now
    oBar.Panels("Estado").text = "ESTADO: " & "IN"
    
    Tbar.buttons("Grabar").Enabled = True
    
    MemVar_2.SetFocus
Else
    oNuevo.Caption = "Nuevo"
    IsNewRecord = False
    oT.MoveLast
    MemVar_1 = oT!IDOperador
    Call MemVar_1_LostFocus
End If

End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
        Call omenu(Button.Key)
End Sub

Private Sub omenu(ByVal Key As String)

On Error GoTo Recover
Select Case Key
    Case "Grabar"
           Call SaveData
           Call MuestraArbol
    Case "Salida"
        Unload Me
    Case "Browse"
            Select Case ActiveControl.Name
            Case "MemVar_1"
                    BrowseCatalogo.x_Tipo = 1
                    BrowseCatalogo.xtabla = "OperadoresMineros"
                    Set BrowseCatalogo.dControl = ActiveControl
                    BrowseCatalogo.Show 1
            Case "MemVar_2"
                    BrowseAcopioPilas.x_Tipo = 1
                    BrowseAcopioPilas.xtabla = "OperadoresMineros"
                    Set BrowseAcopioPilas.dControl = ActiveControl
                    BrowseAcopioPilas.Show 1
            End Select
            
    Case "Imprime"
    
    Case "Borrar"
    
    Case "Top"
        oT.Close
        oT.Open "Select Top 1 * From OperadoresMineros Order By IdOperador"
        Call LoadData
    Case "Bottom"
        oT.Close
        oT.Open "Select Top 1 * From OperadoresMineros Order By IdOperador DESC"
        Call LoadData
    Case "Proximo"
        oT.Close
        oT.Open "Select Top 1 * From OperadoresMineros Where IdOperador>'" & MemVar_1 & "' Order By  IdOperador"
        Call LoadData
    Case "Previo"
        oT.Close
        oT.Open "Select Top 1 * From OperadoresMineros Where IdOperador<'" & MemVar_1 & "' Order By IdOperador DESC"
        Call LoadData
End Select
        
Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error en Rutina Browse," & vbCrLf & Err.Description
    MsgBox MSG, , "Browse"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
        
End Sub

Private Sub LoadData()
On Error GoTo Recover

 If oT.EOF And oT.BOF Then Exit Sub
 
 If Not IsNewRecord Then
     MemVar_1 = oT!IDOperador
     MemVar_2 = oT!Descripcion
     MemVar_3 = oT!Nit
     Fecha = oT!Fecha
     MemVar_4 = oT!Observaciones
     Combo1.ListIndex = IIf(oT!Tipo = "OPERADOR", 0, 1)
 Else
 
 End If
    
 oBar.Panels("Usuario").text = "USUARIO: " & oT!Usuario
 oBar.Panels("Ot").text = "FECHA CREACION: " & oT!Fecha
 oBar.Panels("Estado").text = "ESTADO: " & oT!Estado

 Actualizar.Enabled = (oT!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','EstadoOperador'").Fields(0)) '==> Fields(0) Nos dice si tiene acceso al Objeto

 Tbar.buttons("Grabar").Enabled = (oT!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','" & sProgSISMA & "'").Fields(1)) '==> Fields(1) Nos dice puede hacer cambios al Objeto
 Tbar.buttons("Borrar").Enabled = (oT!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','" & sProgSISMA & "'").Fields(1)) '==> Fields(1) Nos dice puede hacer cambios al Objeto

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer Transaccion," & vbCrLf & Err.Description
    MsgBox MSG, , "LoadData"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
    
End Sub

Private Sub SaveData()
Dim xSql As String
Dim Res As String
Dim xR As New ADODB.Recordset

On Error GoTo Recover

'CREATE PROCEDURE [dbo].[PA_Operadores]
'@Des As Int,
'@Nit AS Int,
'@Tipo AS Varchar(1),
'@Observaciones As Varchar(MAX),
'@Usuario As Varchar(10)

 If IsNewRecord Then
     xSql = "EXEC PA_Operadores '" & MemVar_2 & "','" & MemVar_3 & "','" & Combo1.text & "','" & MemVar_4 & "','" & Susuario & "'"
 
    Set xR = Conn.Execute(xSql)
    Res = xR!Res
    
    If Res <> "OK" Then
       MsgBox "Error al Grabar Nota , Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
       Exit Sub
    End If
 Else
     xSql = "UPDATE OperadoresMineros Set  Descripcion = " & MemVar_2 & "," & Tipo = Combo1 & "," & Observaciones = MemVar_4 & "," & Nit = MemVar_3 & " WHERE IdOperador=" & MemVar_1
 End If

 If IsNewRecord Then
    MemVar_1 = xR!Numero
    Call MuestraArbol
    Call MemVar_1_LostFocus
 End If
 oNuevo.Caption = "Nuevo"
 IsNewRecord = False
 xR.Close
 
Exit Sub

Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Salvar los Datos," & vbCrLf & Err.Description
    MsgBox MSG, , "Savedata()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub

Private Sub MuestraArbol()
Dim Em As New ADODB.Recordset
Dim xSql As String, iCono As Integer

'On Error GoTo Recover

oTree.Nodes.Clear
Set Nodx = oTree.Nodes.Add(, , "0" & "Operadores", "Operadores", 18)

xSql = "SELECT DISTINCT Tipo FROM  OperadoresMineros"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    If Em!Tipo = "OPERADOR" Then
        iCono = 22
    Else
        iCono = 20
    End If
    Set Nodx = oTree.Nodes.Add("0" & "Operadores", tvwChild, "A" & Em!Tipo, Em!Tipo, iCono)
    Em.MoveNext
Wend

xSql = "SELECT * FROM  OperadoresMineros"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("A" & Em!Tipo, tvwChild, "B" & Str(Em!IDOperador), Em!Descripcion, 23)
    Em.MoveNext
Wend
  
For Each loNode In oTree.Nodes
    If loNode.children >= 1 Then
        loNode.Expanded = True
    End If
Next
    
Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error en insercion de nodo," & vbCrLf & Err.Description
    MsgBox MSG, , "Muestra Arbol"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub

Private Sub MemVar_1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
                Call omenu("Browse")
        Case vbKeyUp
        Case vbKeyDown, vbKeyReturn
                MemVar_2.SetFocus
End Select

End Sub

Private Sub MemVar_1_GotFocus()
        Call Mark(MemVar_1)
End Sub

Private Sub MemVar_1_LostFocus()
Dim xSql As String

On Error GoTo Recover

If MemVar_1 <> "" Then
    If Not IsNewRecord Then
         If Not MySeek(oT, Conn, "OperadoresMineros", "IdOperador='" & MemVar_1 & "'") Then
                Call LoadData
         Else
                MsgBox "Operador/PeoveedorNO Registrado, Verifique"
                MemVar_1.SetFocus
        End If
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga La Nota" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_1_LostFocus()"
    Err.Clear
    Resume Next
End If

End Sub

Private Sub oTree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer

Select Case Mid(Node.Key, 1, 1)
Case "B"
    MemVar_1 = Mid(Node.Key, 2, 6)
    Call MemVar_1_LostFocus
End Select

End Sub

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
    dControl.SelStart = 0
    dControl.SelLength = Len(dControl.text)
End If
End Sub




