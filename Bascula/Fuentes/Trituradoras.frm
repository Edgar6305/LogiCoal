VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Trituradoras 
   Caption         =   "Trituradoras"
   ClientHeight    =   4065
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10305
   LinkTopic       =   "Form1"
   ScaleHeight     =   4065
   ScaleWidth      =   10305
   StartUpPosition =   1  'CenterOwner
   Begin MBSplit.Splitter VSplitter 
      Height          =   2835
      Left            =   120
      TabIndex        =   0
      Top             =   780
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   5001
      SplitterPos     =   35
      Begin MSComctlLib.TreeView oTree 
         Height          =   3735
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   6588
         _Version        =   393217
         Style           =   7
         ImageList       =   "ImageList1"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Frame Frame1 
         Height          =   2475
         Left            =   3660
         TabIndex        =   1
         Top             =   180
         Width           =   6195
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
            Left            =   1800
            TabIndex        =   16
            Top             =   1320
            Width           =   3495
         End
         Begin MSMask.MaskEdBox MemVar_3 
            Height          =   315
            Left            =   1800
            TabIndex        =   14
            Top             =   1680
            Width           =   1095
            _ExtentX        =   1931
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
         Begin VB.TextBox MemVar_1 
            BeginProperty Font 
               Name            =   "Arial"
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
            TabIndex        =   5
            Top             =   600
            Width           =   1080
         End
         Begin VB.PictureBox okFind 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3060
            Picture         =   "Trituradoras.frx":0000
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   4
            Top             =   600
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.PictureBox okNew 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3060
            Picture         =   "Trituradoras.frx":0102
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   3
            Top             =   600
            Visible         =   0   'False
            Width           =   240
            Begin ComctlLib.ImageList ImageList3 
               Left            =   10040
               Top             =   500
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   32
               ImageHeight     =   32
               MaskColor       =   -2147483633
               _Version        =   327682
               BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
                  NumListImages   =   11
                  BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Trituradoras.frx":0204
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Trituradoras.frx":051E
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Trituradoras.frx":0838
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Trituradoras.frx":0B52
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Trituradoras.frx":0E6C
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Trituradoras.frx":1186
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Trituradoras.frx":14A0
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Trituradoras.frx":17BA
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Trituradoras.frx":1AD4
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Trituradoras.frx":1DEE
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Trituradoras.frx":2108
                     Key             =   ""
                  EndProperty
               EndProperty
            End
         End
         Begin VB.TextBox MemVar_2 
            BeginProperty Font 
               Name            =   "Arial"
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
            TabIndex        =   2
            Top             =   960
            Width           =   3495
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   5160
            Top             =   1740
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   20
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":2422
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":257C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":26D6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":2830
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":298A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":2AE4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":2C3E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":2D98
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":2EF2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":304C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":31A6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":3300
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":3D12
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":42AC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":4846
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":4DE0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":592A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":69FC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":7B52
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Trituradoras.frx":8168
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Toneladas Hora"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   2940
            TabIndex        =   18
            Top             =   1740
            Width           =   1140
         End
         Begin VB.Label Label2 
            Caption         =   "Acopio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   15
            Top             =   1380
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Id Trituradora"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   420
            TabIndex        =   8
            Top             =   660
            Width           =   960
         End
         Begin VB.Label Label11 
            Caption         =   "Descripción"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   420
            TabIndex        =   7
            Top             =   1020
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Producción Hora"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   420
            TabIndex        =   6
            Top             =   1740
            Width           =   1200
         End
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   750
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   19815
      _ExtentX        =   34951
      _ExtentY        =   1323
      BandCount       =   4
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
      MinHeight4      =   315
      Width4          =   1695
      NewRow4         =   0   'False
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   330
         Left            =   165
         TabIndex        =   17
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
         ItemData        =   "Trituradoras.frx":8C4A
         Left            =   825
         List            =   "Trituradoras.frx":8C5A
         TabIndex        =   12
         Text            =   "TODAS"
         Top             =   390
         Width           =   2355
      End
      Begin KewlButtonz.KewlButtons oNuevo 
         Height          =   315
         Left            =   3405
         TabIndex        =   11
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
         MICON           =   "Trituradoras.frx":8C8F
         PICN            =   "Trituradoras.frx":8CAB
         PICH            =   "Trituradoras.frx":9245
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
      TabIndex        =   13
      Top             =   3750
      Width           =   10305
      _ExtentX        =   18177
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
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
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Trituradoras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim okUnload As Boolean
Dim IsNewRecord As Boolean
Dim OkOpen As Boolean
Dim oT As New ADODB.Recordset

Private Sub Form_Activate()
If Not OkOpen Then
     Call LoadData
     OkOpen = True
End If
Me.SetFocus
Call MuestraArbol
MemVar_1.SetFocus
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
On Error Resume Next
    oT.Close
    MenuNuevo.Flag_xProg = ""
    Unload Me
End Sub

Private Sub Form_Load()
Dim xR As New ADODB.Recordset
Dim xSql As String

xSql = "Select * From Acopios "
Set xR = Conn.Execute(xSql)

If Not xR.EOF Then
    Combo1.text = Format(xR!IdAcopio, "00") + " " + xR!Descripcion + " (" + xR!Ubicacion + ")"
    Do While Not xR.EOF
        Combo1.AddItem Format(xR!IdAcopio, "00") + " " + xR!Descripcion + " (" + xR!Ubicacion + ")"
        xR.MoveNext
    Loop
Else
    MsgBox "NO Hay Acopios Definidos, Verifique"
    Exit Sub
End If

xSql = "Select Top 1 * From Trituradoras"
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
                MemVar_1.SetFocus
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
                MemVar_1.SetFocus
End Select
End Sub

Private Sub oNuevo_Click()
If oNuevo.Caption = "Nuevo" Then
    IsNewRecord = True
    oNuevo.Caption = "Cancelar"
    MemVar_1.text = ""
    MemVar_2.text = ""
    MemVar_3.text = ""
    Fecha = Now

    oBar.Panels("Usuario").text = "USUARIO: " & Susuario
    oBar.Panels("Ot").text = "FECHA CREACION: " & Format(Now, "dd/MM/yyyy hh:mm")
    
    Tbar.buttons("Grabar").Enabled = True
    
Else
    oNuevo.Caption = "Nuevo"
    IsNewRecord = False
    Exit Sub
End If
MemVar_2.SetFocus
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
        Call omenu(Button.Key)
End Sub

Private Sub omenu(ByVal Key As String)

On Error GoTo Recover
        Select Case Key
                Case "Grabar"
                       Call SaveData
                Case "Salida"
                    Unload Me
                Case "Browse"
                        Select Case ActiveControl.Name
                        Case "MemVar_1"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "Select Identificacion, Descripcion From Terceros"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        End Select
                Case "Imprime"

                Case "Borrar"
                    If MsgBox("Esta seguro de Borrar el Trituradoras", vbYesNo, "Borrado de Registro") = vbYes Then
                        If Conn.Execute(" Select Count(*) From Trituracion Where IdTrituradora=" & MemVar_1).Fields(0) = 0 Then
                            Conn.Execute ("Delete Trituradoras  Where IdTrituradora=" & MemVar_1)
                            oT.Close
                            oT.Open "Select Top 1 * From Trituradoras"
                            MemVar_1 = oT!IdTrituradora
                            Call MemVar_1_LostFocus
                        Else
                            MsgBox "NO puede borrar la Trituradora porque presenta Transacciones realizadas", vbCritical, "LogyNext"
                            MemVar_1.SetFocus
                            Exit Sub
                        End If
                    End If
                
                Case "Top"
                    oT.Close
                    oT.Open "Select Top 1 * From Trituradoras Order By IdTrituradora"
                    Call LoadData
                Case "Bottom"
                    oT.Close
                    oT.Open "Select Top 1 * From Trituradoras Order By IdTrituradora DESC"
                    Call LoadData
                Case "Proximo"
                    oT.Close
                    oT.Open "Select Top 1 * From Trituradoras Where IdTrituradora>'" & MemVar_1 & "' Order By  IdTrituradora"
                    Call LoadData
                Case "Previo"
                    oT.Close
                    oT.Open "Select Top 1 * From Trituradoras Where IdTrituradora<'" & MemVar_1 & "' Order By IdTrituradora DESC"
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

 If oT.EOF And oT.BOF Then
         Exit Sub
 End If

 If Not IsNewRecord Then
     okNew.Visible = False
     okFind.Visible = True
     MemVar_1 = oT!IdTrituradora
     MemVar_2 = oT!Descripcion
     MemVar_3 = oT!ProduccionHora
     Combo1.ListIndex = oT!Acopio - 1
 Else
 
 End If

 oBar.Panels("Usuario").text = "USUARIO: " & oT!Usuario
 oBar.Panels("Ot").text = "FECHA CREACION: " & Format(oT!Fecha, "dd/MM/yyyy hh:mm")
 Call MuestraArbol
 
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
Dim ok As Boolean

On Error GoTo Recover

ok = False
If IsNewRecord = False And MemVar_1.text = "" Then ok = True
If MemVar_2.text = "" Then ok = True
If ok Then Exit Sub

'ALTER PROCEDURE [dbo].[PA_Trituradoras]
'@Descripcion As Varchar(55),
'@ProdHoras float,
'@Acopio int,
'@Usuario As Varchar(10)

 If IsNewRecord Then
     xSql = "EXEC PA_Trituradoras '" & MemVar_2 & "'," & MemVar_3 & "," & Val(Mid(Combo1.text, 1, 2)) & ",'" & Susuario & "'"
 
    Set xR = Conn.Execute(xSql)
    Res = xR!Res
    
    If Res <> "OK" Then
       MsgBox "Error al Grabar Trituradora , Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
       Exit Sub
    End If
 Else
     xSql = "UPDATE Trituradoras Set  Descripcion = '" & MemVar_2 & "',Acopio ='" & Val(Mid(Combo1, 1, 2)) & "',ProduccionHora = " & MemVar_3 & " WHERE IdTrituradora=" & MemVar_1
     Conn.Execute (xSql)
 End If

 If IsNewRecord Then
    MemVar_1 = xR!Numero
 End If

Call MuestraArbol
Call MemVar_1_LostFocus

IsNewRecord = False
oNuevo.Caption = "Nuevo"

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
Dim xSql As String

On Error GoTo Recover

oTree.Nodes.Clear
Set Nodx = oTree.Nodes.Add(, , "0" & "Trituradoras", "TRITURADORAS", 18)

xSql = "SELECT  DISTINCT Acopios.Descripcion, Acopios.Ubicacion"
xSql = xSql + " FROM   Acopios INNER JOIN Trituradoras ON Acopios.IdAcopio = Trituradoras.Acopio"
Set Em = Conn.Execute(xSql)
While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("0" & "Trituradoras", tvwChild, "A" & Em!Ubicacion, Em!Ubicacion, 19)
    Em.MoveNext
Wend

xSql = "SELECT  DISTINCT Acopios.Descripcion, Acopios.Ubicacion"
xSql = xSql + " FROM   Acopios INNER JOIN Trituradoras ON Acopios.IdAcopio = Trituradoras.Acopio"
Set Em = Conn.Execute(xSql)
While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("A" & Em!Ubicacion, tvwChild, "B" & Em!Ubicacion & Em!Descripcion, Em!Descripcion, 17)
    Em.MoveNext
Wend

xSql = "SELECT Acopios.IdAcopio, Acopios.Descripcion, Acopios.Ubicacion, Trituradoras.IdTrituradora, Trituradoras.Descripcion AS DesTrituradora"
xSql = xSql + " FROM   Acopios INNER JOIN Trituradoras ON Acopios.IdAcopio = Trituradoras.Acopio"
Set Em = Conn.Execute(xSql)
While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("B" & Em!Ubicacion & Em!Descripcion, tvwChild, "C" & Format(Em!IdTrituradora, "00"), Em!DesTrituradora, 20)
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
                MemVar_1 = ""
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
     If Not MySeek(oT, Conn, "Trituradoras", "IdTrituradora=" & MemVar_1) Then
            Call LoadData
            If IsNewRecord Then
                MsgBox "ID Trituradora Se encuentra en la base de datos, Verifique", vbInformation
                MemVar_1.SetFocus
                Exit Sub
            End If
     Else
            If Not IsNewRecord Then
                MsgBox "Trituradora NO Registrado, Verifique"
                MemVar_1.SetFocus
            End If
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error  Trituradoras" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_1_LostFocus()"
    Err.Clear
    Resume Next
End If

End Sub

Private Sub oTree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer

Select Case Mid(Node.Key, 1, 1)
Case "A"
Case "C"
    MemVar_1 = Val(Mid(Node.Key, 2, 2))
    Call MemVar_1_LostFocus
End Select

End Sub

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
                dControl.SelStart = 0
                dControl.SelLength = Len(dControl.text)
End If
End Sub

