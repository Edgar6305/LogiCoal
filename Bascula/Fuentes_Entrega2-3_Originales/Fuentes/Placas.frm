VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Placas 
   Caption         =   "Registro de Placas"
   ClientHeight    =   5880
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11595
   LinkTopic       =   "Form1"
   ScaleHeight     =   5880
   ScaleWidth      =   11595
   StartUpPosition =   1  'CenterOwner
   Begin MBSplit.Splitter VSplitter 
      Height          =   4455
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   7858
      SplitterPos     =   35
      Begin VB.Frame Frame1 
         Height          =   3975
         Left            =   4860
         TabIndex        =   2
         Top             =   120
         Width           =   6195
         Begin MSMask.MaskEdBox PesoVacioMAX 
            Height          =   315
            Left            =   1800
            TabIndex        =   22
            Top             =   2100
            Width           =   1035
            _ExtentX        =   1826
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
         Begin VB.TextBox MemVar_3 
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
            TabIndex        =   16
            Top             =   1320
            Width           =   3495
         End
         Begin VB.CheckBox MemVar_5 
            Alignment       =   1  'Right Justify
            Caption         =   "Estado"
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
            TabIndex        =   15
            Top             =   2940
            Width           =   1575
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
            TabIndex        =   7
            Top             =   960
            Width           =   1455
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
            Picture         =   "Placas.frx":0000
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   6
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
                     Picture         =   "Placas.frx":0102
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Placas.frx":041C
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Placas.frx":0736
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Placas.frx":0A50
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Placas.frx":0D6A
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Placas.frx":1084
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Placas.frx":139E
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Placas.frx":16B8
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Placas.frx":19D2
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Placas.frx":1CEC
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Placas.frx":2006
                     Key             =   ""
                  EndProperty
               EndProperty
            End
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
            Picture         =   "Placas.frx":2320
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   5
            Top             =   600
            Visible         =   0   'False
            Width           =   240
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
            TabIndex        =   4
            Top             =   600
            Width           =   1080
         End
         Begin VB.TextBox MemVar_4 
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
            TabIndex        =   3
            ToolTipText     =   "Pulse ""F4"" para Buscar el Proveedor"
            Top             =   1680
            Width           =   660
         End
         Begin MSMask.MaskEdBox PesoLlenoMAX 
            Height          =   315
            Left            =   1800
            TabIndex        =   26
            Top             =   2460
            Width           =   1035
            _ExtentX        =   1826
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
         Begin VB.Label LabelConductor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3360
            TabIndex        =   28
            Top             =   960
            Width           =   2595
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Kilos"
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
            Index           =   5
            Left            =   3000
            TabIndex        =   27
            Top             =   2520
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Kilos"
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
            Index           =   3
            Left            =   3000
            TabIndex        =   23
            Top             =   2160
            Width           =   345
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Peso Lleno MAX"
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
            Index           =   2
            Left            =   420
            TabIndex        =   21
            Top             =   2520
            Width           =   1185
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Peso Vacio MIN"
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
            Index           =   1
            Left            =   420
            TabIndex        =   20
            Top             =   2160
            Width           =   1125
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Vehiculo"
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
            TabIndex        =   17
            Top             =   1380
            Width           =   975
         End
         Begin VB.Label Label11 
            Caption         =   "Conductor"
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
            TabIndex        =   11
            Top             =   1020
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Placas"
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
            TabIndex        =   10
            Top             =   660
            Width           =   480
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Transportador"
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
            TabIndex        =   9
            Top             =   1740
            Width           =   1035
         End
         Begin VB.Label LabelTransportador 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2520
            TabIndex        =   8
            Top             =   1680
            Width           =   3435
         End
      End
      Begin MSComctlLib.TreeView oTree 
         Height          =   3735
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   2175
         _ExtentX        =   3836
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
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11040
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   15
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":2422
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":257C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":26D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":2830
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":298A
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":2AE4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":2C3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":2D98
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":2EF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":304C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":31A6
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":3300
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":3D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":42AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Placas.frx":4846
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   750
      Left            =   0
      TabIndex        =   12
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
      Begin KewlButtonz.KewlButtons oNuevo 
         Height          =   315
         Left            =   3405
         TabIndex        =   18
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
         MICON           =   "Placas.frx":4DE0
         PICN            =   "Placas.frx":4DFC
         PICH            =   "Placas.frx":5396
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   330
         Left            =   165
         TabIndex        =   14
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
               Object.ToolTipText     =   "Pr�ximo Registro"
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
         ItemData        =   "Placas.frx":5930
         Left            =   825
         List            =   "Placas.frx":5940
         TabIndex        =   13
         Text            =   "TODAS"
         Top             =   390
         Width           =   2355
      End
   End
   Begin ComctlLib.StatusBar oBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   19
      Top             =   5565
      Width           =   11595
      _ExtentX        =   20452
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
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   315
      Left            =   0
      TabIndex        =   24
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
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
      Caption         =   "Kilos"
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
      Index           =   4
      Left            =   1200
      TabIndex        =   25
      Top             =   60
      Width           =   345
   End
End
Attribute VB_Name = "Placas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Placas As New ADODB.Recordset
Dim okUnload As Boolean
Dim IsNewRecord As Boolean
Dim OkOpen As Boolean
Dim oT As New ADODB.Recordset

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
On Error Resume Next
    Placas.Close
    MenuNuevo.Flag_xProg = ""
    Unload Me
End Sub

Private Sub Form_Load()
Dim xCt As New ADODB.Recordset

xSql = "Select Top 1 * From Placas " & xFilter & " Order By Placas DESC"
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
                MemVar_3.SetFocus
End Select
End Sub

Private Sub MemVar_2_LostFocus()
Dim xC As New ADODB.Recordset

If MemVar_2 <> "" Then
    Set xC = Conn.Execute("Select Nombre From Conductores Where Cedula='" & MemVar_2 & "'")
    If Not xC.EOF Then
       LabelConductor.Caption = xC!Nombre
    Else
        MsgBox "Conductor NO registrado, Verifique", vbInformation
        MemVar_2.SetFocus
    End If
End If
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
                MemVar_4.SetFocus
End Select
End Sub

Private Sub MemVar_4_GotFocus()
        Call Mark(MemVar_4)
End Sub

Private Sub MemVar_4_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
                Call omenu("Browse")
        Case vbKeyUp
                MemVar_3.SetFocus
        Case vbKeyDown, vbKeyReturn
                PesoVacioMAX.SetFocus
End Select
End Sub

Private Sub MemVar_4_LostFocus()
Dim xR As New ADODB.Recordset

If MemVar_4 <> "" Then
    Set xR = Conn.Execute("Select * From Transportador Where IdTransportador=" & MemVar_4)
    If Not xR.EOF Then
        LabelTransportador = xR!Descripcion
    Else
        LabelTransportador = ""
        MemVar_4.SetFocus
    End If
Else
End If
End Sub

Private Sub oNuevo_Click()
If oNuevo.Caption = "Nuevo" Then
    oNuevo.Caption = "Cancelar"
    IsNewRecord = True
    MemVar_1.text = ""
    MemVar_2.text = ""
    MemVar_3.text = ""
    MemVar_4.text = ""
    MemVar_5 = 1
    Fecha = Now
Else
    oNuevo.Caption = "Nuevo"
    IsNewRecord = False
End If
MemVar_1.SetFocus
End Sub

Private Sub PesoVacioMAX_GotFocus()
        Call Mark(PesoVacioMAX)
End Sub

Private Sub PesoVacioMAX_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
                Call omenu("Browse")
        Case vbKeyUp
                MemVar_4.SetFocus
        Case vbKeyDown, vbKeyReturn
                PesoLlenoMAX.SetFocus
End Select
End Sub

Private Sub PesoLlenoMAX_GotFocus()
        Call Mark(PesoLlenoMAX)
End Sub

Private Sub PesoLlenoMAX_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
                Call omenu("Browse")
        Case vbKeyUp
                PesoVacioMAX.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_5.SetFocus
End Select
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
                                BrowseCatalogo.xtabla = "PLACAS"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        Case "MemVar_4"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "Transportador"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        End Select
                Case "Imprime"

                Case "Borrar"
                
                Case "Top"
                    oT.Close
                    oT.Open "Select Top 1 * From Placas Order By Placas"
                    Call LoadData
                Case "Bottom"
                    oT.Close
                    oT.Open "Select Top 1 * From Placas Order By Placas DESC"
                    Call LoadData
                Case "Proximo"
                    oT.Close
                    oT.Open "Select Top 1 * From Placas Where Placas>'" & MemVar_1 & "' Order By  Placas"
                    Call LoadData
                Case "Previo"
                    oT.Close
                    oT.Open "Select Top 1 * From Placas Where Placas<'" & MemVar_1 & "' Order By Placas DESC"
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
         okNew.Visible = True
         okFind.Visible = False
         Exit Sub
 End If

 If Not IsNewRecord Then
     okNew.Visible = False
     okFind.Visible = True

     MemVar_1 = oT!Placas
     MemVar_2 = oT!Conductor
     MemVar_3 = oT!TipoVehiculo
     MemVar_4 = oT!IdTransportador
     PesoVacioMAX = oT!PesoVacioMAX
     PesoLlenoMAX = oT!PesoLlenoMAX
     LabelTransportador.Caption = Conn.Execute("Select Descripcion From Transportador Where IdTransportador=" & MemVar_4).Fields(0)
     MemVar_5 = IIf(oT!Estado, 1, 0)
 Else
 
 End If

 oBar.Panels("Usuario").text = "USUARIO: " & oT!Usuario
 oBar.Panels("Ot").text = "FECHA CREACION: " & oT!Fecha

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
        If MemVar_1.text = "" Then ok = True
        If MemVar_2.text = "" Then ok = True
        If ok Then Exit Sub
        
        If IsNewRecord Then
            oT.AddNew
            oT!Placas = MemVar_1.text
            oT!Conductor = MemVar_2.text
            oT!TipoVehiculo = MemVar_3.text
            oT!IdTransportador = MemVar_4.text
            oT!Fecha = Now
            oT!Usuario = Susuario
            oT!Estado = 1
            oT!PesoVacioMAX = PesoVacioMAX
            oT!PesoLlenoMAX = PesoLlenoMAX
            oT!Tara = 0
        Else
            oT!Conductor = MemVar_2.text
            oT!TipoVehiculo = MemVar_3.text
            oT!IdTransportador = MemVar_4.text
            oT!Fecha = Now
            oT!Usuario = Susuario
            oT!Estado = IIf(MemVar_5, 1, 0)
            oT!PesoVacioMAX = PesoVacioMAX
            oT!PesoLlenoMAX = PesoLlenoMAX
        End If
        
        oT.Update
        
        If IsNewRecord Then
            If Not MySeek(oT, Conn, "Placas", "Placas=" & MemVar_1) Then
                    ' call LoadData
                    okFind.Visible = True
                    okNew.Visible = False
            End If
        End If
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
Set Nodx = oTree.Nodes.Add(, , "0" & "Transportador", "Transportador", 13)

xSql = "SELECT  Distinct Descripcion  FROM  vPlacasTransportador"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("0" & "Transportador", tvwChild, "A" & Em!Descripcion, Em!Descripcion, 14)
    Em.MoveNext
Wend

xSql = "SELECT  *   FROM  vPlacasTransportador"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("A" & Em!Descripcion, tvwChild, "B" & Em!Placas, Em!Placas & " " & Em!TipoVehiculo & " " & Em!Conductor, 14)
    Em.MoveNext
Wend
    
For Each loNode In oTree.Nodes
    If loNode.children = 1 Then
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

MemVar_1 = UCase(MemVar_1)
If Not IsNewRecord Then
     If Not MySeek(oT, Conn, "Placas", "Placas='" & MemVar_1 & "'") Then
            Call LoadData
     Else
            MsgBox "Placas NO Registrada, Verifique"
            MemVar_1.SetFocus
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Placas" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_1_LostFocus()"
    Err.Clear
    Resume Next
End If

End Sub

Private Sub oTree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer

Select Case Mid(Node.Key, 1, 1)
Case "A"
Case "B"
    MemVar_1 = Mid(Node.Key, 2, 8)
    Call MemVar_1_LostFocus
End Select

End Sub

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
                dControl.SelStart = 0
                dControl.SelLength = Len(dControl.text)
End If
End Sub

