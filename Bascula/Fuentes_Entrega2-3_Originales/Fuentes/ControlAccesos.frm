VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form ControlAccesos 
   Caption         =   "Control d Acceso"
   ClientHeight    =   7650
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7755
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "ControlAccesos.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7650
   ScaleWidth      =   7755
   StartUpPosition =   1  'CenterOwner
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
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   1260
      MaxLength       =   30
      TabIndex        =   5
      Top             =   1350
      Width           =   3480
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
      Height          =   330
      Left            =   1260
      MaxLength       =   15
      TabIndex        =   4
      Top             =   1020
      Width           =   1560
   End
   Begin VB.PictureBox okFind 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2940
      Picture         =   "ControlAccesos.frx":058A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   3
      Top             =   1020
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox okNew 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2940
      Picture         =   "ControlAccesos.frx":068C
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   1020
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":078E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":08E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":0A42
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":0B9C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":0CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":0E50
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":0FAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":1104
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":125E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":13B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":1512
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":166C
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":207E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":4171
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":42CB
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":4425
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":457F
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":46D9
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":4833
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":498D
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":4AE7
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":4C41
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":4D9B
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":4EF5
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":504F
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":5A61
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin Threed.SSFrame SSFrame1 
      Height          =   1155
      Left            =   4860
      TabIndex        =   0
      Top             =   840
      Visible         =   0   'False
      Width           =   2775
      _Version        =   65536
      _ExtentX        =   4895
      _ExtentY        =   2037
      _StockProps     =   14
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Begin VB.ComboBox Combo1 
         BackColor       =   &H00E1FFFF&
         Height          =   315
         Left            =   120
         TabIndex        =   26
         Text            =   "Combo1"
         Top             =   240
         Width           =   2595
      End
      Begin KewlButtonz.KewlButtons KewlButtons1 
         Height          =   345
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "Grabar"
         ENAB            =   0   'False
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "ControlAccesos.frx":7B54
         PICN            =   "ControlAccesos.frx":7B70
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
   Begin MSMask.MaskEdBox MemVar_3 
      Height          =   285
      Left            =   1260
      TabIndex        =   6
      Top             =   1710
      Visible         =   0   'False
      Width           =   300
      _ExtentX        =   529
      _ExtentY        =   503
      _Version        =   393216
      MaxLength       =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin Threed.SSFrame oMarco 
      Height          =   5415
      Left            =   120
      TabIndex        =   7
      Top             =   2100
      Width           =   7515
      _Version        =   65536
      _ExtentX        =   13256
      _ExtentY        =   9551
      _StockProps     =   14
      ForeColor       =   8421504
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Font3D          =   1
      ShadowStyle     =   1
      Enabled         =   -1  'True
      Begin VB.TextBox MemVar_4 
         Height          =   285
         Index           =   0
         Left            =   45
         MaxLength       =   25
         TabIndex        =   9
         Top             =   420
         Width           =   2220
      End
      Begin VB.VScrollBar oMove 
         Height          =   5160
         Left            =   7200
         TabIndex        =   8
         Top             =   180
         Width           =   255
      End
      Begin Threed.SSCheck MemVar_5 
         Height          =   255
         Index           =   0
         Left            =   6000
         TabIndex        =   10
         Top             =   420
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Threed.SSCheck MemVar_6 
         Height          =   255
         Index           =   0
         Left            =   6660
         TabIndex        =   11
         Top             =   420
         Width           =   255
         _Version        =   65536
         _ExtentX        =   450
         _ExtentY        =   450
         _StockProps     =   78
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.01
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Descripción"
         Height          =   300
         Index           =   1
         Left            =   2280
         TabIndex        =   16
         Top             =   120
         Width           =   3540
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Programa"
         Height          =   300
         Index           =   0
         Left            =   45
         TabIndex        =   15
         Top             =   120
         Width           =   2220
      End
      Begin VB.Label SubLabel4 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   2265
         TabIndex        =   14
         Top             =   420
         Width           =   3540
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Acceso"
         Height          =   300
         Left            =   5805
         TabIndex        =   13
         Top             =   120
         Width           =   720
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         Height          =   300
         Left            =   6540
         TabIndex        =   12
         Top             =   120
         Width           =   600
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   780
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   12435
      _ExtentX        =   21934
      _ExtentY        =   1376
      BandCount       =   5
      _CBWidth        =   12435
      _CBHeight       =   780
      _Version        =   "6.0.8169"
      Child1          =   "Tbar"
      MinHeight1      =   330
      Width1          =   4275
      NewRow1         =   0   'False
      Child2          =   "Combo2"
      MinHeight2      =   315
      Width2          =   2565
      NewRow2         =   -1  'True
      Child3          =   "SSCommand1"
      MinHeight3      =   315
      Width3          =   2100
      NewRow3         =   0   'False
      Child4          =   "CopiarPerfil"
      MinHeight4      =   315
      Width4          =   1995
      NewRow4         =   0   'False
      MinHeight5      =   360
      Width5          =   1200
      NewRow5         =   0   'False
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   330
         Left            =   165
         TabIndex        =   21
         Top             =   30
         Width           =   12180
         _ExtentX        =   21484
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
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
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   1
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "Imprimir"
                  EndProperty
               EndProperty
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
               Object.Visible         =   0   'False
               Key             =   "Foto"
               Object.ToolTipText     =   "Insetar Imagen"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Excel"
               Object.ToolTipText     =   "Salir"
               ImageIndex      =   13
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Salida"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
      Begin KewlButtonz.KewlButtons CopiarPerfil 
         Height          =   315
         Left            =   4890
         TabIndex        =   20
         Top             =   405
         Visible         =   0   'False
         Width           =   1800
         _ExtentX        =   3175
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Copiar Perfil"
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "ControlAccesos.frx":810A
         PICN            =   "ControlAccesos.frx":8126
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons SSCommand1 
         Height          =   315
         Left            =   2760
         TabIndex        =   19
         Top             =   405
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Cargar Programas"
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
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "ControlAccesos.frx":8478
         PICN            =   "ControlAccesos.frx":8494
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "ControlAccesos.frx":87E6
         Left            =   165
         List            =   "ControlAccesos.frx":87F3
         TabIndex        =   18
         Text            =   "TODAS"
         Top             =   405
         Width           =   2370
      End
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":881A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":8974
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":8ACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":8C28
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":8D82
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":8EDC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":9036
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":9190
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":92EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":9444
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":959E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":96F8
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "ControlAccesos.frx":A10A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Nivel"
      Height          =   195
      Left            =   120
      TabIndex        =   25
      Top             =   1710
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Descripción"
      Height          =   195
      Left            =   120
      TabIndex        =   24
      Top             =   1380
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Perfil"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   23
      Top             =   570
      Width           =   420
   End
   Begin VB.Label Label7 
      Caption         =   "Perfil"
      Height          =   195
      Left            =   120
      TabIndex        =   22
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "ControlAccesos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oT As New ADODB.Recordset
Dim xR As New ADODB.Recordset
Dim xB As New ADODB.Recordset
Dim TblSub_4 As New ADODB.Recordset
Dim TblSub_5 As New ADODB.Recordset

Dim okUnload As Boolean
Const CONTSTOP = 50
Const maxView = 15
Public AutoSeek As Boolean
Const SpaceY = 300
Dim IsNewRecord As Boolean
Dim MaxUser As Integer
Dim OkOpen As Boolean

Private Sub Combo1_Click()
KewlButtons1.Enabled = True
End Sub

Private Sub CopiarPerfil_Click()
SSFrame1.Visible = True
KewlButtons1.Enabled = False
End Sub

Private Sub Form_Activate()
If Not OkOpen Then
    MaxUser = CONTSTOP
    Call LoadControls
    Call wShow
    Call LoadData
    OkOpen = True
End If
Me.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
oT.Close
xR.Close
TblSub_4.Close
TblSub_5.Close
OkOpen = False
End Sub

Private Sub Form_Load()
oT.Open "NombrePerfiles", Conn, 2, 3, 512
xR.Open "Perfiles", Conn, 2, 3, 512
TblSub_4.Open "Programas", Conn, 2, 3, 512
TblSub_5.Open "Programas", Conn

xB.Open "Select *  From NombrePerfiles", Conn
If Not xB.EOF Then Combo1.text = xB!Perfil
Do While Not xB.EOF
   Combo1.AddItem xB!Perfil
   xB.MoveNext
Loop
xB.Close

End Sub

Private Sub LoadControls()
Dim j As Integer
Screen.MousePointer = vbHourglass
For j = 1 To MaxUser
        Load MemVar_4(j)
        Load SubLabel4(j)
        Load MemVar_5(j)
        Load MemVar_6(j)
Next j
Screen.MousePointer = vbDefault
End Sub

Private Sub KewlButtons1_Click()
Dim xSql As String
If MsgBox("Esta seguro de Grabar el Perfil " + MemVar_2 + ", Sobre " + Combo1.text + "si existe sera Sobre Escrito ", vbYesNo, "") = vbYes Then
   xSql = "Delete From Perfiles Where Usuario='" & Combo1.text & "'"
   Conn.Execute (xSql)
   xSql = "INSERT INTO Perfiles (Usuario,Programa,Acceso,Total) "
   xSql = xSql + "SELECT '" & Combo1.text & "',Programa,Acceso,Total FROM Perfiles Where Usuario='" & MemVar_2 & "'"
   Conn.Execute (xSql)
   MsgBox "Copia Terminada", vbInformation
   MemVar_1 = Conn.Execute("Select IDperfil From NombrePerfiles Where Perfil='" & Combo1.text & "'").Fields(0)
   Call MemVar_1_LostFocus
End If
SSFrame1.Visible = False
End Sub

Private Sub SSCommand1_Click()
Dim i As Integer
i = 0
TblSub_5.MoveFirst
Do While Not TblSub_5.EOF
   MemVar_4(i).text = TblSub_5!Programa
   SubLabel4(i).Caption = TblSub_5!Descripcion
   MemVar_5(i).Value = False
   MemVar_6(i).Value = False
   i = i + 1
   TblSub_5.MoveNext
Loop
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
        Call omenu(Button.Key)
End Sub

Private Sub omenu(ByVal Key As String)
        Select Case Key
                Case "Grabar"
                        Call SaveData
                        Call SaveValores
                Case "Salida"
                                Unload Me
                Case "Browse"
                        Select Case ActiveControl.Name
                        Case "MemVar_1"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "NombrePerfiles"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        End Select
                Case "Imprimir"
                Case "Excel"
                    Call ExportaExcel("vPerfilesProgramas")
                Case "Borrar"
                Case "Top"
                    oT.Close
                    oT.Open "Select Top 1 * From NombrePerfiles Order By IDPerfil"
                    Call LoadData
                Case "Bottom"
                    oT.Close
                    oT.Open "Select Top 1 * From NombrePerfiles Order By IDPerfil DESC"
                    Call LoadData
                Case "Proximo"
                    oT.Close
                    oT.Open "Select Top 1 * From NombrePerfiles Where IDPerfil>'" & MemVar_1 & "' Order By IDPerfil"
                    Call LoadData
                Case "Previo"
                    oT.Close
                    oT.Open "Select Top 1 * From NombrePerfiles Where IDPerfil<'" & MemVar_1 & "' Order By IDPerfil DESC"
                    Call LoadData
        End Select
End Sub
Private Sub ExeBrow(oCod As String)
' Codigo Por desarrollar
End Sub

Private Sub LoadData()
If oT.EOF And oT.BOF Then
        okNew.Visible = True
        IsNewRecord = True
        Exit Sub
End If
okNew.Visible = False
okFind.Visible = True
MemVar_1.text = oT!IdPerfil
MemVar_2.text = oT!Perfil
Call LoadValores
End Sub

Private Sub SaveData()
Dim ok As Boolean
ok = False
        If MemVar_1.text = "" Then ok = True
        If ok Then Exit Sub
'        oT!Perfil = MemVar_2
'        oT.Update
'        If IsNewRecord Then
'            If Not MySeek(oT, Conn, "Usuarios", "Login='" & MemVar_1 & "'") Then
'                    ' call LoadData
'            End If
'        End If
'        IsNewRecord = False
End Sub

Private Sub Mark()
If TypeOf ActiveControl Is TextBox Then
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.text)
End If
End Sub

Private Function Validate(X As String, Y As Byte) As Boolean
Dim ok As Boolean
On Error GoTo errores
        Select Case Y
                Case 90
                        ok = True
                Case dbText
                        ok = True
                Case dbLong
                        If Val(X) >= -2147483648# And Val(X) <= 2147483647 Then
                                ok = True
                        End If
                Case dbDouble
                        ok = True
                Case dbSingle
                        ok = True
                Case dbByte
                        If Val(X) >= 0 And Val(X) <= 255 Then
                                ok = True
                        End If
                Case dbInteger
                        If Val(X) >= -32768 And Val(X) <= 32767 Then
                                ok = True
                        End If
                Case dbDate
                                        Dim xTem As Date
                                        xTem = CDate(X)
                                        ok = True
        End Select
        If Not ok Then
        MsgBox "Valor para ese campo no valido", vbOKOnly
        End If
salir:
        Validate = ok
        Exit Function
errores:
        Select Case Err.Number
                Case 13
                        Select Case Y
                                Case dbDate
                                        MsgBox "Fecha no valida", vbOKOnly
                        End Select
        End Select
        ok = False
        Resume salir
End Function
Private Sub MemVar_1_GotFocus()
        Call Mark
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

Private Sub MemVar_1_LostFocus()

If Not MySeek(oT, Conn, "NombrePerfiles", "IDPerfil='" & MemVar_1 & "'") Then
        Call LoadData
Else
        MsgBox "Perfil NO definido, Verifique", vbOKOnly
        MemVar_1.SetFocus
End If

End Sub
Private Sub MemVar_2_GotFocus()
        Call Mark
End Sub
Private Sub MemVar_2_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                Case vbKeyUp
                        MemVar_1.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_4(0).SetFocus
End Select
End Sub
'Private Sub MemVar_2_Change()
'        MemVar_2.DataChanged = True
'End Sub
Private Sub MemVar_2_LostFocus()
        If Not Validate(MemVar_2.text, 10) Then
                        MemVar_2.SetFocus
        Else
        End If
End Sub
Private Sub MemVar_3_GotFocus()
        Call Mark
End Sub
Private Sub MemVar_3_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                Case vbKeyUp
                        MemVar_2.SetFocus
                Case vbKeyDown, vbKeyReturn
                oMove.Value = 0
                MemVar_4(0).SetFocus
End Select
End Sub
'Private Sub MemVar_3_Change()
'        MemVar_3.DataChanged = True
'End Sub
Private Sub MemVar_3_LostFocus()
        If Not Validate(MemVar_3.text, 3) Then
                        MemVar_3.SetFocus
        Else
        End If
End Sub

Public Sub LoadValores()
Dim xP As New ADODB.Recordset
Dim i As Integer, j As Integer
Dim xSql As String

For i = 0 To MaxUser
        Call Limpia(i)
Next i
i = 0
j = 0
oMove.Value = 0
If Not MySeek(xR, Conn, "Perfiles", "Usuario='" & MemVar_2 & "'") Then
    xSql = "Select * From vPerfilesProgramas Where Usuario='" & MemVar_2 & "' Order By Tipo DESC, Programa "
    Set xP = Conn.Execute(xSql)
    
    Do While Not xP.EOF
        MemVar_4(j).text = IIf(IsNull(xP!Programa), "", xP!Programa)
        SubLabel4(j).Caption = xP!Descripcion
    '    If Not MySeek(TblSub_4, Conn, "Programas", "Programa='" & MemVar_4(j) & "'") Then
    '            SubLabel4(j).Caption = TblSub_4!Descripcion
    '    Else
    '            SubLabel4(j).Caption = ""
    '    End If
        MemVar_5(j).Value = xP!Acceso
        MemVar_6(j).Value = xP!total
        j = j + 1
        xP.MoveNext
    Loop
End If
Call AjustaMover
Call wShow
End Sub
Private Sub Limpia(ByVal i As Integer)
        MemVar_4(i).text = ""
        SubLabel4(i).Caption = ""
        MemVar_5(i).Value = 0
        MemVar_6(i).Value = 0
End Sub
Private Sub AjustaMover()
        oMove.Max = MaxUser
        oMove.SmallChange = maxView
        oMove.LargeChange = maxView
        If MaxUser < maxView Then
                oMove.Visible = False
        Else
                oMove.Visible = True
        End If
End Sub
Private Sub wShow()
Dim i As Long
Dim j As Long
        For i = 0 To MaxUser
                If i >= oMove.Value And i <= oMove.Value + maxView Then
                        MemVar_4(i).Top = j * SpaceY + MemVar_4(0).Top
                        MemVar_4(i).Visible = True
                        SubLabel4(i).Top = j * SpaceY + SubLabel4(0).Top
                        SubLabel4(i).Visible = True
                        MemVar_5(i).Top = j * SpaceY + MemVar_5(0).Top
                        MemVar_5(i).Visible = True
                        MemVar_6(i).Top = j * SpaceY + MemVar_6(0).Top
                        MemVar_6(i).Visible = True
                        j = j + 1
                Else
                        MemVar_4(i).Visible = False
                        SubLabel4(i).Visible = False
                        MemVar_5(i).Visible = False
                        MemVar_6(i).Visible = False
                End If
        Next i
End Sub
Private Sub oMove_Change()
        Call wShow
End Sub
Private Sub MemVar_4_GotFocus(Index As Integer)
        Call Mark
End Sub
Private Sub MemVar_4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        If (Shift And 2) = 2 Then
                Select Case KeyCode
                        Case vbKeyI
                                Call InsertRow(Index)
                        Case vbKeyD
                                Call DeleteRow(Index)
                End Select
                Exit Sub
        End If
        Select Case KeyCode
                Case vbKeyReturn
                        MemVar_5(Index).SetFocus
                Case vbKeyLeft
                        If (Shift And 2) = 2 Then
                                If Index > 0 Then
                                        Call rev(Index)
                                        MemVar_6(Index - 1).SetFocus
                                End If
                        End If
                Case vbKeyRight
                        If (Shift And 2) = 2 Then
                                MemVar_5(Index).SetFocus
                        End If
                Case vbKeyDown
                        If Index < MaxUser Then
                                Call revisa(Index)
                                MemVar_4(Index + 1).SetFocus
                        End If
                Case vbKeyUp
                        If Index > 0 Then
                                Call rev(Index)
                                MemVar_4(Index - 1).SetFocus
                        End If
        End Select
End Sub
'Private Sub MemVar_4_Change(Index As Integer)
'        MemVar_4(Index).DataChanged = True
'End Sub
Private Sub MemVar_4_LostFocus(Index As Integer)
        If Not Validate(MemVar_4(Index).text, 10) Then
                MemVar_4(Index).SetFocus
        Else
                'TblSub_4.Seek "=", MemVar_4(Index).Text
                If Not MySeek(TblSub_4, Conn, "Programas", "Programa='" & MemVar_4(Index) & "'") Then
                                SubLabel4(Index).Caption = TblSub_4!Descripcion
                        Else
                        If ActiveControl.Name <> "BrwBtn_4" Then
                        If Not (MemVar_4(Index).text = "") Then
                                MsgBox "El Programa o Control NO existe, Verifique"
                                MemVar_4(Index).SetFocus
                        End If
                        End If
                                SubLabel4(Index).Caption = ""
                        End If
        End If
End Sub
Private Sub MemVar_5_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        If (Shift And 2) = 2 Then
                Select Case KeyCode
                        Case vbKeyI
                                Call InsertRow(Index)
                        Case vbKeyD
                                Call DeleteRow(Index)
                End Select
                Exit Sub
        End If
        Select Case KeyCode
                Case vbKeyReturn
                        MemVar_6(Index).SetFocus
                Case vbKeyLeft
                        If (Shift And 2) = 2 Then
                                MemVar_4(Index).SetFocus
                        End If
                Case vbKeyRight
                        If (Shift And 2) = 2 Then
                                MemVar_6(Index).SetFocus
                        End If
                Case vbKeyDown
                        If Index < MaxUser Then
                                Call revisa(Index)
                                MemVar_5(Index + 1).SetFocus
                        End If
                Case vbKeyUp
                        If Index > 0 Then
                                Call rev(Index)
                                MemVar_5(Index - 1).SetFocus
                        End If
                Case vbKeySpace
                        MemVar_5(Index).Value = Not MemVar_5(Index).Value
        End Select
End Sub
Private Sub MemVar_6_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        If (Shift And 2) = 2 Then
                Select Case KeyCode
                        Case vbKeyI
                                Call InsertRow(Index)
                        Case vbKeyD
                                Call DeleteRow(Index)
                End Select
                Exit Sub
        End If
        Select Case KeyCode
                Case vbKeyReturn
                        If Index < MaxUser Then
                                Call revisa(Index)
                                MemVar_4(Index + 1).SetFocus
                End If
                Case vbKeyLeft
                        If (Shift And 2) = 2 Then
                                MemVar_5(Index).SetFocus
                        End If
                Case vbKeyRight
                        If (Shift And 2) = 2 Then
                                If Index < MaxUser Then
                                        Call revisa(Index)
                                        MemVar_4(Index + 1).SetFocus
                                End If
                        End If
                Case vbKeyDown
                        If Index < MaxUser Then
                                Call revisa(Index)
                                MemVar_6(Index + 1).SetFocus
                        End If
                Case vbKeyUp
                        If Index > 0 Then
                                Call rev(Index)
                                MemVar_6(Index - 1).SetFocus
                        End If
                Case vbKeySpace
                        MemVar_6(Index).Value = Not MemVar_6(Index).Value
        End Select
End Sub
Private Sub revisa(Index As Integer)
        If Index + 1 >= maxView + oMove.Value Then
                oMove.Value = oMove.Value + oMove.SmallChange
        End If
End Sub
Private Sub rev(Index As Integer)
        If Index - 1 < oMove.Value And Index > 0 Then
                If oMove.Value - oMove.SmallChange >= 0 Then
                        oMove.Value = oMove.Value - oMove.SmallChange
                Else
                        oMove.Value = 0
                End If
        End If
End Sub

Public Sub SaveValores()
Dim i As Long
Dim j As Integer
Dim xSql As String

xSql = "Delete Perfiles Where Usuario='" & MemVar_2 & "'"
Conn.Execute (xSql)

For i = 0 To MaxUser
    If MemVar_4(i).text <> "" Then
        xR.AddNew
            xR!Usuario = MemVar_2.text
            xR!Programa = MemVar_4(i).text
            xR!Acceso = MemVar_5(i).Value
            xR!total = MemVar_6(i).Value
        xR.Update
    End If
Next

End Sub

Private Sub DeleteRow(N As Integer)
Dim i As Integer
        Screen.MousePointer = vbHourglass
        For i = N To MaxUser - 1
                MemVar_4(i).text = MemVar_4(i + 1).text
                SubLabel4(i).Caption = SubLabel4(i + 1).Caption
                MemVar_5(i).Value = MemVar_5(i + 1).Value
                MemVar_6(i).Value = MemVar_6(i + 1).Value
        Next i
        Screen.MousePointer = vbDefault
End Sub

Private Sub InsertRow(N As Integer)
Dim i As Integer
        Screen.MousePointer = vbHourglass
        For i = MaxUser To N + 1 Step -1
                MemVar_4(i).text = MemVar_4(i - 1).text
                SubLabel4(i).Caption = SubLabel4(i - 1).Caption
                MemVar_5(i).Value = MemVar_5(i - 1).Value
                MemVar_6(i).Value = MemVar_6(i - 1).Value
        Next i
        MemVar_4(N).text = ""
                SubLabel4(N).Caption = ""
        MemVar_5(N).Value = 0
        MemVar_6(N).Value = 0
        Screen.MousePointer = vbDefault
End Sub

