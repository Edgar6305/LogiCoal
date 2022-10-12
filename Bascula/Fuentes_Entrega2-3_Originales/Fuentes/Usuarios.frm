VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Usuarios 
   Caption         =   "Gestion de Usuarios"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11595
   Icon            =   "Usuarios.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Usuarios.frx":058A
   ScaleHeight     =   5760
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
            NumListImages   =   19
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":0B14
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":0C6E
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":0DC8
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":0F22
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":107C
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":11D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":1330
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":148A
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":15E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":173E
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":1898
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":19F2
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":2404
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":299E
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":2F38
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":34D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":401C
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":50EE
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Usuarios.frx":6244
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
         Height          =   3975
         Left            =   4620
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
            TabIndex        =   17
            Top             =   1680
            Width           =   3495
         End
         Begin VB.TextBox MemVar_6 
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
            IMEMode         =   3  'DISABLE
            Left            =   1800
            PasswordChar    =   "*"
            TabIndex        =   16
            Top             =   2460
            Width           =   1395
         End
         Begin VB.TextBox MemVar_5 
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
            TabIndex        =   15
            Top             =   2100
            Width           =   3495
         End
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
            Picture         =   "Usuarios.frx":685A
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   6
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
            Picture         =   "Usuarios.frx":695C
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   5
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
                     Picture         =   "Usuarios.frx":6A5E
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Usuarios.frx":6D78
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Usuarios.frx":7092
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Usuarios.frx":73AC
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Usuarios.frx":76C6
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Usuarios.frx":79E0
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Usuarios.frx":7CFA
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Usuarios.frx":8014
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Usuarios.frx":832E
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Usuarios.frx":8648
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Usuarios.frx":8962
                     Key             =   ""
                  EndProperty
               EndProperty
            End
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
            TabIndex        =   4
            Top             =   960
            Width           =   3495
         End
         Begin VB.CheckBox MemVar_7 
            Alignment       =   1  'Right Justify
            Caption         =   "Activado"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   420
            TabIndex        =   3
            Top             =   2940
            Value           =   1  'Checked
            Width           =   1575
         End
         Begin VB.TextBox MemVar_3 
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
            TabIndex        =   2
            Top             =   1320
            Width           =   3495
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Perfil"
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
            Top             =   1740
            Width           =   390
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Login"
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
            Top             =   660
            Width           =   435
         End
         Begin VB.Label Label11 
            Caption         =   "Descripcion"
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
            Top             =   1020
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cargo"
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
            Top             =   1380
            Width           =   465
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Email"
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
            Top             =   2160
            Width           =   405
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Password"
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
            Left            =   420
            TabIndex        =   8
            Top             =   2520
            Width           =   735
         End
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   780
      Left            =   0
      TabIndex        =   18
      Top             =   0
      Width           =   19815
      _ExtentX        =   34951
      _ExtentY        =   1376
      BandCount       =   5
      _CBWidth        =   19815
      _CBHeight       =   780
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
      MinHeight5      =   360
      Width5          =   1800
      NewRow5         =   0   'False
      Begin KewlButtonz.KewlButtons Actualizar 
         Height          =   315
         Left            =   4935
         TabIndex        =   23
         Top             =   405
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
         MICON           =   "Usuarios.frx":8C7C
         PICN            =   "Usuarios.frx":8C98
         PICH            =   "Usuarios.frx":9232
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
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
         ItemData        =   "Usuarios.frx":97CC
         Left            =   825
         List            =   "Usuarios.frx":97DC
         TabIndex        =   21
         Text            =   "TODAS"
         Top             =   405
         Width           =   2355
      End
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   330
         Left            =   165
         TabIndex        =   20
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
         TabIndex        =   19
         Top             =   405
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
         MICON           =   "Usuarios.frx":9811
         PICN            =   "Usuarios.frx":982D
         PICH            =   "Usuarios.frx":9DC7
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
      TabIndex        =   22
      Top             =   5445
      Width           =   11595
      _ExtentX        =   20452
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
Attribute VB_Name = "Usuarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Usuarios As New ADODB.Recordset
Dim okUnload As Boolean
Dim IsNewRecord As Boolean
Dim OkOpen As Boolean
Dim oT As New ADODB.Recordset

Private Sub Actualizar_Click()
If MsgBox("Esta Seguro de Actualizar El Usuario ", vbYesNo, "Actualización de Usuarios") = vbYes Then
    xSql = "UPDATE Usuarios_T Set  Estado = AC"
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
On Error Resume Next
    Usuarios_T.Close
    MenuNuevo.Flag_xProg = ""
    Unload Me
End Sub

Private Sub Form_Load()
Dim xCt As New ADODB.Recordset

xSql = "Select Top 1 * From Usuarios_T " & xFilter
oT.Open xSql, Conn, 2, 3, 1

Set VSplitter.LeftOrTopCtl = oTree
Set VSplitter.RightOrBottomCtl = Frame1

Set xCt = Conn.Execute("Select * From NombrePerfiles")
If Not xCt.EOF Then Combo1.text = xCt!Perfil
Do While Not xCt.EOF
    Combo1.AddItem xCt!Perfil
    xCt.MoveNext
Loop
xCt.Close

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

If MemVar_2 = "" Then
    MsgBox "Debe escribir la Descripción del Usuario, Verifique", vbInformation
    MemVar_2.SetFocus
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

Private Sub MemVar_3_LostFocus()
If MemVar_3 = "" Then
    MsgBox "Debe escribir el Cargo del Usuario, Verifique", vbInformation
    MemVar_2.SetFocus
End If

End Sub

Private Sub MemVar_5_GotFocus()
    Call Mark(MemVar_5)
End Sub

Private Sub MemVar_5_LostFocus()
If MemVar_5 = "" Then
    MsgBox "Debe escribir el e-mail del Usuario, Verifique", vbInformation
    MemVar_2.SetFocus
End If

End Sub

Private Sub MemVar_6_GotFocus()
    Call Mark(MemVar_5)
End Sub

Private Sub MemVar_6_LostFocus()
If MemVar_6 = "" Then
    MsgBox "Debe escribir el PassWord del Usuario, Verifique", vbInformation
    MemVar_2.SetFocus
End If

End Sub

Private Sub oNuevo_Click()
If oNuevo.Caption = "Nuevo" Then
    oNuevo.Caption = "Cancelar"
    IsNewRecord = True
    MemVar_1.text = ""
    MemVar_2.text = ""
    MemVar_3.text = ""
    Combo1.ListIndex = 4
    MemVar_5 = ""
    MemVar_6 = ""
    MemVar_7 = 1

    oBar.Panels("Usuario").text = "USUARIO: " & Susuario
    oBar.Panels("Ot").text = "FECHA CREACION: " & Now
    oBar.Panels("Estado").text = "ESTADO: " & "IN"

Else
    oNuevo.Caption = "Nuevo"
    IsNewRecord = False
End If
MemVar_1.SetFocus
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
                    BrowseCatalogo.xtabla = "Usuarios_T"
                    Set BrowseCatalogo.dControl = ActiveControl
                    BrowseCatalogo.Show 1
            End Select
    Case "Imprime"
    
    Case "Borrar"
    
    Case "Top"
        oT.Close
        oT.Open "Select Top 1 * From Usuarios_T Order By Login"
        Call LoadData
    Case "Bottom"
        oT.Close
        oT.Open "Select Top 1 * From Usuarios_T Order By Login DESC"
        Call LoadData
    Case "Proximo"
        oT.Close
        oT.Open "Select Top 1 * From Usuarios_T Where Login>'" & MemVar_1 & "' Order By  Login"
        Call LoadData
    Case "Previo"
        oT.Close
        oT.Open "Select Top 1 * From Usuarios_T Where Login<'" & MemVar_1 & "' Order By Login DESC"
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
     MemVar_1 = oT!Login
     MemVar_2 = oT!Descripcion
     MemVar_3 = oT!Cargo
     Combo1.text = oT!Perfil
     MemVar_5 = oT!email
     MemVar_6 = oT!PassWord
     MemVar_7 = IIf(oT!Activado, 1, 0)
 Else
 
 End If

 oBar.Panels("Usuario").text = "USUARIO: " & oT!Login
 oBar.Panels("Ot").text = "FECHA CREACION: " & oT!FechaCreacion
 oBar.Panels("Estado").text = "ESTADO: " & oT!Estado
 
 Actualizar.Enabled = (oT!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','ActualizarUsuarios'").Fields(0)) '==> Fields(0) Nos dice si tiene acceso al Objeto

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
Dim ok As Boolean

On Error GoTo Recover

    ok = False
    If MemVar_1.text = "" Then ok = True
    If MemVar_2.text = "" Then ok = True
    If ok Then Exit Sub
    
    If IsNewRecord Then
        oT.AddNew
        oT!Login = MemVar_1
        oT!Descripcion = MemVar_2
        oT!Cargo = MemVar_3
        oT!Perfil = Combo1.text
        oT!email = MemVar_5
        oT!PassWord = Enc_Guardar(MemVar_6)
        oT!Activado = MemVar_7
    
    Else
        oT!Descripcion = MemVar_2
        oT!Cargo = MemVar_3
        oT!Perfil = Combo1.text
        oT!email = MemVar_5
        oT!PassWord = Enc_Guardar(MemVar_6)
        oT!Activado = MemVar_7
    End If
    
    oT.Update
    
    If IsNewRecord Then
        If Not MySeek(oT, Conn, "Usuarios_T", "Login=" & MemVar_1) Then
                 Call LoadData
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
Set Nodx = oTree.Nodes.Add(, , "0" & "Usuarios", "Usuarios", 13)

xSql = "SELECT DISTINCT Perfil  FROM  Usuarios_T"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("0" & "Usuarios", tvwChild, "A" & Em!Perfil, Em!Perfil, 14)
    Em.MoveNext
Wend

xSql = "SELECT * FROM Usuarios_T"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("A" & Em!Perfil, tvwChild, "B" & Em!Login, Em!Login & " " & Em!Descripcion, 14)
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

MemVar_1 = UCase(MemVar_1)
If Not IsNewRecord Then
     If Not MySeek(oT, Conn, "Usuarios_T", "Login='" & MemVar_1 & "'") Then
            Call LoadData
     Else
            MsgBox "Usuario NO Registrado, Verifique"
            MemVar_1.SetFocus
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga el Usuario" & vbCrLf & Err.Description
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
    MemVar_1 = Mid(Node.Key, 2, 10)
    Call MemVar_1_LostFocus
End Select

End Sub

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
                dControl.SelStart = 0
                dControl.SelLength = Len(dControl.text)
End If
End Sub


