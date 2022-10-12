VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Trituracion 
   Caption         =   "Orden de Trituracion"
   ClientHeight    =   9825
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13305
   LinkTopic       =   "Form1"
   ScaleHeight     =   9825
   ScaleWidth      =   13305
   StartUpPosition =   1  'CenterOwner
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
         NumListImages   =   24
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":015A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":02B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":040E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":0568
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":06C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":081C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":0976
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":0AD0
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":0C2A
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":0D84
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":0EDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":18F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":1E8A
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":2424
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":29BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":3EC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":4A12
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":4FCB
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":60E1
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":68B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":7985
            Key             =   ""
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":84CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Trituracion.frx":9625
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MBSplit.Splitter VSplitter 
      Height          =   8535
      Left            =   0
      TabIndex        =   0
      Top             =   780
      Width           =   13155
      _ExtentX        =   23204
      _ExtentY        =   15055
      SplitterPos     =   30
      Begin MSComctlLib.TreeView oTree 
         Height          =   3735
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   3555
         _ExtentX        =   6271
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
      Begin VB.Frame Body 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   8055
         Left            =   4440
         TabIndex        =   1
         Top             =   180
         Width           =   8475
         Begin VB.TextBox xIdTrituracion 
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
            Left            =   360
            TabIndex        =   42
            Top             =   6720
            Width           =   315
         End
         Begin VB.TextBox DesParo 
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
            Left            =   5040
            MaxLength       =   55
            TabIndex        =   40
            Top             =   6720
            Width           =   3015
         End
         Begin VB.TextBox MemVar_34 
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
            TabIndex        =   34
            Top             =   1560
            Width           =   1260
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
            TabIndex        =   3
            Top             =   360
            Width           =   1260
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
            MaxLength       =   15
            TabIndex        =   2
            Top             =   1200
            Width           =   1260
         End
         Begin MSMask.MaskEdBox HorasEfectivas 
            Height          =   315
            Left            =   4440
            TabIndex        =   4
            Top             =   1980
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0.00"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker Fecha 
            Height          =   315
            Left            =   1800
            TabIndex        =   5
            Top             =   780
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy hh:mm tt"
            Format          =   115867651
            CurrentDate     =   44578
         End
         Begin Threed.SSFrame oMarco 
            Height          =   1830
            Left            =   360
            TabIndex        =   6
            Top             =   2700
            Width           =   6615
            _Version        =   65536
            _ExtentX        =   11668
            _ExtentY        =   3228
            _StockProps     =   14
            ForeColor       =   8421504
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
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
               Alignment       =   2  'Center
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
               Index           =   0
               Left            =   60
               MaxLength       =   10
               TabIndex        =   9
               Top             =   450
               Width           =   960
            End
            Begin VB.VScrollBar oMove 
               Height          =   1560
               Left            =   6300
               TabIndex        =   8
               Top             =   120
               Width           =   255
            End
            Begin VB.TextBox MemVar_5 
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
               Index           =   0
               Left            =   1020
               MaxLength       =   25
               TabIndex        =   7
               Top             =   450
               Width           =   4120
            End
            Begin MSMask.MaskEdBox MemVar_6 
               Height          =   315
               Index           =   0
               Left            =   5160
               TabIndex        =   10
               Top             =   450
               Width           =   1125
               _ExtentX        =   1984
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               PromptChar      =   "_"
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Pila Destino"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   45
               TabIndex        =   13
               Top             =   120
               Width           =   960
            End
            Begin VB.Label SubLabel3 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
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
               Height          =   285
               Index           =   0
               Left            =   1020
               TabIndex        =   12
               Top             =   120
               Width           =   4120
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Porcentaje"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   5160
               TabIndex        =   11
               Top             =   120
               Width           =   1125
            End
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1425
            Left            =   360
            TabIndex        =   14
            Top             =   4920
            Width           =   7770
            _ExtentX        =   13705
            _ExtentY        =   2514
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
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
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "No Paro"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Fecha Inicio"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Fecha Final"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Minutos Paro"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Descripcion"
               Object.Width           =   4410
            EndProperty
         End
         Begin MSMask.MaskEdBox MemVar_3 
            Height          =   315
            Left            =   1800
            TabIndex        =   15
            Top             =   1980
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker FechaCierre 
            Height          =   315
            Left            =   5640
            TabIndex        =   37
            Top             =   780
            Width           =   2055
            _ExtentX        =   3625
            _ExtentY        =   556
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "dd/MM/yyyy hh:mm tt"
            Format          =   115867651
            CurrentDate     =   44578
         End
         Begin MSComCtl2.DTPicker FecIniParo 
            Height          =   315
            Left            =   720
            TabIndex        =   38
            Top             =   6720
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
            Format          =   115867651
            CurrentDate     =   44578
         End
         Begin MSComCtl2.DTPicker FecFinParo 
            Height          =   315
            Left            =   2820
            TabIndex        =   39
            Top             =   6720
            Width           =   2115
            _ExtentX        =   3731
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
            Format          =   115867651
            CurrentDate     =   44578
         End
         Begin KewlButtonz.KewlButtons NewParo 
            Height          =   555
            Left            =   360
            TabIndex        =   41
            Top             =   7200
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   979
            BTYPE           =   3
            TX              =   "Nuevo Paro"
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
            MICON           =   "Trituracion.frx":9C3B
            PICN            =   "Trituracion.frx":9C57
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin KewlButtonz.KewlButtons GrabarParo 
            Height          =   555
            Left            =   6480
            TabIndex        =   43
            Top             =   7200
            Visible         =   0   'False
            Width           =   1605
            _ExtentX        =   2831
            _ExtentY        =   979
            BTYPE           =   3
            TX              =   "Grabar Paro"
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
            MICON           =   "Trituracion.frx":A8A9
            PICN            =   "Trituracion.frx":A8C5
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Cierre"
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
            Left            =   4500
            TabIndex        =   36
            Top             =   840
            Width           =   930
         End
         Begin VB.Label Label_34 
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
            Left            =   3240
            TabIndex        =   35
            Top             =   1560
            Width           =   4395
         End
         Begin VB.Label Label11 
            Caption         =   "Trituradora"
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
            Index           =   1
            Left            =   420
            TabIndex        =   33
            Top             =   1620
            Width           =   1275
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Detalle del Despacho"
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
            Left            =   360
            TabIndex        =   25
            Top             =   2520
            Width           =   1515
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Paradas Reportadas"
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
            Left            =   360
            TabIndex        =   24
            Top             =   4620
            Width           =   1485
         End
         Begin VB.Label Label11 
            Caption         =   "Cantidad (Ton.) "
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
            Index           =   0
            Left            =   420
            TabIndex        =   23
            Top             =   2040
            Width           =   1275
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Numero de Orden"
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
            TabIndex        =   22
            Top             =   420
            Width           =   1275
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Pila Origen"
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
            TabIndex        =   21
            Top             =   1200
            Width           =   765
         End
         Begin VB.Label LabelPila 
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
            Left            =   3240
            TabIndex        =   20
            Top             =   1200
            Width           =   4395
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha"
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
            TabIndex        =   19
            Top             =   840
            Width           =   450
         End
         Begin VB.Label Label7 
            Caption         =   "Horas Efectivas"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Index           =   0
            Left            =   3240
            TabIndex        =   18
            Top             =   1980
            Width           =   1005
         End
         Begin Tracer.LabelPlus Cerrado 
            Height          =   675
            Left            =   3300
            TabIndex        =   17
            Top             =   180
            Visible         =   0   'False
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   1191
            BackColor       =   255
            BackColorOpacity=   0
            BackShadow      =   0   'False
            BorderColor     =   255
            BorderCornerLeftTop=   4
            BorderCornerRightTop=   4
            BorderCornerBottomRight=   4
            BorderCornerBottomLeft=   4
            BorderWidth     =   1
            CaptionAlignmentH=   2
            CaptionAlignmentV=   1
            Caption         =   "Trituracion.frx":AF43
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   16777215
            ShadowColorOpacity=   0
            CallOutAlign    =   0
            CallOutWidth    =   0
            CallOutLen      =   0
            MousePointer    =   0
            BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            IconForeColor   =   0
            IconOpacity     =   0
            PicturePresent  =   -1  'True
            PictureArr      =   "Trituracion.frx":AF63
         End
         Begin VB.Label Anulado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H000000FF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ANULADO"
            BeginProperty Font 
               Name            =   "Segoe UI Black"
               Size            =   9
               Charset         =   0
               Weight          =   900
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   4380
            TabIndex        =   16
            Top             =   360
            Visible         =   0   'False
            Width           =   1275
         End
      End
   End
   Begin ComctlLib.StatusBar oBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   27
      Top             =   9510
      Width           =   13305
      _ExtentX        =   23469
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
            Text            =   "ESTADO"
            TextSave        =   "ESTADO"
            Key             =   "Ot"
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
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   780
      Left            =   0
      TabIndex        =   28
      Top             =   0
      Width           =   19815
      _ExtentX        =   34951
      _ExtentY        =   1376
      BandCount       =   7
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
      Child4          =   "CierreLotes"
      MinHeight4      =   315
      Width4          =   1785
      NewRow4         =   0   'False
      Child5          =   "MuestraProduccion"
      MinHeight5      =   315
      Width5          =   2325
      NewRow5         =   0   'False
      MinHeight6      =   360
      Width6          =   2400
      NewRow6         =   0   'False
      MinHeight7      =   360
      Width7          =   1500
      NewRow7         =   0   'False
      Begin KewlButtonz.KewlButtons oNuevo 
         Height          =   315
         Left            =   3405
         TabIndex        =   32
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
         MICON           =   "Trituracion.frx":C3F9
         PICN            =   "Trituracion.frx":C415
         PICH            =   "Trituracion.frx":C9AF
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
         ItemData        =   "Trituracion.frx":CF49
         Left            =   825
         List            =   "Trituracion.frx":CF59
         TabIndex        =   31
         Text            =   "TODAS"
         Top             =   405
         Width           =   2355
      End
      Begin KewlButtonz.KewlButtons CierreLotes 
         Height          =   315
         Left            =   4935
         TabIndex        =   30
         Top             =   405
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Cerrar Orden"
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
         MICON           =   "Trituracion.frx":CF86
         PICN            =   "Trituracion.frx":CFA2
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
         TabIndex        =   29
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
   End
End
Attribute VB_Name = "Trituracion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim okUnload As Boolean
Const CONTSTOP = 5
Const maxView = 3
Const SpaceY = 300
Dim IsNewRecord As Boolean
Dim MaxUser As Integer
Dim OkOpen As Boolean
Dim oD As New ADODB.Recordset
Dim sHorasEfectivas As Integer

Private Sub CierreLotes_Click()
Dim xR As New ADODB.Recordset
Dim xSql As String
Dim xRes As String

'ALTER PROCEDURE [dbo].[PA_CierreOrdenTrituracion]
'@IdTrituracion int,
'@Fecha DateTime

If DateDiff("n", Fecha, FechaCierre) / 60 = 12 Then
    If MsgBox("Esta Seguro de Cerrar La Orden de Trituración ", vbYesNo, "Cerrar Ordenes de Trituración") = vbYes Then
        xSql = "SET DateFormat DMY EXEC PA_CierreOrdenTrituracion " & MemVar_1 & ",'" & Format(FechaCierre, "dd/MM/yyyy hh:mm") & "'"
        xRes = Conn.Execute(xSql).Fields(0)
        If xRes <> "OK" Then
            MsgBox xRes, vbCritical, "Error de Cierre de Orden"
            Exit Sub
        End If
        MsgBox "Orden de Trituración No " & MemVar_1 & " Fue Cerrada"
    End If
Else
    MsgBox "La Fecha de Cierre debe ser Mayor 12 Horas que la de Inicio, Verifique", vbInformation
    FechaCierre.SetFocus
End If

Call MemVar_1_LostFocus
MemVar_1.SetFocus
End Sub

Private Sub Form_Activate()
If Not OkOpen Then
        MaxUser = CONTSTOP
        Call LoadControls
        Call wShow
        If MemVar_1 <> "" Then
           'Set oT = Conn.Execute("Select * From PagosCuentas Where Codigo='" & MemVar_1 & "'")
        End If
        Call LoadData
        OkOpen = True
        Me.SetFocus
        Call MuestraArbol
        MemVar_1.SetFocus
End If
End Sub

Private Sub Form_Load()
Dim X As New ADODB.Recordset
    Set VSplitter.LeftOrTopCtl = oTree
    Set VSplitter.RightOrBottomCtl = Body
    
    oD.Open "Trituracion", Conn, 2, 3, 512
    If Not oD.EOF Then oD.MoveLast
    FecIniParo = Now
    FecFinParo = Now
End Sub

Private Sub LoadControls()
Dim j As Integer
    Screen.MousePointer = vbHourglass
    For j = 1 To MaxUser
            Load MemVar_4(j)
            Load MemVar_5(j)
            Load MemVar_6(j)
    Next j
    Screen.MousePointer = vbDefault
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
                        MemVar_5(i).Top = j * SpaceY + MemVar_5(0).Top
                        MemVar_5(i).Visible = True
                        MemVar_6(i).Top = j * SpaceY + MemVar_6(0).Top
                        MemVar_6(i).Visible = True
                        j = j + 1
                Else
                        MemVar_4(i).Visible = False
                        MemVar_5(i).Visible = False
                        MemVar_6(i).Visible = False
                End If
        Next i
End Sub

Private Sub DeleteRow(N As Integer)
Dim i As Integer
    For i = N To MaxUser - 1
            MemVar_4(i).text = MemVar_4(i + 1).text
            MemVar_5(i).text = MemVar_5(i + 1).text
            MemVar_6(i).text = MemVar_6(i + 1).text
    Next i
End Sub

Private Sub InsertRow(N As Integer)
Dim i As Integer
    For i = MaxUser To N + 1 Step -1
            MemVar_4(i).text = MemVar_4(i - 1).text
            MemVar_5(i).text = MemVar_5(i - 1).text
            MemVar_6(i).text = MemVar_6(i - 1).text
    Next i
    MemVar_4(N).text = ""
    MemVar_5(N).text = ""
    MemVar_6(N).text = ""
End Sub

Private Sub Limpia(ByVal i As Integer)
        MemVar_4(i).text = ""
        MemVar_5(i).text = ""
        MemVar_6(i).text = ""
End Sub

Public Sub LoadValores()
Dim i As Integer
Dim xR As New ADODB.Recordset
Dim xSql As String

For i = 0 To MaxUser
  Call Limpia(i)
Next i
i = 0
oMove.Value = 0

xSql = "SELECT * From vTrituracionDetalle Where IdTrituracion=" & MemVar_1
Set xR = Conn.Execute(xSql)

Do While Not xR.EOF
    MemVar_4(i).text = xR!PilaDestino
    MemVar_5(i).text = xR!Descripcion
    MemVar_6(i).text = xR!Porcentaje
    i = i + 1
    xR.MoveNext
Loop

Call AjustaMover
Call wShow

End Sub

Private Sub GrabarParo_Click()
Dim xR As New ADODB.Recordset
Dim xSql As String
Dim xRes As String

'ALTER PROCEDURE [dbo].[PA_ParoTrituracion]
'@IdTrituracion int,
'@FecIni Datetime,
'@FecFin Datetime,
'@Descripcion Varchar(55),
'@Usuario as Varchar(10)

If xIdTrituracion = "" Then
     If MsgBox("Esta Seguro de Adicionar Paro a La Orden de Trituración ", vbYesNo, "Paros Ordenes de Trituración") = vbYes Then
        xSql = "Set DateFormat DMY EXEC PA_ParoTrituracion " & MemVar_1 & ",'" & Format(FecIniParo, "dd/MM/yyyy hh:mm") & "','" & Format(FecFinParo, "dd/MM/yyyy hh:mm") & "','" & DesParo & "','" & Susuario & "'"
        xRes = Conn.Execute(xSql).Fields(0)
        If xRes <> "OK" Then
            MsgBox xRes, vbCritical, "Error de Creacion de Paro"
        End If
    End If
Else
     If MsgBox("Esta Seguro de Modificar Paro a La Orden de Trituración ", vbYesNo, "Paros Ordenes de Trituración") = vbYes Then
        xSql = "Set DateFormat DMY EXEC PA_EditarParoTrituracion " & xIdTrituracion & ",'" & Format(FecIniParo, "dd/MM/yyyy hh:mm") & "','" & Format(FecFinParo, "dd/MM/yyyy hh:mm") & "','" & DesParo & "','" & Susuario & "'"
        xRes = Conn.Execute(xSql).Fields(0)
        If xRes <> "OK" Then
            MsgBox xRes, vbCritical, "Error de Creacion de Paro"
        End If
    End If
End If
GrabarParo.Visible = False
Call CargaParos(MemVar_1)
End Sub

Private Sub ListView1_Click()

    If Val("0" & ListView1.SelectedItem) > 0 Then
        xIdTrituracion = ListView1.SelectedItem
        FecIniParo = ListView1.SelectedItem.SubItems(1)
        FecFinParo = ListView1.SelectedItem.SubItems(2)
        DesParo = ListView1.SelectedItem.SubItems(4)
        GrabarParo.Visible = True
    End If
End Sub

Private Sub MemVar_4_GotFocus(Index As Integer)
        Call Mark(MemVar_4(Index))
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
        Case vbKeyF4
            MemVar_4(Index).text = ""
            Call omenu("Browse")
        Case vbKeyReturn
                MemVar_6(Index).SetFocus
        Case vbKeyLeft
                If (Shift And 2) = 2 Then
                        If Index > 0 Then
                                Call rev(Index)
                                MemVar_4(Index - 1).SetFocus
                        End If
                End If
        Case vbKeyRight
                If (Shift And 2) = 2 Then
                        MemVar_6(Index).SetFocus
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

Private Sub MemVar_4_LostFocus(Index As Integer)
Dim xR As New ADODB.Recordset

On Error Resume Next


If MemVar_4(Index) <> "" Then
    If IsNumeric(MemVar_4(Index)) Then
        Set xR = Conn.Execute("Select * From vPilasGeneral Where IdPila=" & MemVar_4(Index).text)
        If Not xR.EOF Then
            MemVar_5(Index) = xR!Descripcion & " - " & xR!Desacopio & " - " & xR!Ubicacion
            MemVar_6(Index).SetFocus
        Else
            MsgBox "Pila NO localizada, verifique", vbInformation
            MemVar_4(Index).SetFocus
        End If
    Else
        MsgBox "La Pila debe ser Un Numero Entero", vbInformation
        MemVar_4(Index) = ""
        MemVar_4(Index).SetFocus
    End If
End If
End Sub

Private Sub MemVar_6_GotFocus(Index As Integer)
        Call Mark(MemVar_6(Index))
End Sub

Private Sub MemVar_6_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next

Select Case KeyCode
        Case vbKeyF4
                Call omenu("Browse")
        Case vbKeyReturn
                If Index < MaxUser Then
                        Call revisa(Index)
                        MemVar_4(Index + 1).SetFocus
        End If
        Case vbKeyLeft
                If (Shift And 2) = 2 Then
                        MemVar_4(Index).SetFocus
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
End Select

End Sub

Private Sub MuestraCalidad_Click()
Dim xSql As String
Dim xRes As String

If MsgBox("Esta Seguro de Crear Muestra de Calidad ", vbYesNo, "Crear Muestras de Calidad") = vbYes Then
    xSql = " EXEC PA_CreaRegCalidad " & "'DS', " & MemVar_1 & ", 1, 0,'" & Susuario & "'"
    xRes = Conn.Execute(xSql).Fields(0)
    If xRes <> "OK" Then
        MsgBox xRes, vbCritical, "Error de Creacion de Muestra"
    Else
        MsgBox "Muestra de Calidad Creada"
    End If
End If

End Sub

Private Sub MuestraProduccion_Click()
Dim xSql As String
Dim xRes As String
Dim xCantidadCalidad As Single

If MsgBox("Esta Seguro de Crear Muestra de Producción ", vbYesNo, "Crear Muestras de Producción") = vbYes Then

'@TranOrigen varchar(2),
'@Numerotransaccion Int,
'@TipoMuestra int,
'@Cantidad float,
'@Usuario Varchar(10)

    xCantidadCalidad = CantidadRecibidaBascula - MemVar_36
    xSql = " EXEC PA_CreaRegCalidad " & "'TR', " & MemVar_1 & ", 2, " & xCantidadCalidad & ",'" & Susuario & "'"
    
    xRes = Conn.Execute(xSql).Fields(0)
    If xRes <> "OK" Then
        MsgBox xRes, vbCritical, "Error de Creacion de Muestra"
    Else
        xSql = "EXEC PA_ActualizaTRMuestraCalidad " & MemVar_1 & ", " & CantidadRecibidaBascula
        xRes = Conn.Execute(xSql).Fields(0)
        If xRes <> "OK" Then
            MsgBox xRes, vbCritical, "Error Actualizacion Cantidad Ultima Muestra"
        End If
        MsgBox "Muestra de Producción Creada, Ultima Cantidad de Muestra Actualizada"
    End If
End If

End Sub

Private Sub NewParo_Click()
    xIdTrituracion = ""
    FecIniParo = Fecha
    FecFinParo = Fecha
    DesParo = ""
    GrabarParo.Visible = True
End Sub

Private Sub oMove_Change()
        Call wShow
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
                                BrowseCatalogo.xtabla = "Ventas"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        Case "MemVar_2", "MemVar_4"
                                BrowseAcopioPilas.x_Tipo = 1
                                BrowseAcopioPilas.xtabla = "vPilasAcopiosGeneral"
                                Set BrowseAcopioPilas.dControl = ActiveControl
                                BrowseAcopioPilas.Show 1
                        Case "MemVar_34"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "Trituradoras"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                                
                        End Select
                Case "Imprime"
                     RepCotOs.Show
                Case "Borrar"
                    If MsgBox("Esta seguro de Borrar La Orden de Traslado", vbYesNo, "Borrado de Registro") = vbYes Then
                        If CantidadRecibidaBascula = 0 Then
                            Conn.Execute ("Update Traslados Set Estado='AN' Where IdTraslado=" & MemVar_1)
                            Call MySeek(oD, Conn, "Traslados", "IdVenta=" & MemVar_1)
                            Call LoadData
                        Else
                            MsgBox "NO puede borrar el lote porque presenta Cantidad Recibida", vbInformation
                            MemVar_1.SetFocus
                            Exit Sub
                        End If
                    End If
                Case "Top"
                    oD.Close
                    oD.Open "Select Top 1 * From Trituracion Order By IdTrituracion"
                    Call LoadData
                Case "Bottom"
                    oD.Close
                    oD.Open "Select Top 1 * From Trituracion Order By IdTrituracion DESC"
                    Call LoadData
                Case "Proximo"
                    oD.Close
                    oD.Open "Select Top 1 * From Trituracion Where IdTrituracion>'" & MemVar_1 & "' Order By IdTrituracion"
                    Call LoadData
                Case "Previo"
                    oD.Close
                    oD.Open "Select Top 1 * From Trituracion Where  IdTrituracion<'" & MemVar_1 & "' Order By IdTrituracion DESC"
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
Dim xR As New ADODB.Recordset
Dim xEstado As Boolean

On Error GoTo Recover

 If oD.EOF And oD.BOF Then
'         okNew.Visible = True
'         okFind.Visible = False
         Exit Sub
 End If

MemVar_1 = oD!IdTrituracion
Fecha = oD!FechaInicio
FechaCierre = IIf(IsNull(oD!FechaCierre), Now, oD!FechaCierre)
MemVar_2 = oD!PilaOrigen

Set xR = Conn.Execute("SELECT * From vPilasGeneral Where IdPila=" & MemVar_2)
LabelPila = xR!Descripcion & " - " & xR!Desacopio
xR.Close
MemVar_3 = oD!Cantidad
MemVar_34 = oD!IdTrituradora
Call MemVar_34_LostFocus
HorasEfectivas = oD!HorasEfectivas


CierreLotes.Enabled = (oD!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','CierreTrituracion'").Fields(0)) '==> Fields(0) Nos dice si tiene acceso al Objeto

Tbar.buttons("Grabar").Enabled = (oD!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','" & sProgSISMA & "'").Fields(1)) '==> Fields(1) Nos dice puede hacer cambios al Objeto
Tbar.buttons("Borrar").Enabled = (oD!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','" & sProgSISMA & "'").Fields(1)) '==> Fields(1) Nos dice puede hacer cambios al Objeto

Anulado.Visible = (oD!Estado = "AN")
oBar.Panels("Usuario").text = "USUARIO: " & oD!Usuario & " " & Format(oD!FechaInicio, "dd/MM/yyyy hh:mm")
oBar.Panels("Ot").text = "ESTADO: " & oD!Estado

xEstado = IIf(oD!Estado = "IN", 1, 0)
Cerrado.Visible = xEstado
NewParo.Visible = xEstado
Cerrado.Visible = IIf(oD!Estado = "AC", 1, 0)
Anulado.Visible = IIf(oD!Estado = "AN", 1, 0)

Call LoadValores
Call CargaParos(MemVar_1)
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
Dim xSql As String
Dim IdCliente As Integer
Dim Res As String
Dim i As Integer
Dim xSuma As Integer
        
On Error GoTo Recover
        
        If MemVar_4(0) = "" Then
           MsgBox "Faltan los Datos de La Pila Destino", vbInformation
           MemVar_4(0).SetFocus
           Exit Sub
        End If

        xSuma = 0
        For i = 0 To CONTSTOP - 1
                xSuma = xSuma + Val("0" + MemVar_6(i))
        Next
        If xSuma <> 100 Then
            MsgBox "La sumatoria de los Porcentajes debe ser 100, Verifique", vbInformation
            MemVar_4(0).SetFocus
            Exit Sub
        End If

'ALTER PROCEDURE [dbo].[PA_OrdenTrituracion]
'@IdTrituradora int,
'@FecIni datetime,
'@Pila int,
'@Usuario Varchar(15)

        If IsNewRecord Then
             xSql = "SET DateFormat DMY EXEC PA_OrdenTrituracion " & MemVar_34 & ",'" & Format(Fecha, "dd/MM/yyyy hh:mm") & "'," & MemVar_2 & ",'" & Susuario & "'"
            
            Res = Conn.Execute(xSql).Fields(0)
            If Res <> "OK" Then
                 MsgBox "Error al Grabar Encabezado del Trituración, Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
                 Exit Sub
            End If
            
            MemVar_1 = Conn.Execute("Select TOP 1 IdTrituracion From Trituracion Where Usuario='" & Susuario & "' And Estado ='IN' Order By IdTrituracion DESC").Fields(0)
            
        Else
            oD!PilaOrigen = MemVar_2
            oD!FechaInicio = Fecha
            oD!IdTrituradora = MemVar_34
            oD.Update
        End If
        
          For i = 0 To CONTSTOP - 1
            If MemVar_4(i) <> "" Then
                xSql = "EXEC PA_OrdenTrituracionDetalle " & MemVar_1 & "," & MemVar_4(i) & "," & MemVar_6(i) & "," & i
                Res = Conn.Execute(xSql).Fields(0)
                 If Res <> "OK" Then
                     MsgBox "Error al Grabar Detalle de la Trituracion, Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
                     Exit Sub
                 End If
            End If
        Next
    
        If Not MySeek(oD, Conn, "Trituracion", "idTrituracion=" & MemVar_1) Then
                 Call LoadData
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

Private Sub oNuevo_Click()
Dim i As Integer
    
If oNuevo.Caption = "Nuevo" Then
    oNuevo.Caption = "Cancelar"
    Tbar.buttons("Grabar").Enabled = True
    Tbar.buttons("Borrar").Enabled = True

    IsNewRecord = True
    MemVar_1.text = ""
    MemVar_2.text = ""
    MemVar_3.text = 0
    MemVar_34.text = ""
    HorasEfectivas = 0
    
    Fecha = Format(Now, "dd/MM/yyyy") & " 06:00"
    FechaCierre = Format(Now, "dd/MM/yyyy") & " 18:00"
    FecIniParo = Format(Now, "dd/MM/yyyy") & " 06:00"
    FecFinParo = Format(Now, "dd/MM/yyyy") & " 06:00"
    
    LabelPila = ""
    Label_34 = ""
    
    For i = 0 To CONTSTOP - 1
      Call Limpia(i)
    Next i
    
    ListView1.ListItems.Clear
    
    oBar.Panels("Usuario").text = "USUARIO: " & Susuario & " " & Format(Now, "dd/MM/yyyy hh:mm")
    oBar.Panels("Ot").text = "ESTADO: " & "IN"
    
    Cerrado.Visible = False
    Anulado.Visible = False
    CierreLotes.Visible = False
    NewParo.Visible = False
    GrabarParo.Visible = False

Else
    oNuevo.Caption = "Nuevo"
    IsNewRecord = False
    oD.MoveLast
    MemVar_1 = oD!IdTraslado
    Call MemVar_1_LostFocus
End If
MemVar_2.SetFocus
End Sub

Private Sub MemVar_1_GotFocus()
        Call Mark(MemVar_1)
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

Private Sub MemVar_1_LostFocus()

On Error GoTo Recover

If MemVar_1 <> "" Then
    If Not MySeek(oD, Conn, "Trituracion", "IdTrituracion=" & MemVar_1) Then
            Call LoadData
    Else
            MsgBox "Numero de Orden NO Localizada"
            MemVar_1.SetFocus
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Lote" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_1_LostFocus()"
    Err.Clear
    Resume Next
End If
    
End Sub

Private Sub MemVar_2_GotFocus()
        Call Mark(MemVar_2)
End Sub

Private Sub MemVar_2_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                        MemVar_2 = ""
                        Call omenu("Browse")
                Case vbKeyUp
                        MemVar_1.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_34.SetFocus
End Select
End Sub

Private Sub MemVar_2_LostFocus()
Dim xR As New ADODB.Recordset

On Error GoTo Recover

If MemVar_2 <> "" Then
    Set xR = Conn.Execute("SELECT * From vPilasGeneral Where IdPila=" & MemVar_2)
    If Not xR.EOF Then
        LabelPila = xR!Descripcion & " - " & xR!Desacopio
    Else
        MsgBox "Pila NO Localizada, Verifique", vbInformation
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un Error al  Cargar Pila" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_2_LostFocus()"
    Err.Clear
    Resume Next
End If
    
End Sub

Private Sub MemVar_3_GotFocus()
        Call Mark(MemVar_3)
End Sub

Private Sub MemVar_3_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                        MemVar_3 = ""
                        Call omenu("Browse")
                Case vbKeyUp
                Case vbKeyDown, vbKeyReturn
                        MemVar_34.SetFocus
End Select
End Sub

Private Sub MemVar_3_LostFocus()

On Error GoTo Recover

If MemVar_3 > 0 Then
Else
    MsgBox "Cantidad Errada, Verifique", vbInformation
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Lote" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_3_LostFocus()"
    Err.Clear
    Resume Next
End If
    
End Sub

Private Sub MemVar_34_GotFocus()
        Call Mark(MemVar_34)
End Sub

Private Sub MemVar_34_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                        Call omenu("Browse")
                Case vbKeyUp
                        MemVar_3.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_4(0).SetFocus
End Select
End Sub

Private Sub MemVar_34_LostFocus()
Dim xR As New ADODB.Recordset

On Error GoTo Recover

If MemVar_34 <> "" Then
    Set xR = Conn.Execute("Select * FROM Trituradoras WHERE IdTrituradora=" & MemVar_34)
    If Not xR.EOF Then
        Label_34.Caption = xR!Descripcion
    Else
        MsgBox "Trituradora NO localizada, Verifique", vbInformation
        MemVar_34.SetFocus
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Cargar Trituradoras" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_34_LostFocus()"
    Err.Clear
    Resume Next
End If
    
End Sub

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
                dControl.SelStart = 0
                dControl.SelLength = Len(dControl.text)
End If
End Sub

Private Function Validate(X As String, Y As Byte) As Boolean
        Validate = True
End Function

Private Sub MuestraArbol()
Dim Em As New ADODB.Recordset
Dim xSql As String

On Error GoTo Recover

oTree.Nodes.Clear
Set Nodx = oTree.Nodes.Add(, , "A" & "Trituracion", "Trituracion", 13)

xSql = "SELECT  Distinct DesTrituradora FROM  vTrituracion Where Estado='IN' "
Set Em = Conn.Execute(xSql)

If Em.EOF Then Exit Sub

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("A" & "Trituracion", tvwChild, "B" & Em!DesTrituradora, Em!DesTrituradora, 14)
    Em.MoveNext
Wend

xSql = "SELECT *   FROM  vTrituracion Where Estado='IN' "
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("B" & Em!DesTrituradora, tvwChild, "C" & Format(Em!IdTrituracion, "000000"), Format(Em!IdTrituracion, "000000") & " " & Em!Despila & " " & Format(Em!FechaInicio, "dd/MM/YY"), 15)
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
    MSG = "Se produjo un error al  Cargar Orden desde el Arbol," & vbCrLf & Err.Description
    MsgBox MSG, , "MuestraArbol()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub

Private Sub oTree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer

Select Case Mid(Node.Key, 1, 1)
Case "C"
    CargaOrden (Val(Mid(Node.Key, 2, 6)))
End Select

End Sub

Private Sub CargaOrden(ByVal Num As Long)
Dim xSql As String

On Error GoTo Recover
If Not MySeek(oD, Conn, "Trituracion", "IdTrituracion=" & Num) Then
    MemVar_1 = oD!IdTrituracion
    Call LoadData
End If
Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Orden desde el Arbol," & vbCrLf & Err.Description
    MsgBox MSG, , "CargaOrden"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub

Private Sub CargaParos(ByVal xTrituracion As Long)
Dim xSql As String
Dim xTm As New ADODB.Recordset

On Error GoTo Recover

    xSql = "SELECT * FROM  TrituradoraParos WHERE iDTrituracion=" & xTrituracion
    
    Set xTm = Conn.Execute(xSql)
    ListView1.ListItems.Clear
    If xTm.EOF Then
        Set iTmx = ListView1.ListItems.Add()
        iTmx.text = "No Data..."
    End If
    
    Do While Not xTm.EOF
        Set iTmx = ListView1.ListItems.Add()
        iTmx.text = xTm!IdParoTrituradora
        iTmx.SubItems(1) = Format(xTm!FechaInicio, "dd/MM/yyyy hh:mm")
        iTmx.SubItems(2) = IIf(IsNull(xTm!FechaFin), "Corriendo....", Format(xTm!FechaFin, "dd/MM/yyyy hh:mm"))
        iTmx.SubItems(3) = IIf(IsNull(xTm!FechaFin), DateDiff("n", xTm!FechaInicio, Now), DateDiff("n", xTm!FechaInicio, xTm!FechaFin))
        iTmx.SubItems(4) = xTm!Descripcion
        xTm.MoveNext
    Loop
    xTm.Close

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer vFrmLotes_LotesCalidad," & vbCrLf & Err.Description
    MsgBox MSG, , "CargaMuestrasCalidad()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

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



