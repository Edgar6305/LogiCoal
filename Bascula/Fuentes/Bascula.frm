VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Bascula 
   Caption         =   "Bascula"
   ClientHeight    =   9045
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15210
   FillStyle       =   2  'Horizontal Line
   Icon            =   "Bascula.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9045
   ScaleWidth      =   15210
   StartUpPosition =   1  'CenterOwner
   Begin MBSplit.Splitter VSplitter1 
      Height          =   8115
      Left            =   0
      TabIndex        =   13
      Top             =   780
      Width           =   15135
      _ExtentX        =   26696
      _ExtentY        =   14314
      SplitterPos     =   30
      Begin MSComctlLib.TreeView oTree 
         Height          =   7755
         Left            =   120
         TabIndex        =   29
         Top             =   180
         Width           =   4335
         _ExtentX        =   7646
         _ExtentY        =   13679
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
      Begin Threed.SSFrame Body 
         Height          =   7935
         Left            =   4740
         TabIndex        =   14
         Top             =   60
         Width           =   10335
         _Version        =   65536
         _ExtentX        =   18230
         _ExtentY        =   13996
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
         Begin VB.Frame Adicionales 
            Height          =   675
            Left            =   1800
            TabIndex        =   38
            Top             =   2340
            Visible         =   0   'False
            Width           =   4335
            Begin VB.ComboBox Combo1 
               Appearance      =   0  'Flat
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Left            =   2400
               TabIndex        =   40
               Top             =   180
               Width           =   1815
            End
            Begin VB.CheckBox Carpado 
               Alignment       =   1  'Right Justify
               Caption         =   "Encarpado"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   60
               TabIndex        =   39
               Top             =   240
               Width           =   1155
            End
            Begin VB.Label Label7 
               BackColor       =   &H80000004&
               Caption         =   "Tipo Carbón"
               BeginProperty Font 
                  Name            =   "Segoe UI Semibold"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   600
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   3
               Left            =   1320
               TabIndex        =   41
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.TextBox MemVar_31 
            Appearance      =   0  'Flat
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
            Left            =   2940
            TabIndex        =   37
            Top             =   1260
            Width           =   435
         End
         Begin MSCommLib.MSComm MSComm1 
            Left            =   1140
            Tag             =   "50000"
            Top             =   7080
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
         End
         Begin VB.TextBox Memvar_51 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFEEC&
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   975
            Left            =   1800
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   32
            Top             =   3120
            Width           =   4575
         End
         Begin KewlButtonz.KewlButtons UsoTara 
            Height          =   675
            Left            =   5760
            TabIndex        =   10
            Top             =   4200
            Visible         =   0   'False
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   1191
            BTYPE           =   7
            TX              =   "Uso Tara"
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
            MICON           =   "Bascula.frx":058A
            PICN            =   "Bascula.frx":05A6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.TextBox MemVar_3 
            Appearance      =   0  'Flat
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
            Left            =   1800
            TabIndex        =   4
            Top             =   1260
            Width           =   915
         End
         Begin VB.TextBox MemVar_4 
            Appearance      =   0  'Flat
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
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   6
            Top             =   1620
            Width           =   1575
         End
         Begin VB.ComboBox MemVar_1 
            Appearance      =   0  'Flat
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
            ItemData        =   "Bascula.frx":11A8
            Left            =   1800
            List            =   "Bascula.frx":11B8
            TabIndex        =   2
            Top             =   360
            Width           =   1695
         End
         Begin VB.ComboBox MemVar_2 
            Appearance      =   0  'Flat
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
            Left            =   1800
            TabIndex        =   3
            Top             =   720
            Width           =   3795
         End
         Begin VB.TextBox MemVar_5 
            Appearance      =   0  'Flat
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
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   7
            Text            =   "0"
            Top             =   2040
            Width           =   1575
         End
         Begin VB.TextBox MemVar_6 
            Appearance      =   0  'Flat
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
            Left            =   8460
            TabIndex        =   16
            Text            =   "1"
            Top             =   900
            Visible         =   0   'False
            Width           =   795
         End
         Begin VB.TextBox Label_6 
            Appearance      =   0  'Flat
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
            Left            =   9300
            TabIndex        =   15
            Top             =   900
            Visible         =   0   'False
            Width           =   495
         End
         Begin KewlButtonz.KewlButtons Grabar 
            Height          =   675
            Left            =   3720
            TabIndex        =   11
            Tag             =   "0"
            Top             =   4980
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   1191
            BTYPE           =   7
            TX              =   "Grabar"
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
            BCOL            =   15790320
            BCOLO           =   15790320
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Bascula.frx":11DA
            PICN            =   "Bascula.frx":11F6
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin KewlButtonz.KewlButtons ImprimeTiquete 
            Height          =   675
            Left            =   3720
            TabIndex        =   12
            Tag             =   "0"
            Top             =   5760
            Visible         =   0   'False
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   1191
            BTYPE           =   7
            TX              =   "Imprimir"
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
            BCOL            =   16777215
            BCOLO           =   16777215
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Bascula.frx":1874
            PICN            =   "Bascula.frx":1890
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox MemVar_7 
            Height          =   645
            Left            =   1800
            TabIndex        =   17
            Top             =   4200
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   1138
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###,##0"
            PromptChar      =   "_"
         End
         Begin KewlButtonz.KewlButtons LeerPeso 
            Height          =   675
            Left            =   3720
            TabIndex        =   9
            Tag             =   "0"
            Top             =   4200
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   1191
            BTYPE           =   7
            TX              =   "Leer Peso"
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
            BCOL            =   -2147483643
            BCOLO           =   -2147483643
            FCOL            =   0
            FCOLO           =   0
            MCOL            =   12632256
            MPTR            =   1
            MICON           =   "Bascula.frx":2662
            PICN            =   "Bascula.frx":267E
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin MSMask.MaskEdBox MemVar_8 
            Height          =   645
            Left            =   1800
            TabIndex        =   18
            Top             =   4980
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   1138
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MemVar_9 
            Height          =   645
            Left            =   1800
            TabIndex        =   19
            Top             =   5760
            Width           =   1635
            _ExtentX        =   2884
            _ExtentY        =   1138
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   20.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###,##0"
            PromptChar      =   "_"
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   300
            Top             =   7080
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   29
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":3C78
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":3DD2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":3F2C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":4086
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":41E0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":433A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":4494
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":45EE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":4748
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":48A2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":49FC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":4B56
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":5568
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":5B02
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":609C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":6636
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":6D6C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":7242
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":7690
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":87E6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":9180
                  Key             =   ""
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":A466
                  Key             =   ""
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":B038
                  Key             =   ""
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":B80A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":BDC3
                  Key             =   ""
               EndProperty
               BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":C595
                  Key             =   ""
               EndProperty
               BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":C8E7
                  Key             =   ""
               EndProperty
               BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":CC45
                  Key             =   ""
               EndProperty
               BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Bascula.frx":D7D7
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin Crystal.CrystalReport oCr 
            Left            =   2100
            Top             =   7140
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            Destination     =   1
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
            CopiesToPrinter =   2
            DiscardSavedData=   -1  'True
            WindowState     =   2
            PrintFileLinesPerPage=   60
            WindowShowGroupTree=   -1  'True
            WindowAllowDrillDown=   -1  'True
            WindowShowCloseBtn=   -1  'True
            WindowShowSearchBtn=   -1  'True
            WindowShowPrintSetupBtn=   -1  'True
            WindowShowRefreshBtn=   -1  'True
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
            Left            =   3480
            TabIndex        =   36
            Top             =   1620
            Width           =   2115
         End
         Begin Tracer.LabelPlus LabelPlusMsg 
            Height          =   855
            Left            =   6900
            TabIndex        =   35
            Top             =   4140
            Visible         =   0   'False
            Width           =   3315
            _ExtentX        =   5847
            _ExtentY        =   1508
            BackColor       =   -2147483624
            Border          =   -1  'True
            BorderColor     =   32896
            BorderCornerLeftTop=   2
            BorderCornerRightTop=   2
            BorderCornerBottomRight=   2
            BorderCornerBottomLeft=   2
            BorderWidth     =   1
            Caption         =   "Bascula.frx":EA15
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShadowSize      =   3
            ShadowColor     =   32896
            ShadowOffsetX   =   2
            ShadowOffsetY   =   2
            BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PictureArr      =   0
         End
         Begin Tracer.LabelPlus LabelMensajes 
            Height          =   2235
            Left            =   5760
            TabIndex        =   34
            Top             =   5100
            Visible         =   0   'False
            Width           =   4515
            _ExtentX        =   7964
            _ExtentY        =   3942
            BackColor       =   -2147483624
            Border          =   -1  'True
            BorderColor     =   32896
            BorderCornerLeftTop=   2
            BorderCornerRightTop=   2
            BorderCornerBottomRight=   2
            BorderCornerBottomLeft=   2
            BorderWidth     =   1
            Caption         =   "Bascula.frx":EA35
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShadowSize      =   3
            ShadowColor     =   32896
            ShadowOffsetX   =   2
            ShadowOffsetY   =   2
            BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PictureArr      =   0
         End
         Begin Tracer.LabelPlus LabelLotes 
            Height          =   1995
            Left            =   5820
            TabIndex        =   33
            Top             =   360
            Visible         =   0   'False
            Width           =   2295
            _ExtentX        =   4048
            _ExtentY        =   3519
            BackColor       =   16777215
            Border          =   -1  'True
            BorderColor     =   16576
            BorderCornerLeftTop=   2
            BorderCornerRightTop=   2
            BorderCornerBottomRight=   2
            BorderCornerBottomLeft=   2
            BorderWidth     =   1
            CaptionAlignmentH=   1
            Caption         =   "Bascula.frx":EA55
            CaptionPaddingX =   3
            CaptionPaddingY =   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ShadowSize      =   2
            ShadowOffsetX   =   3
            ShadowOffsetY   =   3
            HotLine         =   -1  'True
            HotLineColor    =   8438015
            HotLineWidth    =   25
            HotLinePosition =   1
            BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PictureArr      =   0
         End
         Begin VB.Label Label5 
            Caption         =   "Observaciones"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   420
            TabIndex        =   31
            Top             =   3120
            Width           =   1335
         End
         Begin VB.Label Label1 
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
            Height          =   255
            Left            =   360
            TabIndex        =   28
            Top             =   1260
            Width           =   1155
         End
         Begin VB.Label Label2 
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
            Height          =   255
            Left            =   360
            TabIndex        =   27
            Top             =   1620
            Width           =   1155
         End
         Begin VB.Label Label6 
            Caption         =   "Material"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   26
            Top             =   360
            Width           =   1155
         End
         Begin VB.Label Label3 
            Caption         =   "Orden o Lote"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   25
            Top             =   780
            Width           =   1155
         End
         Begin VB.Label LabelTransportador 
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
            Left            =   3480
            TabIndex        =   5
            Top             =   1260
            Width           =   2115
         End
         Begin VB.Label Label4 
            Caption         =   "Remision"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   2040
            Width           =   1155
         End
         Begin VB.Label Label5 
            Caption         =   "Maquina Cargue"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   8460
            TabIndex        =   23
            Top             =   660
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label Label7 
            Caption         =   "Peso Lleno"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   420
            TabIndex        =   22
            Top             =   4380
            Width           =   1275
         End
         Begin VB.Label Label7 
            Caption         =   "Peso Neto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   420
            TabIndex        =   21
            Top             =   5880
            Width           =   1275
         End
         Begin VB.Label Label7 
            Caption         =   "Peso Vacio"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   420
            TabIndex        =   20
            Top             =   5100
            Width           =   1275
         End
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15255
      _ExtentX        =   26908
      _ExtentY        =   1376
      BandCount       =   5
      _CBWidth        =   15255
      _CBHeight       =   780
      _Version        =   "6.0.8169"
      Child1          =   "Tbar"
      MinHeight1      =   330
      Width1          =   17040
      NewRow1         =   0   'False
      Child2          =   "Transacciones"
      MinHeight2      =   315
      Width2          =   3405
      NewRow2         =   -1  'True
      Child3          =   "AnularTiquete"
      MinHeight3      =   315
      Width3          =   1740
      NewRow3         =   0   'False
      MinHeight4      =   315
      Width4          =   1740
      NewRow4         =   0   'False
      MinHeight5      =   360
      Width5          =   1995
      NewRow5         =   0   'False
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   330
         Left            =   165
         TabIndex        =   30
         Top             =   30
         Width           =   15000
         _ExtentX        =   26458
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
      Begin VB.ComboBox Transacciones 
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   315
         ItemData        =   "Bascula.frx":EA89
         Left            =   165
         List            =   "Bascula.frx":EA9C
         TabIndex        =   1
         Text            =   "1- RECEPCION PESO LLENO"
         Top             =   405
         Width           =   3210
      End
      Begin KewlButtonz.KewlButtons AnularTiquete 
         Height          =   315
         Left            =   3600
         TabIndex        =   8
         Top             =   405
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Anular Tiquete"
         ENAB            =   0   'False
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Bascula.frx":EB1A
         PICN            =   "Bascula.frx":EB36
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   -1  'True
      End
   End
End
Attribute VB_Name = "Bascula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SnombreConductor As String
Dim SIdTransportador As Integer
Dim SnumeroTiquete As Long
Dim SusoTaraSISMA As Boolean
Dim mFont2 As StdFont
Dim sMuestra As Single
Dim SpesoMAX As Single, SpesoMIN As Single
Dim xMsg As String, xMsg2 As String
Dim Splaca As String
Dim Stransportador
Dim OkOpen As Boolean

Private Sub AnularTiquete_Click()
                        
If MsgBox("Esta seguro de Anular El Tiquete ", vbYesNo, "Borrado de Registro") = vbYes Then
    Set Anulaciones.dControl = MemVar_1
    Anulaciones.MemVar_1 = 1
    Anulaciones.MemVar_2 = SnumeroTiquete
    Anulaciones.MemVar_3 = Now
    Anulaciones.Show 1
    If MemVar_1.Tag Then
       Conn.Execute ("Update Bascula Set Estado='AN' Where IdTiquete=" & SnumeroTiquete)
    End If
End If

Call MuestraRecepcion
Call Limpia
End Sub

Private Sub Form_Activate()
If Not OkOpen Then
     OkOpen = True
    Transacciones.SetFocus
End If
Me.SetFocus
End Sub

Private Sub Form_Load()
Dim xR As New ADODB.Recordset
    'set the child controls for the vertical splitter
    Set VSplitter1.LeftOrTopCtl = oTree
    Set VSplitter1.RightOrBottomCtl = Body
    oTree.Height = VSplitter1.Height - 1500
    
    Set mFont2 = New StdFont
    mFont2.Name = "Segoe UI"
    mFont2.Size = 7

    SusoTaraSISMA = False
    Call CargaMateriales
    Set xR = Conn.Execute("Select * From TiposCarbon Order By Descripcion DESC")
    If Not xR.EOF Then Combo1 = Format(xR!IdTipoCarbon, "00") & " " & xR!Descripcion
    Do While Not xR.EOF
        Combo1.AddItem Format(xR!IdTipoCarbon, "00") & " " & xR!Descripcion
        xR.MoveNext
    Loop
    xR.Close
End Sub

Private Sub Form_Resize()
    Dim lSplitHeight As Long
    Dim i As Integer
    If Me.WindowState = vbMinimized Then Exit Sub
    lSplitHeight = Me.ScaleHeight - CoolBar1.Height - 100
    VSplitter1.Height = IIf(lSplitHeight < 0, 0, lSplitHeight)
End Sub

Private Sub Grabar_Click()
Dim xTiquete As Long

Grabar.Visible = False
UsoTara.Visible = False
LeerPeso.Visible = False

On Error GoTo Recover

Select Case Mid(Transacciones, 1, 1)
    Case 3 '
                Call SaveData
                Call Limpia
    Case 1
        If MemVar_7 = 0 Then
                 MsgBox "Peso Lleno está en CERO", vbExclamation
                 LeerPeso.SetFocus
        Else
                Call SaveData
                'xTiquete = Conn.Execute("Select TOP 1 IdTiquete From Bascula Where Placas='" & MemVar_3 & "'  And Usuario='" & Susuario & "'  AND Estado='IN' ORDER BY IdTiquete DESC").Fields(0)
                If SusoTaraSISMA Then
                    ImprimeTiquete.Visible = True
                    ImprimeTiquete.SetFocus
                Else
                    Call Limpia
                End If
        End If
    
    Case 5
        If MemVar_7 = 0 Then
                 MsgBox "Peso Lleno está en CERO", vbExclamation
                 LeerPeso.SetFocus
        Else
                Call SaveData
                xTiquete = SnumeroTiquete
                ImprimeTiquete.Visible = True
                ImprimeTiquete.SetFocus
        End If
    
    Case 2
        If MemVar_8 <= 0 Then
                 MsgBox "Peso Vacio está en CERO", vbExclamation
                 LeerPeso.SetFocus
        Else
                Call SaveData
                xTiquete = SnumeroTiquete
                ImprimeTiquete.Visible = True
                ImprimeTiquete.SetFocus
        End If
        
    Case 4
        If MemVar_8 <= 0 Then
                 MsgBox "Peso Vacio está en CERO", vbExclamation
                 LeerPeso.SetFocus
        Else
                Call SaveData
                xTiquete = Conn.Execute("Select TOP 1 IdTiquete From Bascula Where Placas='" & MemVar_3 & "'  And Usuario='" & Susuario & "'  AND Estado='IN' ORDER BY IdTiquete DESC").Fields(0)
                Call Limpia
        End If
End Select

'==> ALARMAS POR PESO MINIMO Y MAXIMO
If LabelPlusMsg.Visible = True Then
    xSql = "EXEC PA_Alarma " & xTiquete & ",'01003' ,'" & xMsg & xMsg2 & " ','" & Susuario & "'"
    Res = Conn.Execute(xSql).Fields(0)
     
     If Res <> "OK" Then
         MsgBox "Error al Grabar el Registro, Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
     End If
     LabelPlusMsg.Visible = False
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer Puerto COM," & vbCrLf & Err.Description
    MsgBox MSG, , "LeerPeso_Click()"
    Err.Clear   ' Borra campos del objeto Err
    Exit Sub
End If
End Sub

Private Sub Grabar_GotFocus()

On Error Resume Next

If MemVar_1.Tag = 0 Then
    MsgBox "Falta Tipo de Transaccion"
    MemVar_1.SetFocus
ElseIf MemVar_2.Tag = 0 Then
    MsgBox "Falta Informacion de Orden o Lote "
    MemVar_1.SetFocus
ElseIf MemVar_3.Tag = 0 Then
    MsgBox "Falta Placa del Vehiculo"
    MemVar_3.SetFocus
ElseIf MemVar_5.Tag = 0 Then
    MsgBox "Falta Remision"
    MemVar_5.SetFocus
End If

End Sub

Private Sub ImprimeTiquete_Click()
Dim xSql As String
Dim xTrTipo As String

On Error GoTo Recover

If MemVar_7 = 0 Or MemVar_8 = 0 Then
     MsgBox "Falta Una Pesada, Verifique antes de Imprimir", vbExclamation
     LeerPeso.SetFocus
Else
    Select Case Mid(Transacciones, 1, 1)
    Case 1
        If SusoTaraSISMA Then
            SnumeroTiquete = Conn.Execute("Select TOP 1 IdTiquete From Bascula Where Placas='" & MemVar_3 & "' And  IdTransaccion=1 And Usuario='" & Susuario & "' And UsoTara=1 ORDER BY IdTiquete DESC").Fields(0)
                    
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
            xSql = xSql + " Where    Bascula.IdTiquete = " & SnumeroTiquete
        End If
    
        Conn.Execute ("Delete RepTiqueteRecepcion")
        Conn.Execute ("Set Dateformat dmy Insert INTO RepTiqueteRecepcion " & xSql)
        oCr.ReportFileName = sDataReportPath + "RepTiqueteRecepcion.Rpt"
        oCr.Action = 1
        
    Case 2
        If Conn.Execute("SELECT TransaccionOrigen FROM Bascula WHERE IdTiquete=" & SnumeroTiquete).Fields(0) = "LT" Then
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
            xSql = xSql + " Where    Bascula.IdTiquete = " & SnumeroTiquete
        
            Conn.Execute ("Delete RepTiqueteRecepcion")
            Conn.Execute ("Set Dateformat dmy Insert INTO RepTiqueteRecepcion " & xSql)
            oCr.ReportFileName = sDataReportPath + "RepTiqueteRecepcion.Rpt"
            oCr.Action = 1
        Else
            xSql = " SELECT Bascula.IdTiquete, Bascula.Documentoasociado, Materiales.Descripcion, Transportador.Descripcion AS DesTransportador, Bascula.Placas, Bascula.Conductor, Bascula.PesoLleno, Bascula.PesoVacio, Bascula.FechaLleno, Bascula.FechaVacio,"
            xSql = xSql + "                  Bascula.FechaLlegada , Bascula.UsoTara, Bascula.Observaciones, Bascula.Usuario,Usuarios_T.Descripcion AS DesUsuario, Usuarios_T.Cargo"
            xSql = xSql + "   FROM  Bascula INNER JOIN"
            xSql = xSql + "                 Materiales ON Bascula.IdMaterial = Materiales.IdMaterial INNER JOIN"
            xSql = xSql + "                 Transportador ON Bascula.IdTransportador = Transportador.IdTransportador INNER JOIN"
            xSql = xSql + "                 Usuarios_T ON Bascula.Usuario = Usuarios_T.Login"
            xSql = xSql + " Where    Bascula.IdTiquete = " & SnumeroTiquete
            
            Conn.Execute ("Delete RepTiqueteRecepcionOtros")
            Conn.Execute ("Set Dateformat dmy Insert INTO RepTiqueteRecepcionOtros " & xSql)
            oCr.ReportFileName = sDataReportPath + "RepTiqueteRecepcionOtros.Rpt"
            oCr.Action = 1
        End If
            
    Case 5
    
        If Conn.Execute("SELECT TransaccionOrigen FROM Bascula WHERE IdTiquete=" & SnumeroTiquete).Fields(0) = "DS" Then ' <== VENTAS
            xSql = "        SELECT  Bascula.IdTiquete, Bascula.Documentoasociado AS Remision, Bascula.Placas, Bascula.Conductor, Bascula.PesoLleno, Bascula.PesoVacio, "
            xSql = xSql + "         Bascula.FechaLleno, Bascula.FechaVacio, Bascula.FechaLlegada, Bascula.Usuario AS DesUsuario, Bascula.Observaciones,"
            xSql = xSql + "         Materiales.Descripcion AS Desmaterial, Transportador.Descripcion AS DesTransportador, PilasFisicas.Descripcion, Terceros.Identificacion, Terceros.Descripcion AS DesTercero, "
            xSql = xSql + "         Ventas.OrdenCompraCliente , Acopios.Descripcion AS DesAcopio, Acopios.Ubicacion, Ventas.CantidadPedida, Ventas.CantidadDespachada , Usuarios_T.Descripcion AS DesUsuario_T, "
            xSql = xSql + "         Usuarios_T.Cargo, Conductores.Nombre,Bascula.IdTipoCarbon"
            xSql = xSql + " FROM    Bascula INNER JOIN"
            xSql = xSql + "         Ventas ON Bascula.TransaccionOrigen = 'DS' AND Bascula.NumeroTransaccion = Ventas.IdVentas INNER JOIN"
            xSql = xSql + "         VentasDetalle ON Ventas.IdVentas = VentasDetalle.IdVenta INNER JOIN"
            xSql = xSql + "         Pilas ON VentasDetalle.IdPila = Pilas.IdPila INNER JOIN"
            xSql = xSql + "         PilasFisicas ON Pilas.IdPilaFisica = PilasFisicas.IdPilaFisica INNER JOIN"
            xSql = xSql + "         Materiales ON Bascula.IdMaterial = Materiales.IdMaterial INNER JOIN"
            xSql = xSql + "         Transportador ON Bascula.IdTransportador = Transportador.IdTransportador INNER JOIN"
            xSql = xSql + "         Terceros ON Ventas.IdCliente = Terceros.IdCliente INNER JOIN"
            xSql = xSql + "         Acopios ON Pilas.IdAcopio = Acopios.IdAcopio INNER JOIN"
            xSql = xSql + "         Usuarios_T ON Bascula.Usuario = Usuarios_T.Login INNER JOIN"
            xSql = xSql + "         Conductores ON Bascula.Conductor = Conductores.Cedula"
            
            xSql = xSql + " WHERE  Bascula.IdTiquete = " & SnumeroTiquete
        
            Conn.Execute ("Delete RepTiqueteDespacho")
            Conn.Execute ("Set Dateformat dmy Insert INTO RepTiqueteDespacho" & xSql)
            oCr.ReportFileName = sDataReportPath + "RepTiqueteDespacho.Rpt"
            oCr.Action = 1
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
            xSql = xSql + " WHERE   Bascula.IdTiquete = " & SnumeroTiquete
            
            Conn.Execute ("Delete RepTiqueteTraslados")
            Conn.Execute ("Set Dateformat dmy Insert INTO RepTiqueteTraslados" & xSql)
            oCr.ReportFileName = sDataReportPath + "RepTiqueteTraslado.Rpt"
            oCr.Action = 1
        End If
        
        Call Alarmas(SnumeroTiquete)
        
    End Select
    
    Call BorraRpt(oCr, 1)
    Call Limpia
End If
Exit Sub

Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al Imprimir," & vbCrLf & Err.Description
    MsgBox MSG, , "ImprimeTiquete_Click()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub

Private Sub Alarmas(ByVal xTiquete As Long)
Dim xR As New ADODB.Recordset
Dim j As Integer

Set xR = Conn.Execute("Select * From Bascula where IdTiquete=" & xTiquete)
j = 1
If xR!TransaccionOrigen = "DS" Then
    xSql = "Select CantidadDespachada-UltimaMuestra-(IntervalosMuestras*1000)  From Ventas Where IdVentas=" & xR!NumeroTransaccion
    sMuestra = Conn.Execute(xSql).Fields(0) ' Revision de Cantidades para Muestras
    If sMuestra > 0 Then
        LabelMensajes.Visible = True
        LabelMensajes.Caption = Trim(Str(j)) + "- Favor Tomar Muestra de Calidad, se encuentra Pasada en " & Format(sMuestra, "####,###0") & " Kilos" & vbCrLf & vbCrLf
        xSql = "EXEC PA_Alarma " & xTiquete & ",'01001' ," & " ' Favor Tomar Muestra de Calidad, se encuentra Pasada en " & Format(sMuestra, "####,###0") & " Kilos" & " ','" & Susuario & "'"
        Conn.Execute (xSql)
        j = j + 1
    End If

    xSql = "Select ROUND((Cantidadpedida- CantidadDespachada)/CantidadPedida*100,0)  From Ventas Where IdVentas=" & xR!NumeroTransaccion
    sMuestra = Conn.Execute(xSql).Fields(0) ' Revision de Cantidad Despachada vs Cantidad Pedida
    If sMuestra <= 0 Then
        LabelMensajes.Visible = True
        LabelMensajes.Caption = LabelMensajes.Caption + Trim(Str(j)) + "- La Orden de Despacho Llego a su Limite, CERRAR orden, porcentaje excedido " & Format(sMuestra * -1, "#0") & " %"
        xSql = "EXEC PA_Alarma " & xTiquete & ",'01002' ,'" & " La Orden de Despacho Llego a su Limite, CERRAR orden, porcentaje excedido " & Format(sMuestra * -1, "#0") & " %" & " ','" & Susuario & "'"
        Conn.Execute (xSql)
    End If
End If
xR.Close

End Sub

Private Sub Limpia()
    MemVar_2.text = ""
    MemVar_3.text = ""
    MemVar_4.text = ""
    MemVar_5.text = "0"
    MemVar_51.text = ""
    MemVar_6.text = ""
    MemVar_7.text = ""
    MemVar_8.text = ""
    MemVar_9.text = ""
    LabelTransportador = ""
    Transacciones.SetFocus
    LabelLotes.Visible = False
    
End Sub

Private Sub LabelMensajes_Click()
LabelMensajes.Visible = False
End Sub

Private Sub LabelMensajes_MouseEnter()
    LabelMensajes.ForeColor = vbRed
End Sub

Private Sub LabelMensajes_MouseLeave()
LabelMensajes.ForeColor = &H8000&
End Sub

Private Sub LabelPlusMsg_PostPaint(ByVal HDC As Long)
Dim mTop As Long, TextHeight As Long
Dim lWidth As Long
Dim lMargin As Long
Dim xFlag As Boolean, xSql As String

On Error Resume Next

mTop = 5  '100 - .BackColorOpacity / 1.5
lMargin = 5

Select Case Val(Mid(Transacciones, 1, 1))
    Case 1, 5
        With LabelPlusMsg
            xMsg = "El Peso Excede el Peso Máximo Registrado "
            xMsg2 = "Báscula: " & Format(MemVar_7, "###,##0") & " Peso Máximo :  " + Format(SpesoMAX, "###,##0")
            
            lWidth = ((.Width / .GetWindowsDPI / Screen.TwipsPerPixelX) - lMargin * 2)    '100= aproximate height
                                                                     
            TextHeight = .DrawText(HDC, sTitle, lMargin, mTop, lWidth, 200, mFont2, vbWhite, 100, ccEnter, cTop, True)
            TextHeight = TextHeight + .DrawText(HDC, xMsg, lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H80000012, 100, cLeft, cTop, True)
            TextHeight = TextHeight + .DrawText(HDC, xMsg2, lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H80000012, 100, cLeft, cTop, True)
            
        End With
    Case 2, 4
        With LabelPlusMsg
            xMsg = "El Peso está por DEBAJO del peso Mínimo Registrado "
            xMsg2 = "Báscula: " & Format(MemVar_8, "###,##0") & " Peso Mínimo :  " + Format(SpesoMIN, "###,##0")
            
            lWidth = ((.Width / .GetWindowsDPI / Screen.TwipsPerPixelX) - lMargin * 2)    '100= aproximate height
                                                                     
            TextHeight = .DrawText(HDC, sTitle, lMargin, mTop, lWidth, 200, mFont2, vbWhite, 100, ccEnter, cTop, True)
            TextHeight = TextHeight + .DrawText(HDC, xMsg, lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H80000012, 100, cLeft, cTop, True)
            TextHeight = TextHeight + .DrawText(HDC, xMsg2, lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H80000012, 100, cLeft, cTop, True)
        End With
End Select

End Sub

Private Sub LeerPeso_Click()
Dim xNeto As Single
Dim xPeso As String, xMsg As String
Dim xPesoSTR As String
Dim i As Long
Dim xPuerto As Integer

On Error GoTo Recover

'1- RECEPCION PESO LLENO
'2- RECEPCION PESO VACIO
'3- DESPACHOS LLEGADA
'4- DESPACHOS PESO VACIO
'5- DESPACHOS PESO LLENO

Select Case Val(Mid(Transacciones, 1, 1))
    Case 1
        xPuerto = sPuertoBascula1
    Case 2
        xPuerto = sPuertoBascula2
    Case 4
        xPuerto = sPuertoBascula4
    Case 5
        xPuerto = sPuertoBascula5
End Select

MSComm1.CommPort = xPuerto
MSComm1.PortOpen = True
i = 0
Do While i < MSComm1.Tag And MSComm1.InBufferCount < 7
  DoEvents
  i = i + 1
Loop

If i >= MSComm1.Tag Then
    MsgBox "Agotado el tiempo de espera del Lector de Bascula, vuelva a intentarlo", vbDefaultButton3
    LeerPeso.SetFocus
    MSComm1.PortOpen = False
    Exit Sub
End If

xPesoSTR = MSComm1.Input

For i = 1 To Len(xPesoSTR)
    If Asc(Mid(xPesoSTR, i, 1)) >= 48 And Asc(Mid(xPesoSTR, i, 1)) <= 57 Then
        xPeso = xPeso + Mid(xPesoSTR, i, 1)
    End If
Next

MSComm1.PortOpen = False

Select Case Val(Mid(Transacciones, 1, 1))
    Case 1, 5
        MemVar_7 = xPeso
        If MemVar_7 > SpesoMAX Then
            LabelPlusMsg.Visible = True
        End If
    Case 2, 4
        MemVar_8 = xPeso
        If MemVar_8 < SpesoMIN Then
            LabelPlusMsg.Visible = True
        End If
End Select

xNeto = MemVar_7.text - MemVar_8.text
MemVar_9.text = IIf(xNeto < 0, 0, xNeto)

Grabar.Visible = True
Grabar.SetFocus

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer Puerto COM," & sPuertoBascula & vbCrLf & Err.Description
    MsgBox MSG, , "LeerPeso_Click()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub

Private Sub MemVar_1_KeyDown(KeyCode As Integer, Shift As Integer)
Dim T As Integer
T = Mid(Transacciones, 1, 1)

    Select Case KeyCode
            Case vbKeyF4
                    Call omenu("Browse")
            Case vbKeyUp
                    Transacciones.SetFocus
            Case vbKeyDown, vbKeyReturn
                    If T = 1 Or T = 3 Or T = 4 Then
                        MemVar_2.SetFocus
                    Else
                        oTree.SetFocus
                    End If
    End Select
End Sub

Private Sub MemVar_1_LostFocus()
Dim xRec As New ADODB.Recordset
Dim xSql As String

On Error GoTo Recover
MemVar_2.Clear

If MemVar_1 <> "" Then
    Select Case Mid(Transacciones, 1, 1)
        Case 1, 2
            If Val(Mid(MemVar_1.text, 1, 2)) = 1 Then '==>Si es Carbon
                 xSql = "Select * from vLotesBascula Where Estado='IN' "
                 Set xRec = Conn.Execute(xSql)
                 If xRec.EOF Then
                    MsgBox "NO hay Lotes Abiertos, Verifique"
                    MemVar_1.SetFocus
                    Exit Sub
                 Else
                    MemVar_2.text = "Lote : " + Format(xRec!IdLote, "000000") + " Manto " + xRec!Manto + " Bloque " + xRec!Bloque + " Tipo " + xRec!Descripcion
                    Do While Not xRec.EOF
                           MemVar_2.AddItem "Lote : " + Format(xRec!IdLote, "000000") + "  " + xRec!Manto + "  " + xRec!Bloque + " " + xRec!Descripcion
                           xRec.MoveNext
                    Loop
                 End If
                 xRec.Close
            Else
                 MemVar_2.text = "Otros Materiales"
            End If
            
        Case 3, 4
            If Val(Mid(MemVar_1.text, 1, 2)) = 1 Then '==>Si es Carbon
                ' ==> VENTAS
                 xSql = "Select * from vVentasBascula Where Estado='IN' "
                 Set xRec = Conn.Execute(xSql)
                 If Not xRec.EOF Then
                    MemVar_2.text = "Orden : " + Format(xRec!IdVentas, "000000") + "  " + xRec!Descripcion
                    Do While Not xRec.EOF
                           MemVar_2.AddItem "Orden : " + Format(xRec!IdVentas, "000000") + "  " + xRec!Descripcion
                           xRec.MoveNext
                    Loop
                 End If
                
                ' ==> TRASLADOS
                 xSql = "Select * from vTrasladosBascula Where Estado='IN' "
                 Set xRec = Conn.Execute(xSql)
                 If Not xRec.EOF Then
                    MemVar_2.text = "Tras. : " + Format(xRec!IdTraslado, "000000") + "  " + xRec!Desacopio
                    Do While Not xRec.EOF
                           MemVar_2.AddItem "Tras. : " + Format(xRec!IdTraslado, "000000") + "  " + xRec!Desacopio
                           xRec.MoveNext
                    Loop
                 End If
                 xRec.Close
                 
                 If MemVar_2.ListCount = 0 Then
                    MsgBox "NO hay Orden de Despacho ni Traslados, verifique", vbInformation
                    MemVar_1.SetFocus
                 End If
                 
            End If
        Case Else
                 MemVar_2.text = 0
                MemVar_3.SetFocus
    End Select
    MemVar_1.Tag = 1
Else
    MemVar_1.Tag = 0
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer vLotesBascula," & vbCrLf & Err.Description
    MsgBox MSG, , "Memvar_1_LostFocus()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
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
Dim xR As New ADODB.Recordset
Dim xSql As String
Dim Lote As Long
Dim Orden As Long

If Val(Mid(MemVar_1.text, 1, 2)) = 1 Then '==>Si es Carbon
    Select Case Mid(Transacciones, 1, 1)
        Case 1 To 2
            Lote = Val(Mid(MemVar_2.text, 8, 6))
            
            xSql = " SELECT Estado FROM Lotes Where IdLote=" & Lote
            Set xR = Conn.Execute(xSql)
            
            If xR.EOF Then
                MsgBox "Lote NO LOcalizado, Verifique", vbExclamation, "Error de Lote"
                MemVar_2.text = ""
                MemVar_2.Tag = 0
                MemVar_1.SetFocus
                Exit Sub
            ElseIf xR!Estado <> "IN" Then
                MsgBox "Lote Cerrado, Verifique", vbExclamation, "Error de Lote"
                MemVar_2.text = ""
                MemVar_2.Tag = 0
                MemVar_1.SetFocus
                Exit Sub
            End If
               
            MemVar_2.Tag = 1
            If Mid(Transacciones, 1, 1) Then
                LabelLotes.Visible = True
                LabelLotes.Caption = "Datos del Lote"
            End If
            
    Case 3 To 5
            Orden = Val(Mid(MemVar_2.text, 9, 6))
            
            xSql = "Select * from vVentasBascula Where Estado='IN' "
            Set xR = Conn.Execute(xSql)
            
            If xR.EOF Then
                MsgBox "Orden NO Localizada, Verifique", vbExclamation, "Error de Lote"
                MemVar_2.text = ""
                MemVar_2.Tag = 0
                MemVar_1.SetFocus
                Exit Sub
            ElseIf xR!Estado <> "IN" Then
                MsgBox "Oden Cerrada, Verifique", vbExclamation, "Error de Lote"
                MemVar_2.text = ""
                MemVar_2.Tag = 0
                MemVar_1.SetFocus
                Exit Sub
            End If
            MemVar_2.Tag = 1
    End Select
Else
    MemVar_2.Tag = 1
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
Dim xRec As New ADODB.Recordset
Dim xSql As String

On Error GoTo Recover

If MemVar_3 <> "" Then
    MemVar_3 = UCase(MemVar_3)
    Select Case Mid(Transacciones, 1, 1)
        Case 1
            xSql = "Select * From Placas Where Placas='" & MemVar_3 & "'"
            Set xRec = Conn.Execute(xSql)
            If xRec.EOF Then
               MsgBox "NO esta registrada la Placa, Verifique", vbInformation, "Placas"
               MemVar_3.SetFocus
            Else
            
              If Conn.Execute("Select * From vBasculaRecLleno Where Placas='" & MemVar_3 & "'").EOF Then
                   MemVar_4.text = xRec!Conductor
                   MemVar_31 = xRec!IdTransportador
                   Stransportador = MemVar_31
                   SnombreConductor = MemVar_4.text
                   If Conn.Execute("SELECT Nombre FROM Conductores WHERE Cedula='" & MemVar_4 & "'").EOF Then
                        If MsgBox("El Conductor NO esta Creado, desea seguir con la Recepción", vbYesNo, "Crear Conductor") = vbYes Then
                            MemVar_4.text = "00"
                        Else
                            MemVar_4.text = ""
                            MemVar_1.SetFocus
                            Exit Sub
                        End If
                   Else
                   End If
                   LabelConductor = Conn.Execute("SELECT Nombre FROM Conductores WHERE Cedula='" & MemVar_4 & "'").Fields(0)
                   SIdTransportador = xRec!IdTransportador
                   SpesoMAX = xRec!PesoLlenoMAX
                   SpesoMIN = xRec!PesoVacioMAX
                   LabelTransportador.Caption = Conn.Execute("SELECT Descripcion FROM Transportador WHERE IdTransportador=" & SIdTransportador).Fields(0)
                   'Memvar_4.SetFocus
              Else
                   MsgBox "Placa Registrada En Patios, verifique", vbExclamation
                   MemVar_3.text = ""
                   MemVar_3.SetFocus
                   Exit Sub
              End If
            End If
            xRec.Close
        
            If Conn.Execute("Select Tara From Placas Where Placas='" & MemVar_3 & "'").Fields(0) > 0 Then
                UsoTara.Visible = True
            End If
        
        Case 3, 4
            xSql = "Select * From Placas Where Placas='" & MemVar_3 & "'"
            Set xRec = Conn.Execute(xSql)
            If xRec.EOF Then
               MsgBox "NO esta registrada la Placa, Verifique", vbInformation, "Placas"
               MemVar_3 = Splaca
               MemVar_3.SetFocus
               Exit Sub
            Else
            
              If Conn.Execute("Select * From vBasculaDesVacio Where Placas='" & MemVar_3 & "'").EOF Then
                   MemVar_4.text = xRec!Conductor
                   MemVar_31 = xRec!IdTransportador
                   Stransportador = MemVar_31
                   SnombreConductor = MemVar_4.text
                   If Conn.Execute("SELECT Nombre FROM Conductores WHERE Cedula='" & MemVar_4 & "'").EOF Then
                        If MsgBox("El Conductor NO esta Creado, desea seguir con la Recepción", vbYesNo, "Crear Conductor") = vbYes Then
                            MemVar_4.text = "00"
                        Else
                            MemVar_4.text = ""
                            MemVar_1.SetFocus
                            Exit Sub
                        End If
                   Else
                   End If
                   LabelConductor = Conn.Execute("SELECT Nombre FROM Conductores WHERE Cedula='" & MemVar_4 & "'").Fields(0)
                   SIdTransportador = xRec!IdTransportador
                   LabelTransportador.Caption = Conn.Execute("SELECT Descripcion FROM Transportador WHERE IdTransportador=" & SIdTransportador).Fields(0)
                   MemVar_4.SetFocus
              Else
                   MsgBox "Placa Registrada En Patios, verifique", vbExclamation
                   MemVar_3.text = ""
                   MemVar_1.SetFocus
                   Exit Sub
              End If
            End If
            xRec.Close
        
        Case 5
            xSql = "Select * From Placas Where Placas='" & MemVar_3 & "'"
            Set xRec = Conn.Execute(xSql)
            If xRec.EOF Then
               MsgBox "NO esta registrada la Placa, Verifique", vbInformation, "Placas"
               MemVar_3 = Splaca
               MemVar_3.SetFocus
               Exit Sub
            End If
            xRec.Close
    End Select
    MemVar_3.Tag = 1
Else
    MemVar_3.Tag = 0
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer vLotesBascula," & vbCrLf & Err.Description
    MsgBox MSG, , "Memvar_1_LostFocus()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub


Private Sub MemVar_31_GotFocus()
        Call Mark(MemVar_31)
End Sub

Private Sub MemVar_31_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
            Case vbKeyF4
                    Call omenu("Browse")
            Case vbKeyUp
                    MemVar_3.SetFocus
            Case vbKeyDown, vbKeyReturn
                    MemVar_4.SetFocus
    End Select

End Sub

Private Sub MemVar_31_LostFocus()
Dim xR As New ADODB.Recordset

On Error GoTo Recover

If Conn.Execute("Select * From Transportador Where IdTransportador=" & MemVar_31).EOF Then
    MsgBox "Transportador NO localizado, Verifique", vbInformation
    MemVar_31 = Stransportador
    MemVar_31.SetFocus
    Exit Sub
Else
    If MemVar_3 <> "" Then
        If MemVar_31 <> Stransportador Then
            LabelTransportador = Conn.Execute("Select Descripcion From Transportador Where IdTransportador=" & MemVar_31).Fields(0)
            If MsgBox("Esta Seguro del Cambio de Transportador, Será Actualizado ", vbYesNo, "Cambio de Transportador") = vbYes Then
                Conn.Execute ("UPDATE Placas Set IdTransportador=" & MemVar_31 & " WHERE Placas='" & MemVar_3 & "'")
                MemVar_4.SetFocus
            Else
                MemVar_31 = Stransportador
                MemVar_31.SetFocus
            End If
        End If
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al Leer maquina de Cargue," & vbCrLf & Err.Description
    MsgBox MSG, , "Memvar_6_LostFocus()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
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
                    MemVar_5.SetFocus
    End Select
End Sub

Private Sub MemVar_4_LostFocus()
On Error GoTo Recover


If SnombreConductor <> MemVar_4 Then
    If Not Conn.Execute("SELECT Nombre FROM Conductores WHERE Cedula='" & MemVar_4 & "'").EOF Then
        If MsgBox("Esta Seguro del Cambio de Conductor, Será Actualizado ", vbYesNo, "Cambio de Conductor") = vbYes Then
            Conn.Execute ("Update Placas Set Conductor='" & MemVar_4 & "' Where Placas='" & MemVar_3 & "'")
            LabelConductor = Conn.Execute("SELECT Nombre FROM Conductores WHERE Cedula='" & MemVar_4 & "'").Fields(0)
            MsgBox "Nombre Cambiado"
        End If
    Else
        MsgBox "Conductor NO registrado en Base de Datos", vbDefaultButton1, "LogyCoal"
        MemVar_4 = SnombreConductor
        MemVar_4.SetFocus
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al Grabar Nombre," & vbCrLf & Err.Description
    MsgBox MSG, , "Memvar_4_LostFocus()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub

Private Sub MemVar_5_GotFocus()
        Call Mark(MemVar_5)
End Sub

Private Sub MemVar_5_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
    
    Select Case KeyCode
            Case vbKeyF4
                    Call omenu("Browse")
            Case vbKeyUp
                    MemVar_4.SetFocus
            Case vbKeyDown, vbKeyReturn
                    If Mid(Transacciones, 1, 1) <> 3 Then
                       If Mid(Transacciones, 1, 1) = 5 Then
                           MemVar_51.SetFocus
                       Else
                           LeerPeso.SetFocus
                       End If
                    Else
                       MemVar_7 = 0
                       MemVar_8 = 0
                       MemVar_9 = 0
                       Grabar.SetFocus
                     End If
    End Select
End Sub

Private Sub MemVar_5_LostFocus()

If MemVar_5 = "" Then
    MsgBox "Remision NO debe ser un campo vacio", vbInformation
    MemVar_5.Tag = 0
    MemVar_5.SetFocus
Else
    MemVar_5.Tag = 1
End If
End Sub

'Private Sub MemVar_51_KeyDown(KeyCode As Integer, Shift As Integer)
'    Select Case KeyCode
'            Case vbKeyF4
'                    Call omenu("Browse")
'            Case vbKeyUp
'                    MemVar_4.SetFocus
'            Case vbKeyDown, vbKeyReturn
'                    LeerPeso.SetFocus
'    End Select
'End Sub


Private Sub MemVar_6_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyF4
                    Call omenu("Browse")
            Case vbKeyUp
                    MemVar_5.SetFocus
            Case vbKeyDown, vbKeyReturn

    End Select
End Sub

Private Sub MemVar_6_LostFocus()
Dim xRec As New ADODB.Recordset
On Error GoTo Recover

Set xRec = Conn.Execute("Select * From Transportador Where idTransportador=" & MemVar_6)

If Not xRec.EOF Then
   Label_6 = xRec!Descripcion
Else
   MsgBox "Maquina NO Localizada", vbInformation, "Maquina de Cargue"
   MemVar_6.SetFocus
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al Leer maquina de Cargue," & vbCrLf & Err.Description
    MsgBox MSG, , "Memvar_6_LostFocus()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub

Private Sub MemVar_7_GotFocus()
    LeerPeso.SetFocus
End Sub

Private Sub MemVar_8_GotFocus()
    LeerPeso.SetFocus
End Sub

Private Sub MemVar_9_GotFocus()
LeerPeso.SetFocus
    End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
Call omenu(Button.Key)
End Sub

Private Sub omenu(ByVal Key As String)
        Select Case Key
                Case "Grabar"
                        Call SaveData
                Case "Salida"
                        Unload Me
                Case "Foto"
                Case "Browse"
                        Select Case ActiveControl.Name
                        Case "MemVar_3"
                            BrowseCatalogo.x_Tipo = 1
                            BrowseCatalogo.xtabla = "Placas"
                            Set BrowseCatalogo.dControl = ActiveControl
                            BrowseCatalogo.Show 1
                        
                        Case "MemVar_31"
                            BrowseCatalogo.x_Tipo = 1
                            BrowseCatalogo.xtabla = "Transportador"
                            Set BrowseCatalogo.dControl = ActiveControl
                            BrowseCatalogo.Show 1
                
                        Case "MemVar_6"
                            BrowseCatalogo.x_Tipo = 1
                            BrowseCatalogo.xtabla = "Transportador"
                            Set BrowseCatalogo.dControl = ActiveControl
                            BrowseCatalogo.Show 1
                        End Select
                
                Case "Top"
                Case "Bottom"
                Case "Imprimir"
                Case "Borrar"
                Case "Proximo"
                Case "Previo"
        End Select
End Sub

Private Sub Transacciones_Click()
    Call Transacciones_KeyDown(vbKeyReturn, 1)
End Sub

Private Sub Transacciones_GotFocus()
    AnularTiquete.Enabled = False
    MemVar_2 = ""
    MemVar_3 = ""
    MemVar_31 = ""
    MemVar_4 = ""
    MemVar_5 = "0"
    MemVar_51 = ""
    MemVar_6 = ""
    MemVar_7 = 0
    MemVar_8 = 0
    MemVar_9 = 0
    LabelTransportador = ""
    SusoTaraSISMA = False
    LabelLotes.Visible = False
    UsoTara.Visible = False
    LabelConductor = ""
   
End Sub

Private Sub Transacciones_LostFocus()
    T = Mid(Transacciones, 1, 1)
    Select Case T
        Case 1, 2, 4, 5
            Grabar.Visible = False
            ImprimeTiquete.Visible = False
            LeerPeso.Visible = True
        Case 3
            Grabar.Visible = True
            ImprimeTiquete.Visible = False
            LeerPeso.Visible = False
    End Select
End Sub

Private Sub Transacciones_KeyDown(KeyCode As Integer, Shift As Integer)
Dim T As Integer
T = Mid(Transacciones, 1, 1)
UsoTara.Visible = False
Adicionales.Visible = False
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
        Case vbKeyDown, vbKeyReturn
            Call MuestraRecepcion
            If T = 1 Then
                Body.Enabled = True
                MemVar_1.SetFocus
            ElseIf T = 3 Then
                Body.Enabled = True
                MemVar_1.SetFocus
            ElseIf T = 4 Then
                Body.Enabled = True
                MemVar_1.SetFocus
            ElseIf T = 5 Then
                Body.Enabled = True
                Adicionales.Visible = True
                Combo1.ListIndex = 3
                oTree.SetFocus
            Else
                Body.Enabled = False
                oTree.SetFocus
            End If
End Select

End Sub

Private Sub MuestraRecepcion()
Dim Em As New ADODB.Recordset
Dim xSql As String
Dim xSql2 As String
Dim xSql1 As String
Dim xTipo As String

On Error GoTo Recover
oTree.Nodes.Clear

xTipo = Chr(64 + Mid(Transacciones, 1, 1))

Select Case Mid(Transacciones, 1, 1)

Case 1 To 2
    Set Nodx = oTree.Nodes.Add(, , "0" & "Recepciones", "Recepciones", 22)
    
    xSql = "Select Distinct Descripcion From vRecepciones "
    Set Em = Conn.Execute(xSql)

    While Not Em.EOF
        Set Nodx = oTree.Nodes.Add("0" & "Recepciones", tvwChild, "A" & Em!Descripcion, Em!Descripcion, 16)
        Em.MoveNext
    Wend
    
    xSql = "SELECT  * FROM vRecepciones"
    Set Em = Conn.Execute(xSql)
    
    While Not Em.EOF
        Set Nodx = oTree.Nodes.Add("A" & Em!Descripcion, tvwChild, xTipo & Em!IdTiquete, Format(Em!IdTiquete, "000000") & "- " & Em!Placas & " - " & Em!Conductor, 16)
        Em.MoveNext
    Wend
    
Case 3 To 5
    Set Nodx = oTree.Nodes.Add(, , "0" & "Despachos", "Despachos", 22)
    
    Set Nodx = oTree.Nodes.Add("0" & "Despachos", tvwChild, "0" & "Ventas", "Ventas", 22)
    Set Nodx = oTree.Nodes.Add("0" & "Despachos", tvwChild, "0" & "Traslados", "Traslados", 22)
    
    Select Case Mid(Transacciones, 1, 1)
        Case 3 '==> Enturnado de Mulas
             xSql2 = "Select Distinct Descripcion From vDespachos WHERE PesoVacio=0"
             xSql = "SELECT  * FROM vDespachos WHERE PesoVacio=0"
             Set Em = Conn.Execute(xSql2)
            
             While Not Em.EOF
                 Set Nodx = oTree.Nodes.Add("0" & "Ventas", tvwChild, "A" & Em!Descripcion, Em!Descripcion, 29)
                 Em.MoveNext
             Wend
             
             ' ==> TRASLADOS
             xSql2 = "Select Distinct Descripcion From vDespachos2 WHERE PesoVacio=0"
             xSql1 = "SELECT  * FROM vDespachos2 WHERE PesoVacio=0"
             Set Em = Conn.Execute(xSql2)
            
             While Not Em.EOF
                 Set Nodx = oTree.Nodes.Add("0" & "Traslados", tvwChild, "A" & Em!Descripcion, Em!Descripcion, 29)
                 Em.MoveNext
             Wend
        
        Case 4 '==> Peso Vacio de Mulas
             xSql2 = "Select Distinct Descripcion From vDespachos WHERE PesoVacio>=0"
             xSql = "SELECT  * FROM vDespachos WHERE PesoVacio>=0"
             Set Em = Conn.Execute(xSql2)
            
             While Not Em.EOF
                 Set Nodx = oTree.Nodes.Add("0" & "Ventas", tvwChild, "A" & Em!Descripcion, Em!Descripcion, 29)
                 Em.MoveNext
             Wend
             
             ' ==> TRASLADOS
             xSql2 = "Select Distinct Descripcion From vDespachos2 WHERE PesoVacio>=0"
             xSql1 = "SELECT  * FROM vDespachos2 WHERE PesoVacio>=0"
             Set Em = Conn.Execute(xSql2)
            
             While Not Em.EOF
                 Set Nodx = oTree.Nodes.Add("0" & "Traslados", tvwChild, "A" & Em!Descripcion, Em!Descripcion, 29)
                 Em.MoveNext
             Wend
            
        Case 5 '==> PESO LLENO
             xSql2 = "Select Distinct Descripcion From vDespachos WHERE PesoVacio>0"
             xSql = "SELECT  * FROM vDespachos WHERE PesoVacio>0"
             Set Em = Conn.Execute(xSql2)
            
             While Not Em.EOF
                 Set Nodx = oTree.Nodes.Add("0" & "Ventas", tvwChild, "A" & Em!Descripcion, Em!Descripcion, 29)
                 Em.MoveNext
             Wend
             ' ==> TRASLADOS
             xSql2 = "Select Distinct Descripcion From vDespachos2 WHERE PesoVacio>0"
             xSql1 = "SELECT  * FROM vDespachos2 WHERE PesoVacio>0"
             Set Em = Conn.Execute(xSql2)
            
             While Not Em.EOF
                 Set Nodx = oTree.Nodes.Add("0" & "Traslados", tvwChild, "A" & Em!Descripcion, Em!Descripcion, 29)
                 Em.MoveNext
             Wend
    End Select

    Set Em = Conn.Execute(xSql)
    While Not Em.EOF
        Set Nodx = oTree.Nodes.Add("A" & Em!Descripcion, tvwChild, xTipo & Em!IdTiquete, Format(Em!IdTiquete, "000000") & "- " & Em!Placas & " - " & Em!Conductor, 29)
        Em.MoveNext
    Wend
    
    Set Em = Conn.Execute(xSql1)
    While Not Em.EOF
        Set Nodx = oTree.Nodes.Add("A" & Em!Descripcion, tvwChild, xTipo & Em!IdTiquete, Format(Em!IdTiquete, "000000") & "- " & Em!Placas & " - " & Em!Conductor, 29)
        Em.MoveNext
    Wend
    
    Em.Close
End Select

For Each loNode In oTree.Nodes
    If loNode.children >= 1 Then
        loNode.Expanded = True
    End If
Next

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al Mostrar Recepciones," & vbCrLf & Err.Description
    MsgBox MSG, , "Muestra Recepciones"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub

Private Sub oTree_NodeClick(ByVal Node As MSComctlLib.Node)

Select Case Mid(Node.Key, 1, 1)
Case "B"
     Body.Enabled = True
     SnumeroTiquete = Val(Mid(Node.Key, 2, 6))
     Call CargaTiquete(SnumeroTiquete)
     AnularTiquete.Enabled = True

Case "C"
     Body.Enabled = True
     SnumeroTiquete = Val(Mid(Node.Key, 2, 6))
     Call CargaTiquete(SnumeroTiquete)
     AnularTiquete.Enabled = True

Case "D"
    Body.Enabled = True
    SnumeroTiquete = Val(Mid(Node.Key, 2, 6))
    Call CargaTiquete(SnumeroTiquete)
     AnularTiquete.Enabled = True

Case "E"
     Body.Enabled = True
     SnumeroTiquete = Val(Mid(Node.Key, 2, 6))
     Call CargaTiquete(SnumeroTiquete)
     AnularTiquete.Enabled = True
End Select

End Sub

Private Sub CargaTiquete(Numero)
Dim xR As New ADODB.Recordset
Dim xL As New ADODB.Recordset
Dim xNeto As Single

On Error Resume Next
If Mid(Transacciones, 1, 1) = 3 Then Exit Sub

Set xR = Conn.Execute("Select * From Bascula Where IdTiquete=" & Numero)
If xR.EOF Then
    MsgBox "Se presento un error al leer tiquete No " + Str(Numero) + ", Verifique", vbInformation
    Exit Sub
End If

MemVar_1 = Format(xR!IdMaterial, "00") + " " + Conn.Execute("Select Descripcion From Materiales Where IdMaterial=" & xR!IdMaterial).Fields(0)

Select Case Mid(Transacciones, 1, 1)
    Case 2
            If xR!IdMaterial = 1 Then
                 xSql = "Select * from vLotesBascula Where IdLote=" & xR!NumeroTransaccion
                 Set xL = Conn.Execute(xSql)
                 If Not xL.EOF Then
                    MemVar_2.text = "Lote : " + Format(xL!IdLote, "000000") + "  " + xL!Mina + "  " + xL!DesTipo
                 End If
                 xL.Close
            Else
                MemVar_2.text = "Otros Materiales"
            End If
    Case 3 To 5
            If xR!IdMaterial = 1 Then
                Select Case xR!TransaccionOrigen
                Case "DS"
                        xSql = "Select * from vVentasBascula Where IdVentas=" & xR!NumeroTransaccion
                        Set xL = Conn.Execute(xSql)
                        If xL.EOF Then
                           MsgBox "NO hay Oden de ventas Abiertas, Verifique"
                           MemVar_1.SetFocus
                        Else
                           MemVar_2.text = "Orden : " + Format(xL!IdVentas, "000000") + "  " + xL!Descripcion
                           Do While Not xL.EOF
                                  MemVar_2.text = "Orden : " + Format(xL!IdVentas, "000000") + "  " + xL!Descripcion
                                  xL.MoveNext
                           Loop
                        End If
                        xL.Close
                Case "TR"
                        xSql = "Select * from vTrasladosBascula Where IdTraslado=" & xR!NumeroTransaccion
                        Set xL = Conn.Execute(xSql)
                        If xL.EOF Then
                           MsgBox "NO hay Oden de ventas Abiertas, Verifique"
                           MemVar_1.SetFocus
                        Else
                           MemVar_2.text = "Tras. : " + Format(xL!IdTraslado, "000000") + "  " + xL!Desacopio
                           Do While Not xL.EOF
                                  MemVar_2.text = "Tras. : " + Format(xL!IdTraslado, "000000") + "  " + xL!Desacopio
                                  xL.MoveNext
                           Loop
                        End If
                        xL.Close
                
                End Select
            End If
End Select

MemVar_3.text = xR!Placas
Splaca = MemVar_3
SpesoMIN = Conn.Execute("Select PesoVacioMAX From Placas Where Placas='" & MemVar_3 + "'").Fields(0)
SpesoMAX = Conn.Execute("Select PesoLlenoMAX From Placas Where Placas='" & MemVar_3 + "'").Fields(0)
SIdTransportador = xR!IdTransportador
LabelTransportador.Caption = Conn.Execute("SELECT Descripcion FROM Transportador WHERE IdTransportador=" & SIdTransportador).Fields(0)
MemVar_31 = xR!IdTransportador
Stransportador = MemVar_31
MemVar_4.text = xR!Conductor
SnombreConductor = MemVar_4
 LabelConductor = Conn.Execute("SELECT Nombre FROM Conductores WHERE Cedula='" & MemVar_4 & "'").Fields(0)
MemVar_5.text = xR!Documentoasociado
MemVar_6.text = xR!MaquinaCargue
MemVar_7.text = xR!PesoLleno
MemVar_8.text = xR!PesoVacio

xNeto = xR!PesoLleno - xR!PesoVacio
MemVar_9.text = IIf(xNeto < 0, 0, xNeto)

MemVar_1.Tag = 1
MemVar_2.Tag = 1
MemVar_3.Tag = 1
MemVar_4.Tag = 1
MemVar_5.Tag = 1
MemVar_6.Tag = 1
     
If Mid(Transacciones, 1, 1) = 4 Then
    If MemVar_8 <= 0 Then
            LeerPeso.Visible = True
            Grabar.Visible = True
            LeerPeso.SetFocus
    Else
            LeerPeso.Visible = False
            Grabar.Visible = False
    End If
Else
    LeerPeso.Visible = True
    Grabar.Visible = True
    LeerPeso.SetFocus
End If

xR.Close

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al Cargar Tiquiete -CargaTiquete(Numero)," & vbCrLf & Err.Description
    MsgBox MSG, , "Cargar Tiquete"
    Err.Clear   ' Borra campos del objeto Err
    Exit Sub
End If

End Sub

Private Sub SaveData()
Dim xSql As String
Dim Tipo As Integer
Dim NumTra As Integer
Dim Lote As Integer
Dim Material As Integer
Dim Res As String
Dim TiempoTran As Integer
Dim TipoTran As String

Dim lngRec As Long

On Error GoTo Recover
   
MemVar_6.text = 1 '==> Default para ID Equipo de Cargue
   
TiempoTran = Mid(Transacciones, 1, 1)
Material = Val(Mid(MemVar_1.text, 1, 2))
TipoTran = IIf(Mid(MemVar_1, 1, 2) = "01", "LT", "RO")

Select Case TiempoTran
         Case 1                '==>  Recepcion Peso Lleno
                NumTra = 1 ' ==>Transaccion Recepcion
                Lote = Val(Mid(MemVar_2.text, 9, 6))
                xSql = " EXEC PA_PesoInicial " & TiempoTran & ", " & NumTra & ", '" & TipoTran & "', " & Lote & ", " & MemVar_5.text & ", " & Material & ", " & SIdTransportador & ", '" & MemVar_3 & "', '" & MemVar_4 & "',"
                xSql = xSql & MemVar_6 & ", " & MemVar_7 & ",'" & Susuario & "', 'IN'"
                If SusoTaraSISMA Then
                    xSql = xSql & "," & MemVar_8 & ",'" & MemVar_51 & "'"
                Else
                    xSql = xSql & ",0,'" & MemVar_51 & "'"
                End If
                
               Res = Conn.Execute(xSql).Fields(0)
                
                If Res <> "OK" Then
                    MsgBox "Error al Grabar el Registro, Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
                End If
         
         Case 2               '==>  Recepcion Peso Vacio
                xSql = " EXEC PA_PesoFinal " & TiempoTran & ", " & SnumeroTiquete & ", " & MemVar_8.text & ",'" & MemVar_51 & "','" & MemVar_4 & "',0,0"
                
                Res = Conn.Execute(xSql).Fields(0)
                
                If Res <> "OK" Then
                    MsgBox "Error al Grabar el Registro, Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
                End If
                
                Conn.Execute ("UPDATE Placas Set Tara=" & MemVar_8 & " WHERE Placas='" & MemVar_3 & "'")
                
         Case 3                '==>  Turnarse a Cola de Espera
                NumTra = 2 ' ==>Transaccion Despacho
                MemVar_6.text = 1
                MemVar_7.text = 0
                Lote = Val(Mid(MemVar_2.text, 9, 6))
                If Mid(MemVar_2.text, 1, 5) = "Orden" Then
                    TipoTran = "DS"
                Else
                    TipoTran = "TR"
                End If
                xSql = " EXEC PA_PesoInicial " & TiempoTran & ", " & NumTra & ", '" & TipoTran & "', " & Lote & ",'" & MemVar_5 & "', " & Material & ", " & SIdTransportador & ", '" & MemVar_3 & "', '" & MemVar_4 & "',"
                xSql = xSql & MemVar_6 & ", " & MemVar_8 & ",'" & Susuario & "', 'IN',0,'" & MemVar_51 & "'"
                
                Res = Conn.Execute(xSql).Fields(0)
                
                If Res <> "OK" Then
                    MsgBox "Error al Grabar el Registro, Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
                End If
                
         Case 4        '==>  Despachos Peso Vacio

                NumTra = 2 ' ==>Transaccion Despacho
                MemVar_6.text = 1
                MemVar_7.text = 0
                Lote = Val(Mid(MemVar_2.text, 9, 6))
                If Mid(MemVar_2.text, 1, 5) = "Orden" Then
                    TipoTran = "DS"
                Else
                    TipoTran = "TR"
                End If
                
                 If Conn.Execute("Select * From vBasculaDesVacio Where Placas='" & MemVar_3 & "'").EOF Then
                    xSql = " EXEC PA_PesoInicial " & TiempoTran & ", " & NumTra & ", '" & TipoTran & "', " & Lote & ",'" & MemVar_5 & "', " & Material & ", " & SIdTransportador & ", '" & MemVar_3 & "', '" & MemVar_4 & "',"
                    xSql = xSql & MemVar_6 & ", " & MemVar_8 & ", '" & Susuario & "', 'IN',0,'" & MemVar_51 & "'"
                    
                    Res = Conn.Execute(xSql).Fields(0)
                    
                    If Res <> "OK" Then
                        MsgBox "Error al Grabar el Registro, Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
                    End If
                 Else
                    xSql = " EXEC PA_PesoFinal " & TiempoTran & ", " & SnumeroTiquete & ", " & MemVar_8.text & ",'" & MemVar_51 & "','" & MemVar_4 & "',0,0"
                    Res = Conn.Execute(xSql).Fields(0)
                    If Res <> "OK" Then
                        MsgBox "Error al Grabar el Registro, Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
                    End If
                 End If
                Conn.Execute ("UPDATE Placas Set Tara=" & MemVar_8 & " WHERE Placas='" & MemVar_3 & "'")
                
         Case 5        '==>  Despachos Peso Lleno
                xSql = " EXEC PA_PesoFinal " & TiempoTran & ", " & SnumeroTiquete & ", " & MemVar_7.text & ",'" & MemVar_51 & "','" & MemVar_4 & "'," & Val(Mid(Combo1, 1, 2)) & "," & Carpado.Value
                Res = Conn.Execute(xSql).Fields(0)
                If Res <> "OK" Then
                    MsgBox "Error al Grabar el Registro, Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
                End If
 End Select

Call MuestraRecepcion

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al Grabar Tiquete," & vbCrLf & Err.Description
    MsgBox MSG, , "Savedata"
    Err.Clear   ' Borra campos del objeto Err
    Exit Sub
End If
End Sub

Private Sub CargaMateriales()
Dim xRec As New ADODB.Recordset
Dim xSql As String

On Error GoTo Recover

MemVar_1.Clear
xSql = "Select * From Materiales"
Set xRec = Conn.Execute(xSql)
If xRec.EOF Then
   MsgBox "NO hay Materiales Descritos, Verifique"
   Unload Me
Else
   MemVar_1.text = Format(xRec!IdMaterial, "00") + "  " + xRec!Descripcion
   Do While Not xRec.EOF
          MemVar_1.AddItem Format(xRec!IdMaterial, "00") + "  " + xRec!Descripcion
          xRec.MoveNext
   Loop
End If
xRec.Close

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer Materiales," & vbCrLf & Err.Description
    MsgBox MSG, , "CargaMateriales()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
                dControl.SelStart = 0
                dControl.SelLength = Len(dControl.text)
End If
End Sub

Private Sub UsoTara_Click()
Dim xSql As String

If MemVar_8 = 0 Then
    xSql = "Select Tara From Placas Where Placas='" & MemVar_3 & "' "
    MemVar_8 = Conn.Execute(xSql).Fields(0)
    'UsoTara.Visible = False
    SusoTaraSISMA = True
Else
    MemVar_8 = 0
        SusoTaraSISMA = False
End If
MemVar_9 = MemVar_7.text - MemVar_8.text
Grabar.SetFocus
End Sub

Private Sub LabelLotes_PostPaint(ByVal HDC As Long)
    Dim mTop As Long, TextHeight As Long
    Dim lWidth As Long
    Dim lMargin As Long
    Dim i As Integer
    Dim xR As New ADODB.Recordset
    
    On Error Resume Next
    
    mTop = 30   '100 - .BackColorOpacity / 1.5
    lMargin = 5
    
    Set xR = Conn.Execute("Select * From vLotesBascula Where IdLote='" & Val(Mid(MemVar_2, 8, 6)) & "'")
    
    With LabelLotes

        lWidth = ((.Width / .GetWindowsDPI / Screen.TwipsPerPixelX) - lMargin * 2)    '100= aproximate height
                                                                 
        TextHeight = .DrawText(HDC, sTitle, lMargin, mTop, lWidth, 200, mFont2, vbWhite, 100, ccEnter, cTop, True)
        TextHeight = TextHeight + .DrawText(HDC, "Mina      : " + xR!Mina, lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H80000012, 100, cLeft, cTop, True)
        TextHeight = TextHeight + .DrawText(HDC, "Pit          :  " + xR!Tajo, lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H80000012, 100, cLeft, cTop, True)
        TextHeight = TextHeight + .DrawText(HDC, "Manto    :  " + xR!Manto, lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H80000012, 100, cLeft, cTop, True)
        TextHeight = TextHeight + .DrawText(HDC, "Pila        :  " + xR!Descripcion, lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H80000012, 100, cLeft, cTop, True)
       TextHeight = TextHeight + .DrawText(HDC, "Oprador :  " + xR!DesOperador, lMargin, mTop + TextHeight, lWidth, 100, mFont2, &H80000012, 100, cLeft, cTop, True)
       
    End With
    xR.Close
End Sub

