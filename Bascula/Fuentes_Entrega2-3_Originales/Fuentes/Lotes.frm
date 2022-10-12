VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Lotes 
   Caption         =   "Ordenes de Recepción"
   ClientHeight    =   8610
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13440
   LinkTopic       =   "Form1"
   ScaleHeight     =   8610
   ScaleWidth      =   13440
   StartUpPosition =   1  'CenterOwner
   Begin MBSplit.Splitter VSplitter 
      Height          =   7335
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   13275
      _ExtentX        =   23416
      _ExtentY        =   12938
      SplitterPos     =   30
      Begin MSComctlLib.TreeView oTree 
         Height          =   3735
         Left            =   120
         TabIndex        =   19
         Top             =   120
         Width           =   3195
         _ExtentX        =   5636
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
         Height          =   7155
         Left            =   4680
         TabIndex        =   11
         Top             =   60
         Width           =   7695
         Begin VB.TextBox MemVar_7 
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
            Left            =   1980
            MaxLength       =   15
            TabIndex        =   10
            ToolTipText     =   "Pulse ""F4"" para Buscar el Proveedor"
            Top             =   4080
            Width           =   600
         End
         Begin Threed.SSFrame SSFrame1 
            Height          =   2235
            Left            =   420
            TabIndex        =   33
            Top             =   1260
            Width           =   6795
            _Version        =   65536
            _ExtentX        =   11986
            _ExtentY        =   3942
            _StockProps     =   14
            Caption         =   "Ubicación"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Begin VB.TextBox MemVar_53 
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
               Left            =   1500
               MaxLength       =   15
               TabIndex        =   8
               ToolTipText     =   "Pulse ""F4"" para Buscar el Proveedor"
               Top             =   1620
               Width           =   600
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
               Height          =   330
               Left            =   1500
               MaxLength       =   15
               TabIndex        =   3
               Top             =   1080
               Width           =   915
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
               Left            =   1500
               MaxLength       =   15
               TabIndex        =   2
               ToolTipText     =   "Pulse ""F4"" para Buscar el Proveedor"
               Top             =   420
               Width           =   600
            End
            Begin VB.ComboBox MemVar_5 
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
               Left            =   3480
               TabIndex        =   5
               Top             =   1080
               Width           =   1155
            End
            Begin MSMask.MaskEdBox MemVar_4 
               Height          =   330
               Left            =   2520
               TabIndex        =   4
               ToolTipText     =   "Se Digita el Valor de la Cotizacion"
               Top             =   1080
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   15
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
            Begin MSMask.MaskEdBox MemVar_51 
               Height          =   330
               Left            =   4680
               TabIndex        =   6
               ToolTipText     =   "Se Digita el Valor de la Cotizacion"
               Top             =   1080
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   15
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
            Begin MSMask.MaskEdBox MemVar_52 
               Height          =   330
               Left            =   5640
               TabIndex        =   7
               ToolTipText     =   "Se Digita el Valor de la Cotizacion"
               Top             =   1080
               Width           =   900
               _ExtentX        =   1588
               _ExtentY        =   582
               _Version        =   393216
               MaxLength       =   15
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
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Operador"
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
               Left            =   120
               TabIndex        =   42
               Top             =   1620
               Width           =   690
            End
            Begin VB.Label Label53 
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
               Left            =   2220
               TabIndex        =   41
               Top             =   1620
               Width           =   4275
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Bloque"
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
               Left            =   5880
               TabIndex        =   40
               Top             =   780
               Width           =   495
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Pit"
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
               Left            =   4920
               TabIndex        =   39
               Top             =   780
               Width           =   165
            End
            Begin VB.Label Label11 
               Caption         =   "Nivel"
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
               Left            =   1560
               TabIndex        =   38
               Top             =   780
               Width           =   615
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               Caption         =   "Mina o Proveedor"
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
               Left            =   120
               TabIndex        =   37
               Top             =   420
               Width           =   1260
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Panel"
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
               Left            =   2760
               TabIndex        =   36
               Top             =   780
               Width           =   390
            End
            Begin VB.Label Label5 
               Caption         =   "Manto"
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
               Left            =   3840
               TabIndex        =   35
               Top             =   780
               Width           =   555
            End
            Begin VB.Label LabelMina 
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
               Left            =   2220
               TabIndex        =   34
               Top             =   420
               Width           =   2775
            End
         End
         Begin MSMask.MaskEdBox CantidadRecibidaBascula 
            Height          =   315
            Left            =   1980
            TabIndex        =   31
            Top             =   4560
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
            Format          =   "###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSComCtl2.DTPicker Fecha 
            Height          =   315
            Left            =   1920
            TabIndex        =   25
            Top             =   840
            Width           =   1335
            _ExtentX        =   2355
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
            Format          =   116260865
            CurrentDate     =   44578
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
            Left            =   1920
            MaxLength       =   15
            TabIndex        =   1
            Top             =   420
            Width           =   840
         End
         Begin VB.PictureBox okFind 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   2940
            Picture         =   "Lotes.frx":0000
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   13
            Top             =   420
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
            Left            =   2940
            Picture         =   "Lotes.frx":0102
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   12
            Top             =   420
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
                     Picture         =   "Lotes.frx":0204
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Lotes.frx":051E
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Lotes.frx":0838
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Lotes.frx":0B52
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Lotes.frx":0E6C
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Lotes.frx":1186
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Lotes.frx":14A0
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Lotes.frx":17BA
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Lotes.frx":1AD4
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Lotes.frx":1DEE
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Lotes.frx":2108
                     Key             =   ""
                  EndProperty
               EndProperty
            End
         End
         Begin VB.TextBox MemVar_6 
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
            Left            =   1980
            TabIndex        =   9
            Top             =   3660
            Width           =   855
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1125
            Left            =   360
            TabIndex        =   28
            Top             =   5580
            Width           =   7290
            _ExtentX        =   12859
            _ExtentY        =   1984
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483624
            BorderStyle     =   1
            Appearance      =   0
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
               Text            =   "Muestra"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Tipo Muestra"
               Object.Width           =   2117
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Fecha Muestra"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               SubItemIndex    =   3
               Text            =   "Cantidad"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Fecha Entrega"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSComCtl2.DTPicker FechaCierre 
            Height          =   315
            Left            =   4920
            TabIndex        =   46
            Top             =   840
            Visible         =   0   'False
            Width           =   2235
            _ExtentX        =   3942
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
            Format          =   116260865
            CurrentDate     =   44578
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
            Left            =   4260
            TabIndex        =   49
            Top             =   420
            Visible         =   0   'False
            Width           =   1275
         End
         Begin Tracer.LabelPlus Cerrado 
            Height          =   675
            Left            =   3360
            TabIndex        =   48
            Top             =   240
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
            Caption         =   "Lotes.frx":2422
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
            PictureArr      =   "Lotes.frx":2442
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
            Left            =   3480
            TabIndex        =   47
            Top             =   840
            Visible         =   0   'False
            Width           =   930
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Apertura"
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
            Left            =   480
            TabIndex        =   45
            Top             =   840
            Width           =   1140
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Pila"
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
            TabIndex        =   44
            Top             =   4080
            Width           =   240
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
            Left            =   3000
            TabIndex        =   43
            Top             =   4080
            Width           =   2775
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Ton."
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
            Left            =   3540
            TabIndex        =   32
            Top             =   4620
            Width           =   315
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Muestras de Laboratorio"
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
            TabIndex        =   30
            Top             =   5220
            Width           =   1770
         End
         Begin VB.Label Label7 
            Caption         =   "Cantidad Recibida en Bascula"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Index           =   0
            Left            =   420
            TabIndex        =   29
            Top             =   4500
            Width           =   1425
         End
         Begin VB.Label SubLabel1 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3600
            TabIndex        =   17
            Top             =   480
            Width           =   3255
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
            Left            =   480
            TabIndex        =   16
            Top             =   480
            Width           =   1275
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Carbon"
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
            Left            =   420
            TabIndex        =   15
            Top             =   3720
            Width           =   870
         End
         Begin VB.Label LabelTipoCarbon 
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
            Left            =   3000
            TabIndex        =   14
            Top             =   3660
            Width           =   2775
         End
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
            Picture         =   "Lotes.frx":38D8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":3A32
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":3B8C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":3CE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":3E40
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":3F9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":40F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":424E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":43A8
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":4502
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":465C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":47B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":51C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":5762
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Lotes.frx":5CFC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComctlLib.StatusBar oBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   18
      Top             =   8295
      Width           =   13440
      _ExtentX        =   23707
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
            Text            =   "OT USADA"
            TextSave        =   "OT USADA"
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
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   780
      Left            =   0
      TabIndex        =   20
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
      MinHeight4      =   315
      Width4          =   1695
      NewRow4         =   0   'False
      MinHeight5      =   360
      Width5          =   2400
      NewRow5         =   0   'False
      MinHeight6      =   360
      Width6          =   2400
      NewRow6         =   0   'False
      MinHeight7      =   360
      Width7          =   1500
      NewRow7         =   0   'False
      Begin KewlButtonz.KewlButtons MuestraProduccion 
         Height          =   315
         Left            =   9060
         TabIndex        =   27
         Top             =   405
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Muestra de Producción"
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
         MICON           =   "Lotes.frx":6296
         PICN            =   "Lotes.frx":62B2
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons MuestraCalidad 
         Height          =   315
         Left            =   6660
         TabIndex        =   26
         Top             =   405
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Muestra de Calidad"
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
         MICON           =   "Lotes.frx":6610
         PICN            =   "Lotes.frx":662C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons CierreLotes 
         Height          =   315
         Left            =   4860
         TabIndex        =   24
         Top             =   405
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Cerrar Lotes"
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
         MICON           =   "Lotes.frx":697E
         PICN            =   "Lotes.frx":699A
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
         ItemData        =   "Lotes.frx":6E70
         Left            =   825
         List            =   "Lotes.frx":6E80
         TabIndex        =   23
         Text            =   "TODAS"
         Top             =   405
         Width           =   2355
      End
      Begin KewlButtonz.KewlButtons oNuevo 
         Height          =   315
         Left            =   3405
         TabIndex        =   22
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
         MICON           =   "Lotes.frx":6EB5
         PICN            =   "Lotes.frx":6ED1
         PICH            =   "Lotes.frx":746B
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
         TabIndex        =   21
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
Attribute VB_Name = "Lotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Lotes As New ADODB.Recordset
Dim okUnload As Boolean
Dim IsNewRecord As Boolean
Dim OkOpen As Boolean

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub CierreLotes_Click()
Dim xSql As String
Dim xRes As String

 If MsgBox("Esta Seguro de Cerrar  el Lote " + Chr(13) + Chr(10) + "No se podran hacer mas Recepciones en Bascula con este Lote", vbYesNo, "Cierre de Lotes") = vbYes Then
    xSql = " EXEC PA_Cierrelotes " & MemVar_1 & ", " & Susuario
    xRes = Conn.Execute(xSql).Fields(0)
    If xRes <> "OK" Then
        MsgBox vbCrLf & xRes
        Exit Sub
    End If
    xSql = " EXEC PA_CargarLotePilaDetalle  " & MemVar_1
    xRes = Conn.Execute(xSql).Fields(0)
    If xRes <> "OK" Then
        MsgBox vbCrLf & xRes
        Exit Sub
    End If
    Call MemVar_1_LostFocus
End If

End Sub

Private Sub Form_Activate()
If Not OkOpen Then
     Call LoadData
     OkOpen = True
    Me.SetFocus
    Call MuestraArbol
     MemVar_1.SetFocus
End If
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
OkOpen = False
End Sub

Private Sub Form_Load()
Dim xCt As New ADODB.Recordset
Lotes.Open "Lotes", Conn, 2, 3, 512
If Not Lotes.EOF Then Lotes.MoveLast
Set VSplitter.LeftOrTopCtl = oTree
Set VSplitter.RightOrBottomCtl = Frame1
End Sub

Private Sub MuestraCalidad_Click()
Dim xSql As String
Dim xRes As String

If MsgBox("Esta Seguro de Crear Muestra de Calidad ", vbYesNo, "Crear Muestras de Calidad") = vbYes Then
    xSql = " EXEC PA_CreaRegCalidad " & "'LT', " & MemVar_1 & ", 1, 0,'" & Susuario & "'"
    xRes = Conn.Execute(xSql).Fields(0)
    If xRes <> "OK" Then
        MsgBox xRes, vbCritical, "Error de Creacion de Muestra"
    Else
        Call CargaMuestrasCalidad(MemVar_1)
        MsgBox "Muestra de Calidad Creada"
    End If
End If
End Sub

Private Sub MuestraProduccion_Click()
Dim xSql As String
Dim xRes As String

If MsgBox("Esta Seguro de Crear Muestra de Producción ", vbYesNo, "Crear Muestras de Producción") = vbYes Then
    xSql = " EXEC PA_CreaRegCalidad " & "'LT', " & MemVar_1 & ", 2, " & CantidadRecibidaBascula & ",'" & Susuario & "'"
    xRes = Conn.Execute(xSql).Fields(0)
    If xRes <> "OK" Then
        MsgBox xRes, vbCritical, "Error de Creacion de Muestra"
    Else
        Call CargaMuestrasCalidad(MemVar_1)
        MsgBox "Muestra de Producción Creada"
    End If
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
                        Case "MemVar_2"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "Minas"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        Case "MemVar_53"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "OperadoresMineros"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        Case "MemVar_6"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "TiposCarbon"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        Case "MemVar_7"
                               BrowseAcopioPilas.x_Tipo = 1
                                BrowseAcopioPilas.xtabla = "vPilasAcopiosGeneral"
                                Set BrowseAcopioPilas.dControl = ActiveControl
                                BrowseAcopioPilas.Show 1
                        End Select
                Case "Imprime"
                    
                Case "Borrar"
                    If MsgBox("Esta seguro de Borrar el Lote", vbYesNo, "Borrado de Registro") = vbYes Then
                        If CantidadRecibidaBascula = 0 Then
                            Conn.Execute ("Update Lotes Set Estado='AN' Where IdLote=" & MemVar_1)
                            Call MySeek(Lotes, Conn, "Lotes", "IdLote=" & MemVar_1)
                            Call LoadData
                        Else
                            MsgBox "NO puede borrar el lote porque presenta Cantidad Recibida", vbInformation
                            MemVar_1.SetFocus
                            Exit Sub
                        End If
                    End If
Salida:
                Case "Top"
                    Lotes.Close
                    Lotes.Open "Select Top 1 * From Lotes Order By IdLote"
                    Call LoadData
                Case "Bottom"
                    Lotes.Close
                    Lotes.Open "Select Top 1 * From Lotes Order By IdLote DESC"
                    Call LoadData
                Case "Proximo"
                    Lotes.Close
                    Lotes.Open "Select Top 1 * From Lotes Where IdLote>'" & MemVar_1 & "' Order By IdLote"
                    Call LoadData
                Case "Previo"
                    Lotes.Close
                    Lotes.Open "Select Top 1 * From Lotes Where  IdLote<'" & MemVar_1 & "' Order By IdLote DESC"
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

Private Sub ExeBrow(oCod As String)
' Codigo Por desarrollar
End Sub

Private Sub LoadData()
On Error GoTo Recover

 If Lotes.EOF And Lotes.BOF Then
'         okNew.Visible = True
'         okFind.Visible = False
         Exit Sub
 End If

 If Not Lotes.EOF Then
     okNew.Visible = False
     okFind.Visible = True

     MemVar_1 = Lotes!IdLote
     Fecha = Lotes!Fechaapertura
     If IsNull(Lotes!FechaCierre) Then
        FechaCierre.Visible = False
        Label1(2).Visible = False
     Else
        Label1(2).Visible = True
        FechaCierre.Visible = True
        FechaCierre = Lotes!FechaCierre
     End If
     
     MemVar_2 = Lotes!IdMina
     LabelMina.Caption = Conn.Execute("Select Descripcion From Minas Where IdMina=" & MemVar_2).Fields(0)
     MemVar_3 = Lotes!Nivel
     MemVar_4 = Lotes!Panel
     MemVar_5 = Lotes!Manto
     MemVar_51 = Lotes!Tajo
     MemVar_52 = Lotes!Bloque
     MemVar_53 = Lotes!Operador
     MemVar_6 = Lotes!IdTipoCarbon
     MemVar_7 = Lotes!Pila
     Label53 = Conn.Execute("Select Descripcion From OperadoresMineros Where IdOperador=" & MemVar_53).Fields(0)
     LabelPila = Conn.Execute("Select Descripcion From vPilas Where IdPilA=" & MemVar_7).Fields(0)
     LabelTipoCarbon.Caption = Conn.Execute("Select Descripcion From TiposCarbon Where iDTipoCarbon=" & MemVar_6).Fields(0)
     CantidadRecibidaBascula = Lotes!Cantidad
 Else
 
 End If

 CierreLotes.Enabled = (Lotes!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','CierreLotes'").Fields(0)) '==> Fields(0) Nos dice si tiene acceso al Objeto

 Tbar.buttons("Grabar").Enabled = (Lotes!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','" & sProgSISMA & "'").Fields(1)) '==> Fields(1) Nos dice puede hacer cambios al Objeto
 Tbar.buttons("Borrar").Enabled = (Lotes!Estado = "IN" And CantidadRecibidaBascula = 0 And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','" & sProgSISMA & "'").Fields(1)) '==> Fields(1) Nos dice puede hacer cambios al Objeto

 Anulado.Visible = (Lotes!Estado = "AN")
 oBar.Panels("Usuario").text = "USUARIO: " & Lotes!Usuario & " " & Format(Lotes!Fechaapertura, "dd/MM/yyyy hh:mm")
 oBar.Panels("Ot").text = "ESTADO: " & Lotes!Estado

Cerrado.Visible = (Lotes!Estado <> "IN")

Call CargaMuestrasCalidad(MemVar_1)

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

        If IsNewRecord Then
            xSql = "EXEC PA_OrdenRecepcion " & MemVar_2 & ",'" & MemVar_3 & "','" & MemVar_4 & "','" & MemVar_5 & "'," & MemVar_6 & ",'" & Susuario & "','" & MemVar_51 & "','" & MemVar_52 & "'," & MemVar_53 & "," & MemVar_7
            
            Res = Conn.Execute(xSql).Fields(0)
             If Res <> "OK" Then
                 MsgBox "Error al Grabar Orden de Recepcion, Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
                 Exit Sub
             End If
            
        Else
            Lotes!IdMina = MemVar_2.text
            Lotes!Nivel = MemVar_3.text
            Lotes!Panel = MemVar_4.text
            Lotes!Manto = MemVar_5.text
            Lotes!Tajo = MemVar_51
            Lotes!Bloque = MemVar_52
            Lotes!Operador = MemVar_53
            Lotes!Pila = MemVar_7
            Lotes!IdTipoCarbon = MemVar_6.text
            Lotes.Update
        End If
        
        If IsNewRecord Then
            MemVar_1 = Conn.Execute("SELECT IdLote FROM Lotes Where Estado='IN' AND Usuario='" & Susuario & "' ORDER BY IdLote DESC").Fields(0)
            If Not MySeek(Lotes, Conn, "Lotes", "idLote=" & MemVar_1) Then
                     Call LoadData
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

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
                dControl.SelStart = 0
                dControl.SelLength = Len(dControl.text)
End If
End Sub

Private Function Validate(X As String, Y As Byte) As Boolean
        Validate = True
End Function

Private Sub MemVar_1_GotFocus()
        Call Mark(MemVar_1)
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

On Error GoTo Recover

If MemVar_1 <> "" Then
    If Not MySeek(Lotes, Conn, "Lotes", "IdLote=" & MemVar_1) Then
            Call LoadData
    Else
            MsgBox "Numero de Lote NO Localizado"
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
                        Call omenu("Browse")
                Case vbKeyUp
                        MemVar_1.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_3.SetFocus
End Select
End Sub

Private Sub MemVar_2_LostFocus()
Dim xR As New ADODB.Recordset

On Error GoTo Recover

If MemVar_2.text <> "" Then
        Set xR = Conn.Execute(" Select * From Minas Where iDMina=" & MemVar_2)
        If Not xR.EOF Then
                LabelMina.Caption = xR!Descripcion
        Else
                MsgBox "NO se localiza el ID de Mina"
                LabelMina.Caption = ""
                MemVar_2.text = ""
                MemVar_2.SetFocus
        End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Mina," & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_2_LostFocus()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub

Private Sub MemVar_3_GotFocus()
        Call Mark(MemVar_3)
End Sub

Private Sub MemVar_3_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                Case vbKeyUp
                        xNoCot.SetFocus
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
                Case vbKeyUp
                        MemVar_3.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_5.SetFocus
End Select
End Sub

Private Sub MemVar_5_GotFocus()
        Call Mark(MemVar_5)
End Sub

Private Sub MemVar_5_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                        Call omenu("Browse")
                Case vbKeyUp
                        MemVar_4.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_51.SetFocus
End Select
End Sub

Private Sub MemVar_51_GotFocus()
        Call Mark(MemVar_51)
End Sub

Private Sub MemVar_51_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                        Call omenu("Browse")
                Case vbKeyUp
                        MemVar_5.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_52.SetFocus
End Select
End Sub

Private Sub MemVar_52_GotFocus()
        Call Mark(MemVar_52)
End Sub

Private Sub MemVar_52_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                        Call omenu("Browse")
                Case vbKeyUp
                        MemVar_51.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_53.SetFocus
End Select
End Sub

Private Sub MemVar_53_GotFocus()
        Call Mark(MemVar_53)
End Sub

Private Sub MemVar_53_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                        Call omenu("Browse")
                Case vbKeyUp
                        MemVar_52.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_6.SetFocus
End Select
End Sub

Private Sub MemVar_53_LostFocus()
Dim xR As New ADODB.Recordset

On Error GoTo Recover

If MemVar_53.text <> "" Then
        Set xR = Conn.Execute(" Select * From  OperadoresMineros Where IdOperador=" & MemVar_53)
        If Not xR.EOF Then
                Label53.Caption = xR!Descripcion
        Else
                MsgBox "NO se localiza el ID del operador Minero"
                Label53 = ""
                MemVar_53.text = ""
                MemVar_53.SetFocus
        End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Operador Minero," & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_53_LostFocus()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub

Private Sub MemVar_6_GotFocus()
        Call Mark(MemVar_6)
End Sub

Private Sub MemVar_6_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                        Call omenu("Browse")
                Case vbKeyUp
                        MemVar_5.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_7.SetFocus
End Select
End Sub

Private Sub MemVar_6_LostFocus()
Dim xR As New ADODB.Recordset

On Error GoTo Recover

If MemVar_6.text <> "" Then
        Set xR = Conn.Execute(" Select * From  TiposCarbon Where iDtipoCarbon=" & MemVar_6)
        If Not xR.EOF Then
                LabelTipoCarbon.Caption = xR!Descripcion
        Else
                MsgBox "NO se localiza el ID del Tipo de Carbon"
                LabelTipoCarbon = ""
                MemVar_6.text = ""
                MemVar_6.SetFocus
        End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Tipo Carbon," & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_6_LostFocus()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub

Private Sub MemVar_7_GotFocus()
        Call Mark(MemVar_7)
End Sub

Private Sub MemVar_7_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                        Call omenu("Browse")
                Case vbKeyUp
                        MemVar_6.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_1.SetFocus
End Select
End Sub

Private Sub MemVar_7_LostFocus()
Dim xR As New ADODB.Recordset

On Error GoTo Recover

If MemVar_7.text <> "" Then
        Set xR = Conn.Execute("Select * From  vPilas Where IdPila=" & MemVar_7)
        If Not xR.EOF Then
                LabelPila.Caption = xR!Descripcion
        Else
                MsgBox "NO se localiza el ID de la Pila"
                LabelPila = ""
                MemVar_7.text = ""
                MemVar_7.SetFocus
        End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga la Pila de carbon," & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_7_LostFocus()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub


Private Sub MuestraArbol()
Dim Em As New ADODB.Recordset
Dim xSql As String

On Error GoTo Recover

oTree.Nodes.Clear
Set Nodx = oTree.Nodes.Add(, , "Lotes", "Minas", 13)

xSql = "SELECT  Distinct(Mina) FROM  vLotesBascula"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("Lotes", tvwChild, "A" & Em!Mina, Em!Mina, 14)
    Em.MoveNext
Wend

xSql = "SELECT  *   FROM  vLotesBascula Where Estado='IN' "
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("A" & Em!Mina, tvwChild, "A" & Format(Em!IdLote, "000000"), Format(Em!IdLote, "000000") & " " & Em!DesTipo & " " & Format(Em!Cantidad, "###,##0" & " Ton"), 14)
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

Private Sub oTree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer

Select Case Mid(Node.Key, 1, 1)
Case "A"
    CargaLote (Mid(Node.Key, 2, 6))
Case "B"
End Select

End Sub

Private Sub oNuevo_Click()
If oNuevo.Caption = "Nuevo" Then
    oNuevo.Caption = "Cancelar"
    Tbar.buttons("Grabar").Enabled = True
    Tbar.buttons("Borrar").Enabled = True
    IsNewRecord = True
    MemVar_1.text = ""
    MemVar_2.text = ""
    MemVar_3.text = ""
    MemVar_4.text = ""
    MemVar_5.text = ""
    MemVar_6.text = ""
    MemVar_51 = ""
    MemVar_52 = ""
    MemVar_53 = ""
    MemVar_7 = ""
    Fecha = Now
    MemVar_2.SetFocus
    Label53 = ""
    LabelMina = ""
    LabelTipoCarbon = ""
    LabelPila = ""
    CantidadRecibidaBascula = 0
    ListView1.ListItems.Clear
    Label1(2).Visible = False
    FechaCierre.Visible = False
    oBar.Panels("Usuario").text = "USUARIO: " & Susuario & " " & Format(Now, "dd/MM/yyyy hh:mm")
    oBar.Panels("Ot").text = "ESTADO: " & "IN"
Else
    oNuevo.Caption = "Nuevo"
    IsNewRecord = False
    Lotes.MoveLast
    MemVar_1 = Lotes!IdLote
    Call MemVar_1_LostFocus
End If
MemVar_2.SetFocus
End Sub

Private Sub CargaLote(Numero)
Dim xSql As String

If Not MySeek(Lotes, Conn, "Lotes", "IdLote=" & Numero) Then
    MemVar_1 = Lotes!IdLote
    Call LoadData
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Lote desde el Arbol," & vbCrLf & Err.Description
    MsgBox MSG, , "CargaLote"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub

Private Sub CargaMuestrasCalidad(Lote As Integer)
Dim xSql As String
Dim xTm As New ADODB.Recordset

On Error GoTo Recover

    xSql = " SELECT * FROM  vFrmLotes_LotesCalidad WHERE iDLote=" & Lote
    
    Set xTm = Conn.Execute(xSql)
    ListView1.ListItems.Clear
    Do While Not xTm.EOF
        Set iTmx = ListView1.ListItems.Add()
        iTmx.text = xTm!IdMuestra
        iTmx.SubItems(1) = xTm!DesMuestras
        iTmx.SubItems(2) = Format(xTm!Fecha, "dd/MM/yyyy hh:mm")
        iTmx.SubItems(3) = Format(xTm!Cantidad, "####,##0")
        iTmx.SubItems(4) = IIf(IsNull(xTm!FechaEntrega), "", Format(xTm!FechaEntrega, "dd/MM/yyyy hh:mm"))
        xTm.MoveNext
    Loop
    xTm.Close

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer vFrmLotes_LotesCalidad," & vbCrLf & Err.Description
    MsgBox MSG, , "CargaMuestrasCalidad(Lote As Integer)"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub
