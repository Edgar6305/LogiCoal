VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Pilas 
   Caption         =   "Pilas"
   ClientHeight    =   9750
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17400
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
   ScaleHeight     =   9750
   ScaleWidth      =   17400
   StartUpPosition =   1  'CenterOwner
   Begin MBSplit.Splitter VSplitter 
      Height          =   8175
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   17115
      _ExtentX        =   30189
      _ExtentY        =   14420
      SplitterPos     =   25
      Begin MSComctlLib.TreeView oTree 
         Height          =   7995
         Left            =   0
         TabIndex        =   22
         Top             =   0
         Width           =   4215
         _ExtentX        =   7435
         _ExtentY        =   14102
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
         Height          =   7935
         Left            =   4440
         TabIndex        =   1
         Top             =   60
         Width           =   12795
         Begin VB.TextBox MemVar_21 
            Height          =   315
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   58
            Top             =   1140
            Width           =   1260
         End
         Begin VB.TextBox MemVar_2 
            Height          =   315
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   34
            Top             =   720
            Width           =   1260
         End
         Begin VB.TextBox MemVar_1 
            Height          =   315
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   4
            Top             =   360
            Width           =   1260
         End
         Begin VB.PictureBox okFind 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
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
            Height          =   240
            Left            =   3180
            Picture         =   "Pilas.frx":0000
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   3
            Top             =   360
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
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   240
            Left            =   3180
            Picture         =   "Pilas.frx":0102
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   2
            Top             =   360
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
                     Picture         =   "Pilas.frx":0204
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Pilas.frx":051E
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Pilas.frx":0838
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Pilas.frx":0B52
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Pilas.frx":0E6C
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Pilas.frx":1186
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Pilas.frx":14A0
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Pilas.frx":17BA
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Pilas.frx":1AD4
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Pilas.frx":1DEE
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Pilas.frx":2108
                     Key             =   ""
                  EndProperty
               EndProperty
            End
         End
         Begin MSComCtl2.DTPicker FechaInicial 
            Height          =   315
            Left            =   1800
            TabIndex        =   5
            Top             =   1560
            Width           =   1275
            _ExtentX        =   2249
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
            Format          =   117243905
            CurrentDate     =   44578
         End
         Begin Threed.SSFrame oMarco 
            Height          =   2310
            Left            =   360
            TabIndex        =   6
            Top             =   3060
            Width           =   11895
            _Version        =   65536
            _ExtentX        =   20981
            _ExtentY        =   4075
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
            Begin VB.TextBox MemVar_7 
               Alignment       =   2  'Center
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   3200
               MaxLength       =   25
               TabIndex        =   32
               Top             =   450
               Width           =   975
            End
            Begin VB.TextBox MemVar_5 
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   1140
               MaxLength       =   25
               TabIndex        =   9
               Top             =   450
               Width           =   1035
            End
            Begin VB.VScrollBar oMove 
               Height          =   2100
               Left            =   11580
               TabIndex        =   8
               Top             =   120
               Width           =   255
            End
            Begin VB.TextBox MemVar_4 
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   315
               Index           =   0
               Left            =   40
               MaxLength       =   10
               TabIndex        =   7
               Top             =   450
               Width           =   1080
            End
            Begin MSMask.MaskEdBox MemVar_6 
               Height          =   315
               Index           =   0
               Left            =   2160
               TabIndex        =   10
               Top             =   450
               Width           =   1020
               _ExtentX        =   1799
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_8 
               Height          =   315
               Index           =   0
               Left            =   4200
               TabIndex        =   35
               Top             =   450
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_9 
               Height          =   315
               Index           =   0
               Left            =   4860
               TabIndex        =   36
               Top             =   450
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_10 
               Height          =   315
               Index           =   0
               Left            =   5520
               TabIndex        =   37
               Top             =   450
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_11 
               Height          =   315
               Index           =   0
               Left            =   6180
               TabIndex        =   38
               Top             =   450
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_12 
               Height          =   315
               Index           =   0
               Left            =   6840
               TabIndex        =   39
               Top             =   450
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_13 
               Height          =   315
               Index           =   0
               Left            =   7500
               TabIndex        =   40
               Top             =   450
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_14 
               Height          =   315
               Index           =   0
               Left            =   8160
               TabIndex        =   41
               Top             =   450
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_15 
               Height          =   315
               Index           =   0
               Left            =   8820
               TabIndex        =   42
               Top             =   450
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_16 
               Height          =   315
               Index           =   0
               Left            =   9540
               TabIndex        =   43
               Top             =   450
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_17 
               Height          =   315
               Index           =   0
               Left            =   10275
               TabIndex        =   44
               Top             =   450
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_18 
               Height          =   315
               Index           =   0
               Left            =   10920
               TabIndex        =   45
               Top             =   450
               Width           =   645
               _ExtentX        =   1138
               _ExtentY        =   556
               _Version        =   393216
               MaxLength       =   10
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "Segoe UI"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Format          =   "#,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dry-V"
               Height          =   300
               Index           =   13
               Left            =   10920
               TabIndex        =   57
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dry-V"
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
               Index           =   12
               Left            =   10920
               TabIndex        =   56
               Top             =   450
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dry-CF"
               Height          =   300
               Index           =   11
               Left            =   10275
               TabIndex        =   55
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dry-CNZ"
               Height          =   300
               Index           =   10
               Left            =   9540
               TabIndex        =   54
               Top             =   120
               Width           =   705
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dry-BTU"
               Height          =   300
               Index           =   9
               Left            =   8820
               TabIndex        =   53
               Top             =   120
               Width           =   705
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dry-S"
               Height          =   300
               Index           =   8
               Left            =   8160
               TabIndex        =   52
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "V"
               Height          =   300
               Index           =   7
               Left            =   7500
               TabIndex        =   51
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "CF"
               Height          =   300
               Index           =   6
               Left            =   6840
               TabIndex        =   50
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "CNZ"
               Height          =   300
               Index           =   5
               Left            =   6180
               TabIndex        =   49
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "BTU"
               Height          =   300
               Index           =   4
               Left            =   5520
               TabIndex        =   48
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "S"
               Height          =   300
               Index           =   3
               Left            =   4860
               TabIndex        =   47
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "H"
               Height          =   300
               Index           =   2
               Left            =   4200
               TabIndex        =   46
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "No Muestra"
               Height          =   300
               Index           =   1
               Left            =   3180
               TabIndex        =   31
               Top             =   120
               Width           =   1000
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cantidad"
               Height          =   300
               Index           =   0
               Left            =   2160
               TabIndex        =   13
               Top             =   120
               Width           =   1005
            End
            Begin VB.Label SubLabel3 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Numero"
               Height          =   285
               Index           =   0
               Left            =   1140
               TabIndex        =   12
               Top             =   120
               Width           =   1020
            End
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Transaccion"
               Height          =   285
               Left            =   45
               TabIndex        =   11
               Top             =   120
               Width           =   1050
            End
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1845
            Left            =   420
            TabIndex        =   14
            Top             =   6000
            Width           =   11895
            _ExtentX        =   20981
            _ExtentY        =   3254
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
            NumItems        =   7
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Muestra"
               Object.Width           =   1587
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Origen"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   2
               Text            =   "Numero"
               Object.Width           =   1411
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Tipo Muestra"
               Object.Width           =   2646
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   4
               Text            =   "Cantidad"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   5
               Text            =   "Fecha Muestra"
               Object.Width           =   2540
            EndProperty
            BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   6
               Text            =   "Fecha Entrega"
               Object.Width           =   2540
            EndProperty
         End
         Begin MSMask.MaskEdBox MemVar_3 
            Height          =   315
            Left            =   1800
            TabIndex        =   15
            Top             =   2220
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   10380
            Top             =   360
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
                  Picture         =   "Pilas.frx":2422
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":257C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":26D6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":2830
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":298A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":2AE4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":2C3E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":2D98
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":2EF2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":304C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":31A6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":3300
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":3D12
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":42AC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":4846
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":4DE0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":62EA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":6E34
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":73ED
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":8503
                  Key             =   ""
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":8CD5
                  Key             =   ""
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":9DA7
                  Key             =   ""
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":A8F1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Pilas.frx":BA47
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker FechaCierre 
            Height          =   315
            Left            =   4680
            TabIndex        =   30
            Top             =   1560
            Visible         =   0   'False
            Width           =   1275
            _ExtentX        =   2249
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
            Format          =   117243905
            CurrentDate     =   44578
         End
         Begin Tracer.LabelPlus Cerrado 
            Height          =   555
            Left            =   3540
            TabIndex        =   62
            Top             =   180
            Visible         =   0   'False
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   979
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
            Caption         =   "Pilas.frx":C05D
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
            PictureArr      =   "Pilas.frx":C07D
         End
         Begin VB.Label Label11 
            Caption         =   "Toneladas"
            Height          =   315
            Index           =   1
            Left            =   3300
            TabIndex        =   61
            Top             =   2280
            Width           =   1035
         End
         Begin VB.Label LabelAcopio 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3180
            TabIndex        =   60
            Top             =   1140
            Width           =   4815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Ocopio"
            Height          =   195
            Index           =   4
            Left            =   420
            TabIndex        =   59
            Top             =   1200
            Width           =   570
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Pila Fisica"
            Height          =   195
            Index           =   3
            Left            =   420
            TabIndex        =   33
            Top             =   780
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Cierre"
            Height          =   195
            Index           =   2
            Left            =   3300
            TabIndex        =   29
            Top             =   1620
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Apertura"
            Height          =   195
            Index           =   1
            Left            =   420
            TabIndex        =   21
            Top             =   1620
            Width           =   1170
         End
         Begin VB.Label LabelPilaFisica 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3180
            TabIndex        =   20
            Top             =   720
            Width           =   4815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Numero de Pila"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   19
            Top             =   420
            Width           =   1170
         End
         Begin VB.Label Label11 
            Caption         =   "Cantidad "
            Height          =   315
            Index           =   0
            Left            =   420
            TabIndex        =   18
            Top             =   2280
            Width           =   1035
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
            TabIndex        =   17
            Top             =   5700
            Width           =   1770
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Detalle de la Pila"
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
            TabIndex        =   16
            Top             =   2880
            Width           =   1155
         End
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   780
      Left            =   0
      TabIndex        =   23
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
      Begin KewlButtonz.KewlButtons CierrePilas 
         Height          =   315
         Left            =   4860
         TabIndex        =   27
         Top             =   405
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Cierre de Pilas"
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
         MICON           =   "Pilas.frx":D513
         PICN            =   "Pilas.frx":D52F
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
         ItemData        =   "Pilas.frx":DA05
         Left            =   825
         List            =   "Pilas.frx":DA15
         TabIndex        =   26
         Text            =   "TODAS"
         Top             =   405
         Width           =   2355
      End
      Begin KewlButtonz.KewlButtons oNuevo 
         Height          =   315
         Left            =   3405
         TabIndex        =   25
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
         MICON           =   "Pilas.frx":DA42
         PICN            =   "Pilas.frx":DA5E
         PICH            =   "Pilas.frx":DFF8
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
         TabIndex        =   24
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
   Begin ComctlLib.StatusBar oBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   28
      Top             =   9435
      Width           =   17400
      _ExtentX        =   30692
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
End
Attribute VB_Name = "Pilas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim okUnload As Boolean
Const CONTSTOP = 30
Const maxView = 5
Const SpaceY = 300
Dim IsNewRecord As Boolean
Dim MaxUser As Integer
Dim OkOpen As Boolean
Dim oD As New ADODB.Recordset

Private Sub CierrePilas_Click()
Dim xSql As String

xSql = "UPDATE Pilas Set Estado='AC' WHERE idPila=" & MemVar_1
If MsgBox("Esta seguro de Cerrar la Pila " + LabelPilaFisica + " " + LabelAcopio, vbYesNo, "Cerrado de Pila") = vbYes Then
    Conn.Execute (xSql)
End If
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
End If
Me.SetFocus
Call MuestraArbol
End Sub

Private Sub Form_Load()
Dim X As New ADODB.Recordset
    Set VSplitter.LeftOrTopCtl = oTree
    Set VSplitter.RightOrBottomCtl = Body
    
    oD.Open "Select * From Pilas", Conn, adOpenDynamic, adLockOptimistic, 1
    If Not oD.EOF Then oD.MoveLast
End Sub

Private Sub LoadControls()
Dim j As Integer
    For j = 1 To MaxUser
            Load MemVar_4(j)
            Load MemVar_5(j)
            Load MemVar_6(j)
            Load MemVar_7(j)
            Load MemVar_8(j)
            Load MemVar_9(j)
            Load MemVar_10(j)
            Load MemVar_11(j)
            Load MemVar_12(j)
            Load MemVar_13(j)
            Load MemVar_14(j)
            Load MemVar_15(j)
            Load MemVar_16(j)
            Load MemVar_17(j)
            Load MemVar_18(j)
    Next j
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
                        MemVar_7(i).Top = j * SpaceY + MemVar_7(0).Top
                        MemVar_7(i).Visible = True
                        MemVar_8(i).Top = j * SpaceY + MemVar_8(0).Top
                        MemVar_8(i).Visible = True
                        MemVar_9(i).Top = j * SpaceY + MemVar_9(0).Top
                        MemVar_9(i).Visible = True
                        MemVar_10(i).Top = j * SpaceY + MemVar_10(0).Top
                        MemVar_10(i).Visible = True
                        MemVar_11(i).Top = j * SpaceY + MemVar_11(0).Top
                        MemVar_11(i).Visible = True
                        MemVar_12(i).Top = j * SpaceY + MemVar_12(0).Top
                        MemVar_12(i).Visible = True
                        MemVar_13(i).Top = j * SpaceY + MemVar_13(0).Top
                        MemVar_13(i).Visible = True
                        MemVar_14(i).Top = j * SpaceY + MemVar_14(0).Top
                        MemVar_14(i).Visible = True
                        MemVar_15(i).Top = j * SpaceY + MemVar_15(0).Top
                        MemVar_15(i).Visible = True
                        MemVar_16(i).Top = j * SpaceY + MemVar_16(0).Top
                        MemVar_16(i).Visible = True
                        MemVar_17(i).Top = j * SpaceY + MemVar_17(0).Top
                        MemVar_17(i).Visible = True
                        MemVar_18(i).Top = j * SpaceY + MemVar_18(0).Top
                        MemVar_18(i).Visible = True
                        j = j + 1
                Else
                        MemVar_4(i).Visible = False
                        MemVar_5(i).Visible = False
                        MemVar_6(i).Visible = False
                        MemVar_7(i).Visible = False
                        MemVar_8(i).Visible = False
                        MemVar_9(i).Visible = False
                        MemVar_10(i).Visible = False
                        MemVar_11(i).Visible = False
                        MemVar_12(i).Visible = False
                        MemVar_13(i).Visible = False
                        MemVar_14(i).Visible = False
                        MemVar_15(i).Visible = False
                        MemVar_16(i).Visible = False
                        MemVar_17(i).Visible = False
                        MemVar_18(i).Visible = False
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
    MemVar_7(i).text = ""
    MemVar_8(i).text = ""
    MemVar_9(i).text = ""
    MemVar_10(i).text = ""
    MemVar_11(i).text = ""
    MemVar_12(i).text = ""
    MemVar_13(i).text = ""
    MemVar_14(i).text = ""
    MemVar_15(i).text = ""
    MemVar_16(i).text = ""
    MemVar_17(i).text = ""
    MemVar_18(i).text = ""
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

xSql = "SELECT TOP 31 * From PilasDetalle Where IdPila=" & MemVar_1 & " Order By Transaccion, Numero DESC"
Set xR = Conn.Execute(xSql)

Do While Not xR.EOF
    MemVar_4(i).text = xR!Transaccion
    MemVar_5(i).text = xR!Numero
    MemVar_6(i).text = xR!Cantidad
    MemVar_7(i).text = xR!IdMuestra
    MemVar_8(i).text = xR!H
    MemVar_9(i).text = xR!S
    MemVar_10(i).text = xR!BTU
    MemVar_11(i).text = xR!CNZ
    MemVar_12(i).text = xR!CF
    MemVar_13(i).text = xR!v
    MemVar_14(i).text = xR!Dry_S
    MemVar_15(i).text = xR!Dry_BTU
    MemVar_16(i).text = xR!Dry_CNZ
    MemVar_17(i).text = xR!Dry_CF
    MemVar_18(i).text = xR!Dry_V
    i = i + 1
    xR.MoveNext
Loop

Call AjustaMover
Call wShow

End Sub

Private Sub MemVar_4_GotFocus(Index As Integer)
        Call Mark(MemVar_4(Index))
End Sub

Private Sub MemVar_4_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                        Call omenu("Browse")
                Case vbKeyUp
                        MemVar_31.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_6(0).SetFocus
End Select
End Sub

Private Sub MemVar_4_LostFocus(Index As Integer)
Dim xR As New ADODB.Recordset


Set xR = Conn.Execute("Select * From vPilasGeneral Where IdPila=" & MemVar_4(Index).text)

If Not xR.EOF Then
    MemVar_5(Index) = xR!Descripcion
    MemVar_6(Index).SetFocus
Else
    MsgBox "Pila NO localizada, verifique", vbInformation
    MemVar_4(Index).SetFocus
End If

End Sub

Private Sub MemVar_6_GotFocus(Index As Integer)
        Call Mark(MemVar_6(Index))
End Sub

Private Sub oMove_Change()
        Call wShow
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
        Call omenu(Button.Key)
End Sub

Private Sub omenu(ByVal Key As String)
'On Error GoTo Recover
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
                                BrowseCatalogo.xtabla = "Pilas"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        Case "MemVar_2"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "PilasFisicas"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        Case "MemVar_21"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "Acopios"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        
                        End Select
                Case "Imprime"

                Case "Borrar"
                On Error GoTo BorraRec
                If MsgBox("Esta seguro de Borrar el Registro " + Chr(13) + Chr(10) + "Las relaciones si las hay seran borradas", vbYesNo, "Borrado de Registro") = vbYes Then
                    'oT.Delete
                    oT.MoveFirst
                    Call LoadData
                End If
                GoTo Salida
BorraRec:
                MsgBox "Error al borrar el registro, existe relaciones con otros archivos las cuales " + Chr(13) + Chr(10) + "             no permiten que se borre este registro, verifique"
                On Error GoTo 0
Salida:
                Case "Top"
                    oD.Close
                    oD.Open "Select Top 1 * From Pilas Order By IdPila"
                    Call LoadData
                Case "Bottom"
                    oD.Close
                    oD.Open "Select Top 1 * From Pilas Order By IdPila DESC"
                    Call LoadData
                Case "Proximo"
                    oD.Close
                    oD.Open "Select Top 1 * From Pilas Where IdPila>'" & MemVar_1 & "' Order By IdPila"
                    Call LoadData
                Case "Previo"
                    oD.Close
                    oD.Open "Select Top 1 * From Pilas Where  IdPila<'" & MemVar_1 & "' Order By IdPila DESC"
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

 If oD.EOF And oD.BOF Then
'         okNew.Visible = True
'         okFind.Visible = False
         Exit Sub
 End If

MemVar_1 = oD!IdPila
MemVar_2 = oD!IDPilaFisica
LabelPilaFisica = Conn.Execute(" SELECT Descripcion From PilasFisicas Where IdPilaFisica=" & MemVar_2).Fields(0)
MemVar_21 = oD!IdAcopio
LabelAcopio = Conn.Execute(" SELECT Descripcion  From Acopios Where IdAcopio=" & MemVar_21).Fields(0)

Fechainicio = oD!Fechainicio
FehaCierre = oD!FechaCierre

MemVar_3 = oD!Cantidad

oBar.Panels("Usuario").text = "USUARIO: " & oD!Usuario
oBar.Panels("Ot").text = "ESTADO: " & oD!Estado

CierrePilas.Enabled = (oD!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','CierrePilas'").Fields(0)) '==> Fields(0) Nos dice si tiene acceso al Objeto

Tbar.buttons("Grabar").Enabled = (oD!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','" & sProgSISMA & "'").Fields(1)) '==> Fields(1) Nos dice puede hacer cambios al Objeto
Tbar.buttons("Borrar").Enabled = (oD!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','" & sProgSISMA & "'").Fields(1)) '==> Fields(1) Nos dice puede hacer cambios al Objeto

Call LoadValores
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
Dim xSql As String
Dim xRes As String

'On Error GoTo Recover

     If IsNewRecord Then
        xSql = "EXEC PA_Pilas " & MemVar_2 & "," & MemVar_21 & ",'" & Susuario & "'"
        Res = Conn.Execute(xSql).Fields(0)
         
         If Res <> "OK" Then
             MsgBox "Error al Grabar Pila, Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
             Exit Sub
         End If
     Else
         xSql = "UPDATE Pilas Set  IdPilaFisica = " & MemVar_2 & ",IdAcopio =" & MemVar_21 & " WHERE IdPila=" & MemVar_1
         Conn.Execute (xSql)
     End If
     
     If IsNewRecord Then
        MemVar_1 = Conn.Execute(" Select IdPila From Pilas Where Estado= 'IN' And Usuario='" & Susuario & "' Order By IdPila DESC").Fields(0)
        Call MuestraArbol
        Call MemVar_1_LostFocus
     End If
     oNuevo.Caption = "Nuevo"
     IsNewRecord = False

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
    MemVar_3.text = ""
    MemVar_21.text = ""
    FechaInicial = Now
    For i = 0 To CONTSTOP - 1
        Call Limpia(i)
    Next i
    LabelPilaFisica = ""
    LabelAcopio = ""
    ListView1.ListItems.Clear
    oBar.Panels("Usuario").text = "USUARIO: " & Susuario & " " & Format(Now, "dd/MM/yyyy hh:mm")
    oBar.Panels("Ot").text = "ESTADO: " & "IN"
Else
    oNuevo.Caption = "Nuevo"
    IsNewRecord = False
    oD.MoveLast
    MemVar_1 = oD!IdPila
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
                        Call omenu("Browse")
                Case vbKeyUp
                Case vbKeyDown, vbKeyReturn
                        MemVar_2.SetFocus
End Select
End Sub

Private Sub MemVar_1_LostFocus()

On Error GoTo Recover

 If Not MySeek(oD, Conn, "Pilas", "IdPila=" & MemVar_1) Then
        Call LoadData
 Else
        MsgBox "Numero de Orden NO Localizada"
        MemVar_1.SetFocus
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
                Case vbKeyDown, vbKeyReturn
                        MemVar_21.SetFocus
End Select
End Sub

Private Sub MemVar_2_LostFocus()
Dim xR As New ADODB.Recordset

On Error GoTo Recover
If MemVar_2 <> "" Then
    Set xR = Conn.Execute("SELECT * From PilasFisicas Where IdPilaFisica=" & MemVar_2)
    
    If Not xR.EOF Then
        LabelPilaFisica = xR!Descripcion
    Else
        MsgBox "Pila NO Localizado, Verifique", vbInformation
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Cargar Pila Fisica" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_2_LostFocus()"
    Err.Clear
    Resume Next
End If
    
End Sub

Private Sub MemVar_21_GotFocus()
        Call Mark(MemVar_21)
End Sub

Private Sub MemVar_21_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                        Call omenu("Browse")
                Case vbKeyUp
                        MemVar_2.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_2.SetFocus
End Select
End Sub

Private Sub MemVar_21_LostFocus()
Dim xR As New ADODB.Recordset

On Error GoTo Recover
If MemVar_21 <> "" Then
    Set xR = Conn.Execute("SELECT * From Acopios Where IdAcopio=" & MemVar_21)
    
    If Not xR.EOF Then
        LabelAcopio = xR!Descripcion
    Else
        MsgBox "Acopio NO Localizado, Verifique", vbInformation
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Cargar Acopio" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_21_LostFocus()"
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
Set Nodx = oTree.Nodes.Add(, , "A" & "Acopios", "Acopios", 24)

xSql = "Select Distinct Ubicacion From Acopios"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("A" & "Acopios", tvwChild, "B" & Em!Ubicacion, "ACOPIOS " & Em!Ubicacion, 23)
    Em.MoveNext
Wend

xSql = "Select Distinct Descripcion, Ubicacion, IdAcopio From Acopios"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("B" & Em!Ubicacion, tvwChild, "C" & Em!Ubicacion & Em!Descripcion, Em!IdAcopio & " " & Em!Descripcion, 22)
    Em.MoveNext
Wend

xSql = "SELECT  Acopios.Ubicacion, Pilas.IdPila, PilasFisicas.Descripcion, PilasFisicas.TipoCarbon, Acopios.Descripcion AS DesAcopio"
xSql = xSql + " FROM      Acopios INNER JOIN Pilas ON Acopios.IdAcopio = Pilas.IdAcopio INNER JOIN"
xSql = xSql + "                  PilasFisicas ON Pilas.IdPilaFisica = PilasFisicas.IdPilaFisica"
xSql = xSql + " WHERE  Pilas.Estado='IN'"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("C" & Em!Ubicacion & Em!Desacopio, tvwChild, "D" & Em!IdPila, Format(Em!IdPila, "00") & " - " & Em!Descripcion, 21)
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
Case "D"
    MemVar_1 = Mid(Node.Key, 2, 6)
    Call MemVar_1_LostFocus
End Select

End Sub

Private Sub CargaOrden(ByVal Num As Long)
Dim xSql As String

On Error GoTo Recover
If Not MySeek(oD, Conn, "Ventas", "IdVentas=" & Numero) Then
    MemVar_1 = oD!IdVentas
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

Private Sub CargaMuestrasCalidad(ByVal Despacho As Long)
Dim xSql As String
Dim xTm As New ADODB.Recordset
Dim i As Integer

On Error GoTo Recover
 ListView1.ListItems.Clear

For i = 0 To CONTSTOP
    If MemVar_4(i) <> "" Then
    
        xSql = " SELECT  Calidad.IdMuestra, Calidad.TransaccionOrigen, Calidad.Numero, Calidad.TipoMuestra, Calidad.Cantidad,"
        xSql = xSql + "       Calidad.Fecha, Calidad.Usuario, Calidad.H, Calidad.S, Calidad.BTU, Calidad.CNZ, Calidad.CF, Calidad.V,"
        xSql = xSql + "       Calidad.Dry_S, Calidad.Dry_BTU, Calidad.Dry_CNZ, Calidad.Dry_CF, Calidad.Dry_V, Calidad.UsuarioEntrega,"
        xSql = xSql + "       Calidad.FechaEntrega , Calidad.Estado, TipoMuestras.Descripcion"
        xSql = xSql + " FROM       Calidad INNER JOIN TipoMuestras ON Calidad.TipoMuestra = TipoMuestras.IdTipoMuestra"
        xSql = xSql + " Where      Calidad.IdMuestra = " & MemVar_7(i)

        Set xTm = Conn.Execute(xSql)
        
        If xTm.EOF Then
            Set iTmx = ListView1.ListItems.Add()
            iTmx.text = "No Data..."
        End If
        Do While Not xTm.EOF
            Set iTmx = ListView1.ListItems.Add()
            iTmx.text = xTm!IdMuestra
            iTmx.SubItems(1) = xTm!TransaccionOrigen
            iTmx.SubItems(2) = xTm!Numero
            iTmx.SubItems(3) = xTm!Descripcion
            iTmx.SubItems(4) = Format(xTm!Cantidad, "####,##0")
            iTmx.SubItems(5) = Format(xTm!Fecha, "dd/MM/yyyy hh:mm")
            iTmx.SubItems(6) = IIf(IsNull(xTm!FechaEntrega), "", Format(xTm!FechaEntrega, "dd/MM/yyyy hh:mm"))
            xTm.MoveNext
        Loop
        xTm.Close
    Else
    End If
Next

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer vFrmLotes_LotesCalidad," & vbCrLf & Err.Description
    MsgBox MSG, , "CargaMuestrasCalidad()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub


