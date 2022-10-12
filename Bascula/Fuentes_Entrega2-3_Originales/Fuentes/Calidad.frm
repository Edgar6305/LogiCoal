VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Calidad 
   Caption         =   "Gestion de Calidad"
   ClientHeight    =   9525
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18135
   LinkTopic       =   "Form1"
   ScaleHeight     =   9525
   ScaleWidth      =   18135
   StartUpPosition =   1  'CenterOwner
   Begin MBSplit.Splitter VSplitter 
      Height          =   8175
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   17355
      _ExtentX        =   30612
      _ExtentY        =   14420
      SplitterPos     =   25
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
         TabIndex        =   2
         Top             =   60
         Width           =   12795
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
            Picture         =   "Calidad.frx":0000
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   5
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
                     Picture         =   "Calidad.frx":0102
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Calidad.frx":041C
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Calidad.frx":0736
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Calidad.frx":0A50
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Calidad.frx":0D6A
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Calidad.frx":1084
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Calidad.frx":139E
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Calidad.frx":16B8
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Calidad.frx":19D2
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Calidad.frx":1CEC
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "Calidad.frx":2006
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
            Picture         =   "Calidad.frx":2320
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.TextBox MemVar_1 
            Height          =   315
            Left            =   2040
            MaxLength       =   15
            TabIndex        =   3
            Top             =   360
            Width           =   1020
         End
         Begin MSComCtl2.DTPicker FechaInicial 
            Height          =   315
            Left            =   2040
            TabIndex        =   6
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
            Format          =   115277825
            CurrentDate     =   44578
         End
         Begin Threed.SSFrame oMarco 
            Height          =   2310
            Left            =   360
            TabIndex        =   7
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
               TabIndex        =   11
               Top             =   450
               Width           =   1080
            End
            Begin VB.VScrollBar oMove 
               Height          =   2100
               Left            =   11580
               TabIndex        =   10
               Top             =   120
               Width           =   255
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
               TabIndex        =   8
               Top             =   450
               Width           =   975
            End
            Begin MSMask.MaskEdBox MemVar_6 
               Height          =   315
               Index           =   0
               Left            =   2160
               TabIndex        =   12
               Top             =   480
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
               TabIndex        =   13
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
               TabIndex        =   14
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
               TabIndex        =   15
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
               Format          =   "#,##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_11 
               Height          =   315
               Index           =   0
               Left            =   6180
               TabIndex        =   16
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
               TabIndex        =   17
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
               TabIndex        =   18
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
               TabIndex        =   19
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
               TabIndex        =   20
               Top             =   480
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
               Format          =   "#,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_16 
               Height          =   315
               Index           =   0
               Left            =   9540
               TabIndex        =   21
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
               TabIndex        =   22
               Top             =   420
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
               TabIndex        =   23
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
            Begin VB.Label Label3 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Transaccion"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   45
               TabIndex        =   39
               Top             =   120
               Width           =   1050
            End
            Begin VB.Label SubLabel3 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Numero"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   0
               Left            =   1140
               TabIndex        =   38
               Top             =   120
               Width           =   1020
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Cantidad"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   2160
               TabIndex        =   37
               Top             =   120
               Width           =   1005
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "No Muestra"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   3180
               TabIndex        =   36
               Top             =   120
               Width           =   1000
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "H"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   4200
               TabIndex        =   35
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "S"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   4860
               TabIndex        =   34
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "BTU"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   4
               Left            =   5520
               TabIndex        =   33
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "CNZ"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   5
               Left            =   6180
               TabIndex        =   32
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "CF"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   6
               Left            =   6840
               TabIndex        =   31
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "V"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   7
               Left            =   7500
               TabIndex        =   30
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dry-S"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   8
               Left            =   8160
               TabIndex        =   29
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dry-BTU"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   9
               Left            =   8820
               TabIndex        =   28
               Top             =   120
               Width           =   705
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dry-CNZ"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   10
               Left            =   9540
               TabIndex        =   27
               Top             =   120
               Width           =   705
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dry-CF"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   11
               Left            =   10275
               TabIndex        =   26
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
               TabIndex        =   25
               Top             =   450
               Width           =   645
            End
            Begin VB.Label Label4 
               Alignment       =   2  'Center
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Dry-V"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   13
               Left            =   10920
               TabIndex        =   24
               Top             =   120
               Width           =   645
            End
         End
         Begin MSComctlLib.ListView ListView1 
            Height          =   1845
            Left            =   420
            TabIndex        =   40
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
            Left            =   2040
            TabIndex        =   41
            Top             =   2280
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
                  Picture         =   "Calidad.frx":2422
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":257C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":26D6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":2830
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":298A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":2AE4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":2C3E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":2D98
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":2EF2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":304C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":31A6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":3300
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":3D12
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":42AC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":4846
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":4DE0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":62EA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":6E34
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":73ED
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":8503
                  Key             =   ""
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":8CD5
                  Key             =   ""
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":9DA7
                  Key             =   ""
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":A8F1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Calidad.frx":BA47
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin MSComCtl2.DTPicker FechaCierre 
            Height          =   315
            Left            =   2040
            TabIndex        =   42
            Top             =   1920
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
            Format          =   115277825
            CurrentDate     =   44578
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Detalle de la Pila"
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
            TabIndex        =   51
            Top             =   2880
            Width           =   1275
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Muestras de Laboratorio"
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
            TabIndex        =   50
            Top             =   5700
            Width           =   1890
         End
         Begin VB.Label Label11 
            Caption         =   "Cantidad "
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
            Index           =   0
            Left            =   420
            TabIndex        =   49
            Top             =   2340
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Numero de Muestra"
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
            TabIndex        =   48
            Top             =   420
            Width           =   1530
         End
         Begin VB.Label LabelPilaFisica 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3180
            TabIndex        =   47
            Top             =   720
            Width           =   4815
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Apertura"
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
            TabIndex        =   46
            Top             =   1620
            Width           =   1170
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha Cierre"
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
            TabIndex        =   45
            Top             =   1980
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label Label11 
            Caption         =   "Toneladas"
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
            Index           =   1
            Left            =   3480
            TabIndex        =   44
            Top             =   2340
            Width           =   855
         End
         Begin Tracer.LabelPlus Cerrado 
            Height          =   555
            Left            =   3540
            TabIndex        =   43
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
            Caption         =   "Calidad.frx":C05D
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
            PictureArr      =   "Calidad.frx":C07D
         End
      End
      Begin MSComctlLib.TreeView oTree 
         Height          =   7995
         Left            =   60
         TabIndex        =   1
         Top             =   60
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
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   780
      Left            =   0
      TabIndex        =   52
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
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   330
         Left            =   165
         TabIndex        =   56
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
         TabIndex        =   55
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
         MICON           =   "Calidad.frx":D513
         PICN            =   "Calidad.frx":D52F
         PICH            =   "Calidad.frx":DAC9
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
         ItemData        =   "Calidad.frx":E063
         Left            =   825
         List            =   "Calidad.frx":E073
         TabIndex        =   54
         Text            =   "TODAS"
         Top             =   405
         Width           =   2355
      End
      Begin KewlButtonz.KewlButtons CierreLotes 
         Height          =   315
         Left            =   4860
         TabIndex        =   53
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
         MICON           =   "Calidad.frx":E0A0
         PICN            =   "Calidad.frx":E0BC
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
      TabIndex        =   57
      Top             =   9210
      Width           =   18135
      _ExtentX        =   31988
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
Attribute VB_Name = "Calidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
