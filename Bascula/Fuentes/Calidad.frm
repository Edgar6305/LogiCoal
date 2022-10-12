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
   ClientHeight    =   8715
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9930
   LinkTopic       =   "Form1"
   ScaleHeight     =   8715
   ScaleWidth      =   9930
   StartUpPosition =   1  'CenterOwner
   Begin MBSplit.Splitter VSplitter 
      Height          =   7155
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   9615
      _ExtentX        =   16960
      _ExtentY        =   12621
      SplitterPos     =   36
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6735
         Left            =   3780
         TabIndex        =   2
         Top             =   300
         Width           =   4755
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
            Top             =   780
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
            Format          =   115867649
            CurrentDate     =   44578
         End
         Begin Threed.SSFrame oMarco 
            Height          =   4230
            Left            =   420
            TabIndex        =   7
            Top             =   2400
            Width           =   3015
            _Version        =   65536
            _ExtentX        =   5318
            _ExtentY        =   7461
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
            Begin MSMask.MaskEdBox MemVar_8 
               Height          =   315
               Left            =   1680
               TabIndex        =   8
               Top             =   120
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
               Format          =   "#0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_9 
               Height          =   315
               Left            =   1680
               TabIndex        =   9
               Top             =   480
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
               Format          =   "#0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_10 
               Height          =   315
               Left            =   1680
               TabIndex        =   10
               Top             =   840
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
               Format          =   "##,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_11 
               Height          =   315
               Left            =   1680
               TabIndex        =   11
               Top             =   1200
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
               Format          =   "#0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_12 
               Height          =   315
               Left            =   1680
               TabIndex        =   12
               Top             =   1560
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
               Format          =   "#0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_13 
               Height          =   315
               Left            =   1680
               TabIndex        =   13
               Top             =   1920
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
               Format          =   "#0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_14 
               Height          =   315
               Left            =   1680
               TabIndex        =   14
               Top             =   2280
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
               Format          =   "#0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_15 
               Height          =   315
               Left            =   1680
               TabIndex        =   15
               Top             =   2640
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
               Format          =   "##,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_16 
               Height          =   315
               Left            =   1680
               TabIndex        =   16
               Top             =   3000
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
               Format          =   "#0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_17 
               Height          =   315
               Left            =   1680
               TabIndex        =   17
               Top             =   3360
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
               Format          =   "#0.00"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox MemVar_18 
               Height          =   315
               Left            =   1680
               TabIndex        =   18
               Top             =   3720
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
               Format          =   "#0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
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
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   2
               Left            =   60
               TabIndex        =   29
               Top             =   120
               Width           =   645
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
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
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   3
               Left            =   60
               TabIndex        =   28
               Top             =   480
               Width           =   645
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
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
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   4
               Left            =   60
               TabIndex        =   27
               Top             =   840
               Width           =   645
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
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
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   5
               Left            =   60
               TabIndex        =   26
               Top             =   1200
               Width           =   645
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
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
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   6
               Left            =   60
               TabIndex        =   25
               Top             =   1560
               Width           =   645
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
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
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   7
               Left            =   60
               TabIndex        =   24
               Top             =   1920
               Width           =   645
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
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
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   8
               Left            =   60
               TabIndex        =   23
               Top             =   2280
               Width           =   645
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
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
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   9
               Left            =   60
               TabIndex        =   22
               Top             =   2640
               Width           =   705
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
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
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   10
               Left            =   60
               TabIndex        =   21
               Top             =   3000
               Width           =   705
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
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
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   11
               Left            =   60
               TabIndex        =   20
               Top             =   3360
               Width           =   645
            End
            Begin VB.Label Label4 
               Appearance      =   0  'Flat
               BackColor       =   &H80000004&
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
               ForeColor       =   &H80000008&
               Height          =   300
               Index           =   13
               Left            =   60
               TabIndex        =   19
               Top             =   3720
               Width           =   645
            End
         End
         Begin MSMask.MaskEdBox MemVar_3 
            Height          =   315
            Left            =   2040
            TabIndex        =   30
            Top             =   1500
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            ForeColor       =   16711680
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
            Format          =   "###,##0.00"
            PromptChar      =   "_"
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   3660
            Top             =   780
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
            TabIndex        =   31
            Top             =   1140
            Visible         =   0   'False
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
            Format          =   115867649
            CurrentDate     =   44578
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Detalle de la Muestra"
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
            TabIndex        =   38
            Top             =   2040
            Width           =   1635
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
            TabIndex        =   37
            Top             =   1560
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
            TabIndex        =   36
            Top             =   420
            Width           =   1530
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
            TabIndex        =   35
            Top             =   840
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
            TabIndex        =   34
            Top             =   1200
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
            Left            =   3420
            TabIndex        =   33
            Top             =   1500
            Width           =   855
         End
         Begin Tracer.LabelPlus Cerrado 
            Height          =   555
            Left            =   3540
            TabIndex        =   32
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
         Height          =   6915
         Left            =   60
         TabIndex        =   1
         Top             =   60
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   12197
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
      TabIndex        =   39
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
      Width4          =   2100
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
         TabIndex        =   43
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
         TabIndex        =   42
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
         TabIndex        =   41
         Text            =   "TODAS"
         Top             =   405
         Width           =   2355
      End
      Begin KewlButtonz.KewlButtons CierreLotes 
         Height          =   315
         Left            =   4935
         TabIndex        =   40
         Top             =   405
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Cierre de Muestra"
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
      TabIndex        =   44
      Top             =   8400
      Width           =   9930
      _ExtentX        =   17515
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
Attribute VB_Name = "Calidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim okUnload As Boolean
Dim IsNewRecord As Boolean
Dim OkOpen As Boolean
Dim oT As New ADODB.Recordset

Private Sub Actualizar_Click()
If MsgBox("Esta Seguro de Actualizar La Nota de Inventarios ", vbYesNo, "Notas de Inventarios") = vbYes Then
    xSql = "UPDATE Ajustes Set Estado = 'AC' Where IdNotas=" & MemVar_1
    Conn.Execute (xSql)
    Call MemVar_1_LostFocus
End If
End Sub

Private Sub Form_Activate()
If Not OkOpen Then
'     Call LoadData
     OkOpen = True
'    MemVar_1.SetFocus
End If
Me.SetFocus
Call MuestraArbol
oTree.SetFocus
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

xSql = "Select Top 1 * From Calidad " & xFilter
oT.Open xSql, Conn, 2, 3, 1

Set VSplitter.LeftOrTopCtl = oTree
Set VSplitter.RightOrBottomCtl = Frame1

End Sub

Private Sub MemVar_3_GotFocus()
        Call Mark(MemVar_3)
End Sub

Private Sub MemVar_3_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
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

Private Sub MemVar_8_GotFocus()
    Call Mark(MemVar_8)
End Sub

Private Sub MemVar_8_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
                MemVar_1.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_9.SetFocus
End Select
End Sub

Private Sub MemVar_9_GotFocus()
    Call Mark(MemVar_9)
End Sub

Private Sub MemVar_9_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
                MemVar_9.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_10.SetFocus
End Select
End Sub

Private Sub MemVar_10_GotFocus()
    Call Mark(MemVar_10)
End Sub

Private Sub MemVar_10_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
                MemVar_9.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_11.SetFocus
End Select
End Sub

Private Sub MemVar_11_GotFocus()
    Call Mark(MemVar_11)
End Sub

Private Sub MemVar_11_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
                MemVar_10.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_12.SetFocus
End Select
End Sub

Private Sub MemVar_12_GotFocus()
    Call Mark(MemVar_12)
End Sub

Private Sub MemVar_12_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
                MemVar_11.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_13.SetFocus
End Select
End Sub

Private Sub MemVar_13_GotFocus()
    Call Mark(MemVar_13)
End Sub

Private Sub MemVar_13_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
                MemVar_12.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_14.SetFocus
End Select
End Sub

Private Sub MemVar_14_GotFocus()
    Call Mark(MemVar_14)
End Sub

Private Sub MemVar_14_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
                MemVar_13.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_15.SetFocus
End Select
End Sub

Private Sub MemVar_15_GotFocus()
    Call Mark(MemVar_15)
End Sub

Private Sub MemVar_15_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
                MemVar_14.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_16.SetFocus
End Select
End Sub

Private Sub MemVar_16_GotFocus()
    Call Mark(MemVar_16)
End Sub

Private Sub MemVar_16_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
                MemVar_15.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_17.SetFocus
End Select
End Sub

Private Sub MemVar_17_GotFocus()
    Call Mark(MemVar_17)
End Sub

Private Sub MemVar_17_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
                MemVar_16.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_18.SetFocus
End Select
End Sub

Private Sub MemVar_18_GotFocus()
    Call Mark(MemVar_18)
End Sub

Private Sub MemVar_18_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
                MemVar_17.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_1.SetFocus
                Call SaveData
End Select
End Sub



Private Sub oNuevo_Click()
If oNuevo.Caption = "Nuevo" Then
    oNuevo.Caption = "Cancelar"
    IsNewRecord = True
    MemVar_8.text = ""
    MemVar_9.text = ""
    MemVar_10.text = ""
    MemVar_11.text = ""
    MemVar_12.text = ""
    MemVar_13.text = ""
    MemVar_14.text = ""
    MemVar_15.text = ""
    MemVar_16.text = ""
    MemVar_17.text = ""
    MemVar_18.text = ""
    
    oBar.Panels("Usuario").text = "USUARIO: " & Susuario
    oBar.Panels("Ot").text = "FECHA CREACION: " & Now
    oBar.Panels("Estado").text = "ESTADO: " & "IN"
    
    Tbar.buttons("Grabar").Enabled = True
    
    MemVar_8.SetFocus
Else
    oNuevo.Caption = "Nuevo"
    IsNewRecord = False
    MemVar_1.SetFocus
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
                    BrowseCatalogo.xtabla = "NotasDedCre"
                    Set BrowseCatalogo.dControl = ActiveControl
                    BrowseCatalogo.Show 1
            Case "MemVar_2"
                    BrowseAcopioPilas.x_Tipo = 1
                    BrowseAcopioPilas.xtabla = "vPilasAcopiosGeneral"
                    Set BrowseAcopioPilas.dControl = ActiveControl
                    BrowseAcopioPilas.Show 1
            End Select
            
    Case "Imprime"
    
    Case "Borrar"
    
    Case "Top"
        oT.Close
        oT.Open "Select Top 1 * From Ajustes Order By IdNotas"
        Call LoadData
    Case "Bottom"
        oT.Close
        oT.Open "Select Top 1 * From Ajustes Order By IdNotas DESC"
        Call LoadData
    Case "Proximo"
        oT.Close
        oT.Open "Select Top 1 * From Ajustes Where IdNotas>'" & MemVar_1 & "' Order By  IdNotas"
        Call LoadData
    Case "Previo"
        oT.Close
        oT.Open "Select Top 1 * From Ajustes Where IdNotas<'" & MemVar_1 & "' Order By IdNotas DESC"
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
'On Error GoTo Recover

 If oT.EOF And oT.BOF Then Exit Sub
 
 FechaInicial = oT!Fecha
 MemVar_3 = oT!Cantidad
 MemVar_8 = oT!H
 MemVar_9 = oT!S
 MemVar_10 = oT!BTU
 MemVar_11 = oT!CNZ
 MemVar_12 = oT!CF
 MemVar_13 = oT!v
 MemVar_14 = oT!Dry_S
 MemVar_15 = oT!Dry_BTU
 MemVar_16 = oT!Dry_CNZ
 MemVar_17 = oT!Dry_CF
 MemVar_18 = oT!Dry_V
    
 oBar.Panels("Usuario").text = "USUARIO: " & oT!UsuarioEntrega
 oBar.Panels("Ot").text = "FECHA CREACION: " & oT!FechaEntrega
 oBar.Panels("Estado").text = "ESTADO: " & oT!Estado

' Actualizar.Enabled = (oT!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','EstadoNotas'").Fields(0)) '==> Fields(0) Nos dice si tiene acceso al Objeto

 Tbar.buttons("Grabar").Enabled = (oT!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','" & sProgSISMA & "'").Fields(1)) '==> Fields(1) Nos dice puede hacer cambios al Objeto
 Tbar.buttons("Borrar").Enabled = (oT!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','" & sProgSISMA & "'").Fields(1)) '==> Fields(1) Nos dice puede hacer cambios al Objeto

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer Transaccion," & vbCrLf & Err.Description
    MsgBox MSG, , "LoadData"
    Err.Clear   ' Borra campos del objeto Err
    Exit Sub
End If
    
End Sub

Private Sub SaveData()
Dim xSql As String
Dim Res As String
Dim xR As New ADODB.Recordset

'On Error GoTo Recover

 xSql = "UPDATE Calidad "
 xSql = xSql + " Set H=" & MemVar_8 & ", "
 xSql = xSql + " S=" & MemVar_9 & ", "
 xSql = xSql + " BTU=" & MemVar_10 & ", "
 xSql = xSql + " CNZ=" & MemVar_11 & ", "
 xSql = xSql + " CF=" & MemVar_12 & ", "
 xSql = xSql + " v=" & MemVar_13 & ", "
 xSql = xSql + " Dry_S=" & MemVar_14 & ", "
 xSql = xSql + " Dry_BTU=" & MemVar_15 & ","
 xSql = xSql + " Dry_CNZ=" & MemVar_16 & ", "
 xSql = xSql + " Dry_CF=" & MemVar_17 & ","
 xSql = xSql + " Dry_V=" & MemVar_18 & ","
 xSql = xSql + " UsuarioEntrega='" & Susuario & "',"
 xSql = xSql + " FechaEntrega='" & Format(Now, "dd/MM/yyyy hh:mm") & "',"
 xSql = xSql + " Estado='AC'"

 xSql = xSql + " WHERE IdMuestra=" & MemVar_1
 Conn.Execute (xSql)
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
Set Nodx = oTree.Nodes.Add(, , "0" & "Muestras", "Muestras", 15)

xSql = "SELECT TransaccionOrigen, "
xSql = xSql + "CASE WHEN TransaccionOrigen ='DS' THEN 'VENTAS'"
xSql = xSql + "     WHEN TransaccionOrigen ='TR' THEN 'TRASLADOS'"
xSql = xSql + "     WHEN TransaccionOrigen ='LT' THEN 'PRODUCCION' ELSE '' END AS Descripcion"
xSql = xSql + " FROM Calidad"
xSql = xSql + " GROUP BY TransaccionOrigen"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("0" & "Muestras", tvwChild, "A" & Em!TransaccionOrigen, Em!Descripcion, 19)
    Em.MoveNext
Wend

xSql = "SELECT * FROM  Calidad Where Estado='IN'"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("A" & Em!TransaccionOrigen, tvwChild, "B" & Format(Em!IdMuestra, "######"), Format(Em!IdMuestra, "000000") + " " + Format(Em!Fecha, "dd/MM/yyyy hh:mm"), 17)
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
        Case vbKeyUp
        Case vbKeyDown, vbKeyReturn
                MemVar_8.SetFocus
End Select

End Sub

Private Sub MemVar_1_GotFocus()
        Call Mark(MemVar_1)
End Sub

Private Sub MemVar_1_LostFocus()
Dim xSql As String

On Error GoTo Recover

If Not MySeek(oT, Conn, "Calidad", "IdMuestra='" & MemVar_1 & "'") Then
    Call LoadData
Else
    MsgBox "Muestra de Calidad NO Registrado, Verifique"
    MemVar_1.SetFocus
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga La Muestra" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_1_LostFocus()"
    Err.Clear
    Exit Sub
End If

End Sub

Private Sub oTree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer

Select Case Mid(Node.Key, 1, 1)
Case "B"
    MemVar_1 = Mid(Node.Key, 2, 10)
    Call MemVar_1_LostFocus
    MemVar_8.SetFocus
End Select

End Sub

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
    dControl.SelStart = 0
    dControl.SelLength = Len(dControl.text)
End If
End Sub


