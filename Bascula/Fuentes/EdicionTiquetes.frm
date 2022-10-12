VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form EdicionTiquetes 
   Caption         =   "Edición de Tiquetes"
   ClientHeight    =   8010
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7815
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
   ScaleHeight     =   8010
   ScaleWidth      =   7815
   StartUpPosition =   1  'CenterOwner
   Begin MBSplit.Splitter VSplitter 
      Height          =   6675
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   11774
      SplitterPos     =   0
      Begin MSComctlLib.TreeView oTree 
         Height          =   5895
         Left            =   120
         TabIndex        =   17
         Top             =   60
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   10398
         _Version        =   393217
         Style           =   7
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
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6315
         Left            =   1200
         TabIndex        =   11
         Top             =   180
         Width           =   6195
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   3060
            TabIndex        =   41
            Top             =   4320
            Width           =   2895
         End
         Begin VB.CheckBox Carpado 
            Alignment       =   1  'Right Justify
            Caption         =   "Carpado"
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
            Left            =   480
            TabIndex        =   40
            Top             =   4320
            Width           =   1515
         End
         Begin VB.TextBox MemVar_11 
            BackColor       =   &H00FFFEEA&
            Height          =   1455
            Left            =   1800
            MultiLine       =   -1  'True
            TabIndex        =   34
            Top             =   4740
            Width           =   4095
         End
         Begin MSMask.MaskEdBox MemVar_10 
            Height          =   315
            Left            =   1800
            TabIndex        =   10
            Top             =   3900
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "###,##0"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox MemVar_9 
            Height          =   315
            Left            =   1800
            TabIndex        =   9
            Top             =   3540
            Width           =   975
            _ExtentX        =   1720
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "###,##0"
            PromptChar      =   "_"
         End
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
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   7
            Top             =   2760
            Width           =   975
         End
         Begin VB.TextBox MemVar_8 
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
            TabIndex        =   8
            Top             =   3120
            Width           =   975
         End
         Begin VB.TextBox Memvar_4 
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
            Top             =   1680
            Width           =   1005
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
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   5
            Top             =   2040
            Width           =   555
         End
         Begin VB.TextBox Memvar_6 
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
            TabIndex        =   6
            Top             =   2400
            Width           =   555
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
            TabIndex        =   1
            Top             =   600
            Width           =   1080
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
            Left            =   3060
            Picture         =   "EdicionTiquetes.frx":0000
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   13
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
            Left            =   3060
            Picture         =   "EdicionTiquetes.frx":0102
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   12
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
                     Picture         =   "EdicionTiquetes.frx":0204
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "EdicionTiquetes.frx":051E
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "EdicionTiquetes.frx":0838
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "EdicionTiquetes.frx":0B52
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "EdicionTiquetes.frx":0E6C
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "EdicionTiquetes.frx":1186
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "EdicionTiquetes.frx":14A0
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "EdicionTiquetes.frx":17BA
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "EdicionTiquetes.frx":1AD4
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "EdicionTiquetes.frx":1DEE
                     Key             =   ""
                  EndProperty
                  BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "EdicionTiquetes.frx":2108
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
            MaxLength       =   15
            TabIndex        =   2
            Top             =   960
            Width           =   555
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
            TabIndex        =   3
            Top             =   1320
            Width           =   1035
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   5220
            Top             =   240
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
                  Picture         =   "EdicionTiquetes.frx":2422
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":257C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":26D6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":2830
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":298A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":2AE4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":2C3E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":2D98
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":2EF2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":304C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":31A6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":3300
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":3D12
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":42AC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":4846
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":4DE0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":592A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":69FC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "EdicionTiquetes.frx":7B52
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label FechaVacio 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3060
            TabIndex        =   39
            Top             =   3900
            Width           =   1575
         End
         Begin VB.Label FechaLleno 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3060
            TabIndex        =   38
            Top             =   3540
            Width           =   1575
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
            Left            =   3420
            TabIndex        =   37
            Top             =   600
            Visible         =   0   'False
            Width           =   1275
         End
         Begin VB.Label Label11 
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
            Height          =   195
            Index           =   4
            Left            =   540
            TabIndex        =   35
            Top             =   4740
            Width           =   1095
         End
         Begin VB.Label LabelDestino 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   555
            Left            =   3060
            TabIndex        =   33
            Top             =   1320
            Width           =   2475
         End
         Begin VB.Label LabelConductor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3060
            TabIndex        =   32
            Top             =   3120
            Width           =   2895
         End
         Begin VB.Label LabelTransportador 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3060
            TabIndex        =   31
            Top             =   2400
            Width           =   2475
         End
         Begin VB.Label LabelMaterial 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3060
            TabIndex        =   30
            Top             =   2040
            Width           =   2475
         End
         Begin VB.Label LabelOrigen 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3060
            TabIndex        =   29
            Top             =   960
            Width           =   2475
         End
         Begin VB.Label Label11 
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
            Height          =   195
            Index           =   3
            Left            =   500
            TabIndex        =   28
            Top             =   3900
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Height          =   210
            Index           =   3
            Left            =   495
            TabIndex        =   27
            Top             =   3540
            Width           =   795
         End
         Begin VB.Label Label11 
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
            Height          =   195
            Index           =   2
            Left            =   500
            TabIndex        =   26
            Top             =   2820
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
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
            Height          =   210
            Index           =   2
            Left            =   495
            TabIndex        =   25
            Top             =   3180
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "ID Transportador"
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
            Left            =   495
            TabIndex        =   24
            Top             =   2460
            Width           =   1215
         End
         Begin VB.Label Label11 
            Caption         =   "ID Material"
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
            Index           =   1
            Left            =   500
            TabIndex        =   23
            Top             =   2100
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Remisión"
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
            Left            =   495
            TabIndex        =   22
            Top             =   1740
            Width           =   645
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "No Tiquete"
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
            Left            =   495
            TabIndex        =   16
            Top             =   660
            Width           =   765
         End
         Begin VB.Label Label11 
            Caption         =   "Transacción"
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
            Index           =   0
            Left            =   500
            TabIndex        =   15
            Top             =   1020
            Width           =   1035
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Numero"
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
            Left            =   495
            TabIndex        =   14
            Top             =   1380
            Width           =   555
         End
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   750
      Left            =   0
      TabIndex        =   18
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
      Child3          =   "AnularTiquete"
      MinHeight3      =   315
      Width3          =   1740
      NewRow3         =   0   'False
      MinHeight4      =   315
      Width4          =   1695
      NewRow4         =   0   'False
      Begin KewlButtonz.KewlButtons AnularTiquete 
         Height          =   315
         Left            =   3405
         TabIndex        =   36
         Top             =   390
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Anular Tiquete"
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
         BCOL            =   14215660
         BCOLO           =   14215660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "EdicionTiquetes.frx":8168
         PICN            =   "EdicionTiquetes.frx":8184
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   -1  'True
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
         ItemData        =   "EdicionTiquetes.frx":871E
         Left            =   825
         List            =   "EdicionTiquetes.frx":872E
         TabIndex        =   20
         Text            =   "TODAS"
         Top             =   390
         Width           =   2355
      End
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   330
         Left            =   165
         TabIndex        =   19
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
               Object.Visible         =   0   'False
               Key             =   "Borrar"
               Object.ToolTipText     =   "Borrar Registro Actual"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
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
               Object.Visible         =   0   'False
               Key             =   "Foto"
               Object.ToolTipText     =   "Insetar Imagen"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
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
      TabIndex        =   21
      Top             =   7695
      Width           =   7815
      _ExtentX        =   13785
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
Attribute VB_Name = "EdicionTiquetes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Acopios As New ADODB.Recordset
Dim okUnload As Boolean
Dim IsNewRecord As Boolean
Dim OkOpen As Boolean
Dim oT As New ADODB.Recordset
Dim xR As New ADODB.Recordset

Private Sub Anulado_Click()
Dim xText As String

xText = F_Anulado(1, MemVar_1)
MsgBox xText, vbDefaultButton1, "LogyCoal"

End Sub

Private Sub AnularTiquete_Click()
If MsgBox("Esta seguro de Anular El Tiquete ", vbYesNo, "Borrado de Registro") = vbYes Then
    Set Anulaciones.dControl = MemVar_1
    Anulaciones.MemVar_1 = 1
    Anulaciones.MemVar_2 = MemVar_1
    Anulaciones.MemVar_3 = Now
    Anulaciones.Show 1
    If MemVar_1.Tag Then
       Conn.Execute ("Update Bascula Set Estado='AN' Where IdTiquete=" & MemVar_1)
    End If
End If
End Sub

Private Sub Form_Activate()
If Not OkOpen Then
     Call LoadData
     OkOpen = True
     MemVar_1.SetFocus
End If
Me.SetFocus
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
Dim xCt As New ADODB.Recordset
Dim xR As New ADODB.Recordset

On Error Resume Next

xSql = "Select Top 1 * From Bascula Order By IdTiquete DESC"
oT.Open xSql, Conn, 2, 3, 1

Set VSplitter.LeftOrTopCtl = oTree
Set VSplitter.RightOrBottomCtl = Frame1

Set xR = Conn.Execute("Select * From TiposCarbon Order by Descripcion DESC")
Do While Not xR.EOF
    Combo1.AddItem Format(xR!IdTipoCarbon, "00") & " " & xR!Descripcion
    xR.MoveNext
Loop
xR.Close

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error en Form_Load()," & vbCrLf & Err.Description
    MsgBox MSG, , "LoadData"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub

Private Sub LoadData()
Dim xDes As String

On Error GoTo Recover

 If oT.EOF And oT.BOF Then
         okNew.Visible = True
         okFind.Visible = False
         Exit Sub
 Else
         okNew.Visible = False
         okFind.Visible = True
 End If

MemVar_1 = oT!IdTiquete
MemVar_2 = oT!TransaccionOrigen
LabelOrigen = Conn.Execute("Select Descripcion From TipoTransaccion Where IdTipo='" & MemVar_2 & "'").Fields(0)
MemVar_3 = oT!NumeroTransaccion
LabelDestino = ""
If MemVar_2 <> "RO" Then
    xDes = Conn.Execute("Select  Descripcion From vDestinos Where Tipo='" & MemVar_2 & "' AND IdVentas=" & MemVar_3).Fields(0)
    Select Case MemVar_2
    Case "DS"
        LabelDestino = "Venta No" + MemVar_3 + vbCrLf + "Destino " + xDes
    Case "TR"
        LabelDestino = "Traslado No" + MemVar_3 + vbCrLf + "Destino " + xDes
    Case "LT"
        LabelDestino = "Lote No" + MemVar_3 + vbCrLf + "Manto " + xDes
    End Select
End If
MemVar_4 = oT!Documentoasociado
MemVar_5 = oT!IdMaterial
LabelMaterial = Conn.Execute("Select Descripcion From Materiales Where IdMaterial='" & MemVar_5 & "'").Fields(0)
MemVar_6 = oT!IdTransportador
LabelTransportador = Conn.Execute("Select Descripcion From Transportador Where IdTransportador='" & MemVar_6 & "'").Fields(0)
MemVar_7 = oT!Placas
MemVar_8 = oT!Conductor
LabelConductor = Conn.Execute("Select Nombre From Conductores Where Cedula='" & MemVar_8 & "'").Fields(0)
MemVar_9 = oT!PesoLleno
MemVar_10 = oT!PesoVacio
FechaLleno = Format(oT!FechaLleno, "dd/MM/yyyy hh:mm")
FechaVacio = Format(oT!FechaVacio, "dd/MM/yyyy hh:mm")
MemVar_11 = IIf(IsNull(oT!Observaciones), "", oT!Observaciones)
Carpado = IIf(oT!Carpado, 1, 0)
Combo1.ListIndex = CargaTipo(Combo1, oT!IdTipoCarbon)

Anulado.Visible = (oT!Estado = "AN")
Tbar.buttons("Grabar").Enabled = Not (oT!Estado = "AN")

AnularTiquete.Enabled = (Not (oT!Estado = "AN") And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','AnularTiquete'").Fields(0)) '==> Fields(0) Nos dice si tiene acceso al Objeto

oBar.Panels("Usuario").text = "USUARIO: " & oT!Usuario
oBar.Panels("Ot").text = "ESTADO: " & oT!Estado

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer Transaccion," & vbCrLf & Err.Description
    MsgBox MSG, , "LoadData"
    Err.Clear   ' Borra campos del objeto Err
    Exit Sub
End If
    
End Sub

Private Function CargaTipo(Combo1 As Object, Tipo As Integer) As Integer
Dim i As Integer, j As Integer
j = -1
For i = 0 To Combo1.ListCount - 1
    If Val(Mid(Combo1.List(i), 1, 2)) = Tipo Then
        j = i
    End If
Next i
CargaTipo = j
End Function

Private Sub SaveData()
Dim ok As Boolean
Dim xPesoV As Single, xPesoL As Single
Dim xFecV As Date, xFecL As Date

On Error GoTo Recover

    ok = False
    If MemVar_1.text = "" Then ok = True
    If MemVar_2.text = "" Then ok = True
    If ok Then Exit Sub
        
    xPesoV = oT!PesoVacio
    xPesoL = oT!PesoLleno
        
    oT!IdTiquete = MemVar_1
    oT!TransaccionOrigen = MemVar_2
    oT!NumeroTransaccion = MemVar_3
    oT!Documentoasociado = MemVar_4
    oT!IdMaterial = MemVar_5
    oT!IdTransportador = MemVar_6
    oT!Placas = MemVar_7
    oT!Conductor = MemVar_8
    oT!PesoLleno = MemVar_9
    oT!PesoVacio = MemVar_10
    oT!Observaciones = MemVar_11
    oT!Carpado = Carpado
    oT!IdTipoCarbon = Val(Mid(Combo1, 1, 2))
    oT.Update
    
    If (xPesoV = 0) Then
        If MsgBox("Desea Cerrar el Tiquete", vbYesNo, "Cerrado de Tiquete-Estado='AC'") = vbYes Then
           Conn.Execute ("Set DateFormat DMY Update Bascula Set Estado='AC', FechaVacio='" & Format(Now, "dd/MM/yyyy hh:mm:ss") & "' Where IdTiquete=" & MemVar_1)
        End If
    End If
     
     
    If Not MySeek(oT, Conn, "Bascula", "IdTiquete=" & MemVar_1) Then
            Call LoadData
            okFind.Visible = True
            okNew.Visible = False
    End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Salvar los Datos," & vbCrLf & Err.Description
    MsgBox MSG, , "Savedata()"
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

If MemVar_1 <> "" Then
    If Not MySeek(oT, Conn, "Bascula", "idTiquete=" & MemVar_1) Then
        Call LoadData
    Else
        MsgBox "No se Localizó el Tiquete", vbInformation
        MemVar_1.SetFocus
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Registro" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_1_LostFocus()"
    Err.Clear
    Resume Next
End If
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

Private Sub MemVar_2_GotFocus()
        Call Mark(MemVar_2)
End Sub

Private Sub MemVar_2_LostFocus()
Dim xSql As String

On Error GoTo Recover
If MemVar_2 <> "" Then
    MemVar_2 = UCase(MemVar_2)
    Set xR = Conn.Execute("Select * From TipoTransaccion Where IdTipo='" & MemVar_2 & "'")
    
     If Not xR.EOF Then
            LabelOrigen = Conn.Execute("Select Descripcion From TipoTransaccion Where IdTipo='" & MemVar_2 & "'").Fields(0)
     Else
            MsgBox "Transacción NO Registrado, Verifique"
            MemVar_2.SetFocus
    End If
    xR.Close
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Tipos de Transacciones" & vbCrLf & Err.Description
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
                Call omenu("Browse")
        Case vbKeyUp
                MemVar_2.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_4.SetFocus
End Select

End Sub

Private Sub MemVar_3_LostFocus()
If MemVar_3 <> "" Then
    If MemVar_2 <> "RO" Then
        If Conn.Execute("Select  Descripcion From vDestinos Where Tipo='" & MemVar_2 & "' AND IdVentas=" & MemVar_3).EOF Then
            MsgBox "El Numero de Transaccion NO existe, Verifique", vbInformation
            MemVar_3.SetFocus
        Else
            xDes = Conn.Execute("Select  Descripcion From vDestinos Where Tipo='" & MemVar_2 & "' AND IdVentas=" & MemVar_3).Fields(0)
            Select Case MemVar_2
            Case "DS"
                LabelDestino = "Venta No" + MemVar_3 + vbCrLf + "Destino " + xDes
            Case "TR"
                LabelDestino = "Traslado No" + MemVar_3 + vbCrLf + "Destino " + xDes
            Case "LT"
                LabelDestino = "Lote No" + MemVar_3 + vbCrLf + "Manto " + xDes
           Case "RO"
                LabelDestino = "Recepcion OTROS Materiales"
            End Select
        End If
    End If
End If
Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Tipos de Transacciones" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_2_LostFocus()"
    Err.Clear
    Exit Sub
End If
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

Private Sub MemVar_5_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
                Call omenu("Browse")
        Case vbKeyUp
                MemVar_4.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_6.SetFocus
End Select

End Sub

Private Sub MemVar_5_GotFocus()
        Call Mark(MemVar_5)
End Sub

Private Sub MemVar_5_LostFocus()
Dim xSql As String

On Error GoTo Recover

If MemVar_5 <> "" Then
    Set xR = Conn.Execute("Select * From Materiales Where IdMaterial=" & MemVar_5)
     If Not xR.EOF Then
            Call LoadData
     Else
            MsgBox "Material NO Registrado"
            MemVar_5.SetFocus
    End If
    xR.Close
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Materiales" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_5_LostFocus()"
    Err.Clear
    Resume Next
End If

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

Private Sub MemVar_6_GotFocus()
        Call Mark(MemVar_6)
End Sub

Private Sub MemVar_6_LostFocus()
Dim xSql As String

On Error GoTo Recover

If MemVar_6 <> "" Then
    Set xR = Conn.Execute("Select * From Transportador Where IdTransportador=" & MemVar_6)
     
     If Not xR.EOF Then
            ' OK
     Else
            MsgBox "Transportador NO Registrado"
            MemVar_6.SetFocus
    End If
    xR.Close
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Transportador" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_6_LostFocus()"
    Err.Clear
    Resume Next
End If

End Sub

Private Sub MemVar_7_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
                Call omenu("Browse")
        Case vbKeyUp
                MemVar_6.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_8.SetFocus
End Select

End Sub

Private Sub MemVar_7_GotFocus()
        Call Mark(MemVar_7)
End Sub

Private Sub MemVar_7_LostFocus()
Dim xSql As String

On Error GoTo Recover

If MemVar_7 <> "" Then
    Set xR = Conn.Execute("Select * From Placas Where Placas='" & MemVar_7 & "'")
    
     If Not xR.EOF Then
            'OK
     Else
            MsgBox "Placa NO Registrada"
            MemVar_6.SetFocus
    End If
xR.Close
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Placas" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_7_LostFocus()"
    Err.Clear
    Resume Next
End If

End Sub


Private Sub MemVar_8_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
                Call omenu("Browse")
        Case vbKeyUp
                MemVar_7.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_9.SetFocus
End Select

End Sub

Private Sub MemVar_8_GotFocus()
        Call Mark(MemVar_8)
End Sub

Private Sub MemVar_8_LostFocus()
Dim xSql As String

On Error GoTo Recover

If MemVar_8 <> "" Then
    Set xR = Conn.Execute("Select * From Conductores Where Cedula='" & MemVar_8 & "'")
    
     If Not xR.EOF Then
            'OK
     Else
            MsgBox "Conductor NO Registrado"
            MemVar_8.SetFocus
    End If
    xR.Close
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga Conductores" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_8_LostFocus()"
    Err.Clear
    Resume Next
End If

End Sub

Private Sub MemVar_9_GotFocus()
        Call Mark(MemVar_9)
End Sub

Private Sub MemVar_9_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
                MemVar_8.SetFocus
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
                MemVar_1.SetFocus
End Select
End Sub

Private Sub MemVar_10_LostFocus()
    If MemVar_10 >= MemVar_9 Then
       MsgBox "Peso VACIO mayor que Peso LLeno, Verifique "
       MemVar_10.SetFocus
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
                Case "Salida"
                    Unload Me
                Case "Browse"
                        Select Case ActiveControl.Name
                        Case "MemVar_1"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "Bascula"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        Case "MemVar_2"
                                BrowseCatalogo.x_Tipo = 2
                                BrowseCatalogo.xtabla = "SELECT DISTINCT vDestinos.Tipo, TipoTransaccion.Descripcion FROM vDestinos INNER JOIN TipoTransaccion ON vDestinos.Tipo = TipoTransaccion.IdTipo"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        Case "MemVar_3"
                                BrowseCatalogo.x_Tipo = 2
                                BrowseCatalogo.xtabla = "Select idVentas, Descripcion From vDestinos Where Tipo='" & MemVar_2 & "'"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                                
                        Case "MemVar_5"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "Materiales"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        Case "Memvar_6"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "Transportador"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        Case "MemVar_7"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "Placas"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        Case "MemVar_8"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "Conductores"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show 1
                        
                        End Select
                Case "Imprime"

                Case "Borrar"
                
                Case "Top"
                    oT.Close
                    oT.Open "Select Top 1 * From Bascula Order By IdTiquete"
                    Call LoadData
                Case "Bottom"
                    oT.Close
                    oT.Open "Select Top 1 * From Bascula Order By IdTiquete DESC"
                    Call LoadData
                Case "Proximo"
                    oT.Close
                    oT.Open "Select Top 1 * From Bascula Where IdTiquete>'" & MemVar_1 & "' Order By  IdTiquete"
                    Call LoadData
                Case "Previo"
                    oT.Close
                    oT.Open "Select Top 1 * From Bascula Where IdTiquete<'" & MemVar_1 & "' Order By IdTiquete DESC"
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

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
                dControl.SelStart = 0
                dControl.SelLength = Len(dControl.text)
End If
End Sub
