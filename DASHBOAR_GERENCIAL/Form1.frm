VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Despachos Mes"
   ClientHeight    =   13500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   27480
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13500
   ScaleWidth      =   27480
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   7560
      Top             =   180
   End
   Begin DashBoard.ucChartBar Char_DespachosHora 
      Height          =   2535
      Left            =   9780
      TabIndex        =   45
      Top             =   10680
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   4471
      Title           =   "Despachos por Dia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Border          =   -1  'True
      ChartStyle      =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelsPositions =   3
      BorderRound     =   5
      BorderColor     =   -2147483638
   End
   Begin DashBoard.ucChartBar Char_DespachosDia 
      Height          =   2535
      Left            =   360
      TabIndex        =   44
      Top             =   10680
      Width           =   9195
      _ExtentX        =   16219
      _ExtentY        =   4471
      Title           =   "Despachos por Dia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Border          =   -1  'True
      ChartStyle      =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelsPositions =   3
      BorderRound     =   5
      BorderColor     =   -2147483638
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   315
      Index           =   1
      Left            =   2160
      TabIndex        =   43
      Top             =   1260
      Width           =   3615
   End
   Begin VB.Label Label14 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Despachos Turno 1"
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
      Left            =   18960
      TabIndex        =   42
      Top             =   1320
      Width           =   1515
   End
   Begin DashBoard.ucProgressCircular DespachosTurno1 
      Height          =   975
      Left            =   19140
      TabIndex        =   41
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Ton."
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   9
      PB_Color1       =   10527742
      PB_Border       =   -1  'True
      Max             =   2000
      Value           =   1000
      DisplayInPercent=   0   'False
      PF_ForeColor    =   133614
      AnimationInterval=   100
   End
   Begin DashBoard.LabelPlus LabelPlus11 
      Height          =   1755
      Left            =   18900
      TabIndex        =   40
      Top             =   180
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   3096
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   2
      Caption         =   "Form1.frx":0D02
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CallOutPosicion =   3
      HotLine         =   -1  'True
      HotLineColor    =   255
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
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Despachos Turno 2"
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
      Left            =   20820
      TabIndex        =   39
      Top             =   1320
      Width           =   1515
   End
   Begin DashBoard.ucProgressCircular DespachosTurno2 
      Height          =   975
      Left            =   21000
      TabIndex        =   38
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Ton."
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   9
      PB_Color1       =   10527742
      PB_Border       =   -1  'True
      Max             =   2000
      Value           =   1000
      DisplayInPercent=   0   'False
      PF_ForeColor    =   133614
      AnimationInterval=   100
   End
   Begin DashBoard.LabelPlus LabelPlus10 
      Height          =   1755
      Left            =   20760
      TabIndex        =   37
      Top             =   180
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   3096
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   2
      Caption         =   "Form1.frx":0D36
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CallOutPosicion =   3
      HotLine         =   -1  'True
      HotLineColor    =   255
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
   Begin DashBoard.ucProgressCircular ucProgressCircular2 
      Height          =   975
      Left            =   21060
      TabIndex        =   36
      Top             =   300
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Ton."
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   9
      PB_Color1       =   10527742
      PB_Border       =   -1  'True
      Max             =   4000
      Value           =   1000
      DisplayInPercent=   0   'False
      PF_ForeColor    =   133614
      AnimationInterval=   100
   End
   Begin VB.Label Label12 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Despachos Dia"
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
      Left            =   21060
      TabIndex        =   35
      Top             =   1380
      Width           =   1215
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CarbónTurno 2"
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
      Left            =   10140
      TabIndex        =   34
      Top             =   1260
      Width           =   1215
   End
   Begin DashBoard.ucProgressCircular CarbonTurno2 
      Height          =   975
      Left            =   10140
      TabIndex        =   33
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Ton."
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   9
      PB_Color1       =   16501147
      PB_Border       =   -1  'True
      Max             =   2000
      Value           =   500
      DisplayInPercent=   0   'False
      PF_ForeColor    =   13526537
      AnimationInterval=   100
   End
   Begin DashBoard.LabelPlus LabelPlus3 
      Height          =   1755
      Left            =   10080
      TabIndex        =   32
      Top             =   120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   3096
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   2
      Caption         =   "Form1.frx":0D6A
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotLine         =   -1  'True
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
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "CarbónTurno 1"
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
      Left            =   8640
      TabIndex        =   31
      Top             =   1260
      Width           =   1215
   End
   Begin DashBoard.ucProgressCircular CarbonTurno1 
      Height          =   975
      Left            =   8640
      TabIndex        =   30
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Ton."
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   9
      PB_Color1       =   16501147
      PB_Border       =   -1  'True
      PB_BorderColor  =   -2147483646
      Max             =   2000
      Value           =   500
      DisplayInPercent=   0   'False
      PF_ForeColor    =   13526537
      AnimationInterval=   100
   End
   Begin DashBoard.LabelPlus LabelPlus2 
      Height          =   1755
      Left            =   8580
      TabIndex        =   29
      Top             =   120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   3096
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   2
      Caption         =   "Form1.frx":0D9E
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotLine         =   -1  'True
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
   Begin DashBoard.ucProgressCircular ucProgressCircular1 
      Height          =   975
      Left            =   8700
      TabIndex        =   28
      Top             =   180
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Ton."
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   9
      PB_Border       =   -1  'True
      DisplayInPercent=   0   'False
      AnimationInterval=   100
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Carbón Dia"
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
      Left            =   8700
      TabIndex        =   27
      Top             =   1260
      Width           =   1215
   End
   Begin DashBoard.ucChartBar ChartBar4 
      Height          =   3435
      Left            =   17460
      TabIndex        =   26
      Top             =   2220
      Width           =   9675
      _ExtentX        =   17066
      _ExtentY        =   6059
      Title           =   "Producción por Operador "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Border          =   -1  'True
      ChartStyle      =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelsPositions =   3
      BorderRound     =   5
      BorderColor     =   -2147483638
   End
   Begin DashBoard.ucChartBar Traslados 
      Height          =   3435
      Left            =   10800
      TabIndex        =   25
      Top             =   2280
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6059
      Title           =   "Traslados Por Transportadora Dia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   6.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Border          =   -1  'True
      ChartStyle      =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderRound     =   5
      BorderColor     =   -2147483638
   End
   Begin DashBoard.ucChartArea VentasDia 
      Height          =   4455
      Left            =   300
      TabIndex        =   24
      Top             =   5940
      Width           =   18795
      _ExtentX        =   33152
      _ExtentY        =   7858
      Title           =   "Ventas Mensual "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Border          =   -1  'True
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
      BorderRound     =   5
      BorderColor     =   -2147483638
   End
   Begin DashBoard.ucPieChart PieVentas 
      Height          =   2115
      Left            =   19440
      TabIndex        =   23
      Top             =   5880
      Width           =   7695
      _ExtentX        =   13573
      _ExtentY        =   3731
      Title           =   "Ventas Dia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Border          =   -1  'True
      BorderColor     =   -2147483638
      LabelsVisible   =   -1  'True
      ChartStyle      =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
      LabelsPositions =   2
      BorderRound     =   5
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Despachos Dia"
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
      Left            =   22620
      TabIndex        =   22
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Despachos Mes"
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
      Left            =   24240
      TabIndex        =   21
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Despachos Año"
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
      Left            =   25860
      TabIndex        =   20
      Top             =   1320
      Width           =   1215
   End
   Begin DashBoard.ucChartBar Ventas 
      Height          =   3435
      Left            =   180
      TabIndex        =   19
      Top             =   2280
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   6059
      Title           =   "Despachos Por Transportadora Dia"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   6.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Border          =   -1  'True
      ChartStyle      =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderRound     =   5
      BorderColor     =   -2147483638
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Carbón Año"
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
      Left            =   14700
      TabIndex        =   18
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Carbón Mes"
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
      Left            =   13140
      TabIndex        =   17
      Top             =   1260
      Width           =   1215
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Carbón Dia"
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
      Left            =   11700
      TabIndex        =   16
      Top             =   1260
      Width           =   1215
   End
   Begin DashBoard.ucProgressCircular CarbonDia 
      Height          =   975
      Left            =   11700
      TabIndex        =   15
      Top             =   180
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Ton."
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   9
      PB_Color1       =   16501147
      PB_Border       =   -1  'True
      Max             =   4000
      Value           =   1000
      DisplayInPercent=   0   'False
      PF_ForeColor    =   13526537
      AnimationInterval=   100
   End
   Begin DashBoard.ucProgressCircular CarbonMes 
      Height          =   975
      Left            =   13080
      TabIndex        =   14
      Top             =   180
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Ton."
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   9
      PB_Color1       =   16501147
      PB_Border       =   -1  'True
      Max             =   90000
      Value           =   30000
      DisplayInPercent=   0   'False
      PF_ForeColor    =   13526537
      AnimationInterval=   100
   End
   Begin DashBoard.ucProgressCircular CarbonAnio 
      Height          =   975
      Left            =   14580
      TabIndex        =   13
      Top             =   180
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Ton."
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   9
      PB_Color1       =   16501147
      PB_Border       =   -1  'True
      Max             =   7.20000e5
      Value           =   1.20000e5
      DisplayInPercent=   0   'False
      PF_ForeColor    =   13526537
      AnimationInterval=   100
   End
   Begin DashBoard.LabelPlus LabelPlus9 
      Height          =   1755
      Left            =   11580
      TabIndex        =   12
      Top             =   120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   3096
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   2
      Caption         =   "Form1.frx":0DD2
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotLine         =   -1  'True
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
   Begin DashBoard.LabelPlus LabelPlus8 
      Height          =   1755
      Left            =   13080
      TabIndex        =   11
      Top             =   120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   3096
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   2
      Caption         =   "Form1.frx":0E06
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotLine         =   -1  'True
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
   Begin DashBoard.LabelPlus LabelPlus7 
      Height          =   1755
      Left            =   14580
      TabIndex        =   10
      Top             =   120
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   3096
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   2
      Caption         =   "Form1.frx":0E3C
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotLine         =   -1  'True
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
   Begin VB.Image Image1 
      Height          =   675
      Left            =   6240
      Picture         =   "Form1.frx":0E74
      Top             =   660
      Width           =   1545
   End
   Begin DashBoard.ucProgressCircular DespachosAnio 
      Height          =   975
      Left            =   25740
      TabIndex        =   9
      Top             =   300
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Ton."
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   9
      PB_Color1       =   10527742
      PB_Border       =   -1  'True
      Max             =   90000
      Value           =   30000
      DisplayInPercent=   0   'False
      PF_ForeColor    =   133614
      AnimationInterval=   100
   End
   Begin DashBoard.LabelPlus LabelPlus6 
      Height          =   1755
      Left            =   25740
      TabIndex        =   8
      Top             =   180
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   3096
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   2
      Caption         =   "Form1.frx":458E
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotLine         =   -1  'True
      HotLineColor    =   255
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
   Begin DashBoard.ucProgressCircular DespachosMes 
      Height          =   975
      Left            =   24120
      TabIndex        =   7
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1720
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Ton."
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   9
      PB_Color1       =   10527742
      PB_Border       =   -1  'True
      Max             =   90000
      Value           =   45000
      DisplayInPercent=   0   'False
      PF_ForeColor    =   133614
      AnimationInterval=   100
   End
   Begin DashBoard.ucProgressCircular DespachosDia 
      Height          =   975
      Left            =   22620
      TabIndex        =   6
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1720
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -5
      Caption2        =   "Ton."
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   9
      PB_Color1       =   10527742
      PB_Border       =   -1  'True
      Max             =   4000
      Value           =   1000
      DisplayInPercent=   0   'False
      PF_ForeColor    =   133614
      AnimationInterval=   100
   End
   Begin DashBoard.LabelPlus LabelPlus5 
      Height          =   1755
      Left            =   24180
      TabIndex        =   5
      Top             =   180
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   3096
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   2
      Caption         =   "Form1.frx":45C6
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotLine         =   -1  'True
      HotLineColor    =   255
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
   Begin DashBoard.LabelPlus LabelPlus4 
      Height          =   1755
      Left            =   22560
      TabIndex        =   4
      Top             =   180
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   3096
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   2
      Caption         =   "Form1.frx":45FC
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CallOutPosicion =   3
      HotLine         =   -1  'True
      HotLineColor    =   255
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "DASHBOAR LOGINEXT"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   435
      Left            =   2100
      TabIndex        =   3
      Top             =   240
      Width           =   3855
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Caption         =   "Datos de Producción Acopio El Brillante"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   315
      Index           =   0
      Left            =   1500
      TabIndex        =   2
      Top             =   720
      Width           =   4875
   End
   Begin DashBoard.LabelPlus LabelPlus1 
      Height          =   1815
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   3201
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   2
      CaptionAlignmentV=   2
      Caption         =   "Form1.frx":4630
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconPaddingX    =   10
      IconPaddingY    =   10
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":4650
   End
   Begin DashBoard.ucChartBar ChartBar3 
      Height          =   2175
      Left            =   19440
      TabIndex        =   0
      Top             =   8220
      Width           =   7635
      _ExtentX        =   13467
      _ExtentY        =   3836
      Title           =   "Doble Troque Por Hora Turno "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Border          =   -1  'True
      ChartStyle      =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelsPositions =   3
      BorderRound     =   5
      BorderColor     =   -2147483638
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cPalette As Collection
Dim i As Long, j As Long
Dim Value As Collection
Dim Lables As Collection
Dim Icons As Collection
Dim keys As Collection
Dim CustomColors As Collection
Dim colDate As Collection
Dim cResizer As ClsResizer
Dim Turno As Integer
Dim FechaTurno As Date
Dim Conn As New ADODB.Connection
Dim DisplayIndex As Integer
Dim Pasadas As Long
Dim MyNow As Date

Private Sub Form_Load()
Dim xFile, xA, xDb, xRr, xCb, xDbM, xRrM As String, i As Integer
Dim xSr As String, xPs As String
Dim xM As New ADODB.Recordset

On Error GoTo Recover

    If App.PrevInstance = True Then
        MsgBox "El programa ya está siendo ejecutado"
        End
    End If
        
'    Set cResizer = New ClsResizer
'    With cResizer
'        .AddControlFont "ucProgressCircular", "Caption1_Font", "Caption2_Font"
'
'        .AddControlFont "ucChartArea", "Font", "TitleFont"
'        .AddControlFont "ucChartBar", "Font", "TitleFont"
'        .AddControlFont "ucPieChart", "Font", "TitleFont"
'        .AddControlFont "ucTreeMaps", "Font", "TitleFont"
'        .AddControlFont "LabelPlus", "Font"
'
'        .AddControlProperty "ucProgressCircular", "PB_Width", "PF_Width", "Caption1_OffsetY", "Caption2_OffsetY"
'        .AddControlProperty "LabelPlus", "HotLineWidth"
'        .AddControlProperty "ucChartArea", "LinesWidth"
'        .AddControlProperty "ucPieChart", "SeparatorLineWidth", "DonutWidth"
'
'        .SaveControlsPositions Me
'    End With
'
    Set cPalette = NewCollection(&H4744E3, &H50C187, &HABA56C, &H48BDBF, &H4D91F4, &H7450, &H3DB0EF, vbBlue, vbGreen, vbRed, vbYellow, vbCyan, &H4D91F4, &H48BDBF, &H50C187, &HABA56C)
        
    Conn.Provider = "SQLOLEDB"
    'Conn.Properties("Integrated Security") = SSPI
    Conn.Properties("Data Source") = "CIFSRV0001"
    'Conn.Properties("Data Source") = "LOCALHOST"
    Conn.Properties("Initial Catalog") = "Tracer"
    Conn.Properties("user ID") = "sisma_app"
    Conn.Properties("password") = "V$123bcd"
    Conn.CursorLocation = adUseServer
    Conn.CommandTimeout = 0
    Conn.Open

    
    MyNow = Now 'CDate(Format(#8/31/2022 5:30:00 PM#, "dd/MM/yyyy hh:mm"))
    Label3(1).Caption = Format(MyNow, "dd-MMMM-YYYY")
    DisplayIndex = 0
    Pasadas = 0
    
    Call Display

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error," & vbCrLf & Err.Description
    MsgBox MSG, , "DASHBOARD"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub

Private Sub Display()
Dim xMin As Integer
Dim i As Long

        Call CamionesHora
        Call ProducionCarbonTurnos
        Call VentasCliente
        Call TrasladosAcopios
        Call PorcentajeVentas
        Call VentasporDia
        Call ProduccionMes
        Call DespachosDiaChar
        Call DespachosHora
        
End Sub

Private Sub DespachosHora()
    Dim Value   As Collection
    Dim colHoras As Collection
    Dim i As Integer, j As Integer
    Dim xSql As String
    Dim xDe As New ADODB.Recordset
    Dim Turno As Integer, oFFset As Integer
    Dim FechaTurno As Date
    Dim xOp As New ADODB.Recordset
    Dim Acum As Long
    
    On Error GoTo Recover
    
    Char_DespachosHora.Clear
    Char_DespachosHora.Title = "Despachos Hora (" & Format(MyNow, "MMM-dd") & ")"
    
    Set colHoras = New Collection
    j = 3
    oFFset = 0
    
    Turno = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(MyNow, "dd/MM/yyyy hh:mm") & "',1").Fields(0)
    FechaTurno = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(MyNow, "dd/MM/yyyy hh:mm") & "',2").Fields(0)

    UltimoDia = Conn.Execute("Select DAY(EOMONTH(GetDate())) AS DaysInMonth").Fields(0)
    
    For i = 6 To 23
        colHoras.Add (i)
    Next i
    For i = 0 To 5
        colHoras.Add (i)
    Next i
    
    Char_DespachosHora.AddAxisItems colHoras, , 305, 2
    
    xSql = "SET DateFormat DMY Select * From FT_DB_DespachosHora('" & Format(MyNow, "dd/MM/yyyy") & "', 2)"
    
    Set xOp = Conn.Execute(xSql)
    Set Value = New Collection
    Acum = 0
    
    If Not xOp.EOF Then
        For i = 1 To 24
            Value.Add xOp.Fields(i).Value
            Acum = Acum + xOp.Fields(i).Value
        Next i
    End If
    Char_DespachosHora.AddSerie "DS" & " - " & Format(Acum, "#,###") & " viajes", vbBlue, Value
    xOp.Close

Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error, DespachosDia() " & vbCrLf & Err.Description
    MsgBox MSG, , "DASHBOARD"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub


Private Sub DespachosDiaChar()
    Dim Value   As Collection
    Dim colHoras As Collection
    Dim i As Integer, j As Integer
    Dim xSql As String
    Dim xDe As New ADODB.Recordset
    Dim Turno As Integer, oFFset As Integer
    Dim FechaTurno As Date
    Dim xOp As New ADODB.Recordset
    Dim Acum As Long
    
    On Error GoTo Recover
    
    Char_DespachosDia.Clear
    Char_DespachosDia.Title = "Despachos Dia (" & MonthName(Month(MyNow)) & ")"
    
    Set colHoras = New Collection
    j = 2
    oFFset = 0
    
    Turno = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(MyNow, "dd/MM/yyyy hh:mm") & "',1").Fields(0)
    FechaTurno = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(MyNow, "dd/MM/yyyy hh:mm") & "',2").Fields(0)

    UltimoDia = Conn.Execute("Select DAY(EOMONTH(GetDate())) AS DaysInMonth").Fields(0)
    
    For i = 0 To UltimoDia
        colHoras.Add (i + 1)
    Next i

    Char_DespachosDia.AddAxisItems colHoras, , 305, 2
    
    xSql = "SET DateFormat DMY Select * From FT_DB_DespachosDia('" & Format(MyNow, "dd/MM/yyyy") & "', 2)"
    
    Set xOp = Conn.Execute(xSql)
    Set Value = New Collection
    Acum = 0
    
    Do While Not xOp.EOF
        Value.Add xOp.Fields(1).Value
        Acum = Acum + xOp.Fields(1).Value
        xOp.MoveNext
    Loop
    Char_DespachosDia.AddSerie "DS" & " - " & Format(Acum, "#,###") & " viajes", vbGreen, Value
    xOp.Close

Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error, DespachosDia() " & vbCrLf & Err.Description
    MsgBox MSG, , "DASHBOARD"
    Err.Clear   ' Borra campos del objeto Err
    Exit Sub
End If

End Sub

Private Sub CamionesHora()

    Dim Value   As Collection
    Dim colHoras As Collection
    Dim i As Integer, j As Integer
    Dim xSql As String
    Dim xDe As New ADODB.Recordset
    Dim Turno As Integer, oFFset As Integer
    Dim FechaTurno As Date
    Dim xOp As New ADODB.Recordset
    
    ChartBar3.Clear
    
    Set colHoras = New Collection
    j = 6
    oFFset = 0
    
    Turno = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(MyNow, "dd/MM/yyyy hh:mm") & "',1").Fields(0)
    FechaTurno = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(MyNow, "dd/MM/yyyy hh:mm") & "',2").Fields(0)

    ChartBar3.Title = "Camiones x Hora Turno " + Str(Turno)

    For i = 0 To 11
        If Turno = 1 Then
            colHoras.Add Format(i + 6, "00")
        Else
            colHoras.Add Format(i + 18 - oFFset, "00")
            If i + 18 = 23 Then oFFset = 24
        End If
    Next
    
    ChartBar3.AddAxisItems colHoras, , 305, 2
    
    Set xOp = Conn.Execute("Select * From OperadoresMineros")
    Do While Not xOp.EOF
        If Turno = 1 Then
            xSql = "Set DateFormat DMY Select * From FT_DB_RecepcionHora('" & Turno & "','" & Format(FechaTurno, "dd/MM/yyyy") & "', 'LT'" & "," & xOp!IdOperador & ")"
        Else
            xSql = "Set DateFormat DMY Select * From FT_DB_RecepcionHora2('" & Turno & "','" & Format(FechaTurno, "dd/MM/yyyy") & "', 'LT'" & "," & xOp!IdOperador & ")"
        End If
        Set xDe = Conn.Execute(xSql)
        
        Set Value = New Collection
        For i = 1 To 12
            Value.Add xDe.Fields(i).Value
        Next
    
        ChartBar3.AddSerie xOp!Descripcion, cPalette(j), Value
        j = j + 1
        xOp.MoveNext
    Loop
    xOp.Close
End Sub

Private Sub ProducionCarbonTurnos()
Dim xSql As String
Dim xR As New ADODB.Recordset

On Error GoTo Recover

FI = Format(MyNow, "dd/MM/yyyy")
FF = Format(MyNow, "dd/MM/yyyy") & " 23:59:59"

xSql = "SET DateFormat DMY SELECT * FROM FT_DB_CarbonPlanTurno('" & Format(MyNow, "dd/MM/yyyy") & "',1,1)"
Set xR = Conn.Execute(xSql)
LabelPlus2.Caption = "Plan " & Format(xR!Pto / 2, "##,###")
CarbonTurno1.Value = xR!Neto

xSql = "SET DateFormat DMY SELECT * FROM FT_DB_CarbonPlanTurno('" & Format(MyNow, "dd/MM/yyyy") & "',2,1)"
Set xR = Conn.Execute(xSql)
LabelPlus3.Caption = "Plan " & Format(xR!Pto / 2, "##,###")
CarbonTurno2.Value = xR!Neto

xSql = "SET DateFormat DMY SELECT * FROM FT_DB_CarbonPlanDIA('" & Format(MyNow, "dd/MM/yyyy") & "',1)"
Set xR = Conn.Execute(xSql)
LabelPlus9.Caption = "Plan " & Format(xR!Pto, "##,###")
CarbonDia.Value = xR!Neto

xSql = "SET DateFormat DMY SELECT * FROM FT_DB_CarbonPlanMES('" & Format(MyNow, "dd/MM/yyyy") & "',1)"
Set xR = Conn.Execute(xSql)
LabelPlus8.Caption = "Plan " & Format(xR!Pto, "##,###")
CarbonMes.Value = xR!Neto

xSql = "SET DateFormat DMY SELECT * FROM FT_DB_CarbonPlanANIO('" & Format(MyNow, "dd/MM/yyyy") & "',1)"
Set xR = Conn.Execute(xSql)
LabelPlus7.Caption = "Plan " & Format(xR!Pto, "##,###")
CarbonAnio.Value = xR!Neto
'======================================= DESPACHOS ================================================================

xSql = "SET DateFormat DMY SELECT * FROM FT_DB_CarbonPlanTurno('" & Format(MyNow, "dd/MM/yyyy") & "',1,2)"
Set xR = Conn.Execute(xSql)
LabelPlus11.Caption = "Plan " & Format(xR!Pto / 2, "##,###")
DespachosTurno1.Value = xR!Neto

xSql = "SET DateFormat DMY SELECT * FROM FT_DB_CarbonPlanTurno('" & Format(MyNow, "dd/MM/yyyy") & "',2,2)"
Set xR = Conn.Execute(xSql)
LabelPlus10.Caption = "Plan " & Format(xR!Pto / 2, "##,###")
DespachosTurno2.Value = xR!Neto

xSql = "SET DateFormat DMY SELECT * FROM FT_DB_CarbonPlanDIA('" & Format(MyNow, "dd/MM/yyyy") & "',2)"
Set xR = Conn.Execute(xSql)
LabelPlus4.Caption = "Plan " & Format(xR!Pto, "##,###")
DespachosDia.Value = xR!Neto

xSql = "SET DateFormat DMY SELECT * FROM FT_DB_CarbonPlanMES('" & Format(MyNow, "dd/MM/yyyy") & "',2)"
Set xR = Conn.Execute(xSql)
LabelPlus5.Caption = "Plan " & Format(xR!Pto, "##,###")
DespachosMes.Value = xR!Neto

xSql = "SET DateFormat DMY SELECT * FROM FT_DB_CarbonPlanANIO('" & Format(MyNow, "dd/MM/yyyy") & "',2)"
Set xR = Conn.Execute(xSql)
LabelPlus6.Caption = "Plan " & Format(xR!Pto, "##,###")
DespachosAnio.Value = xR!Neto

xR.Close

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error," & vbCrLf & Err.Description
    MsgBox MSG, , "ProducionCarbonTurnos()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub

Private Sub VentasCliente()
    Dim Value As Collection
    Dim colHoras As Collection
    Dim i As Integer, j As Integer
    Dim xSql As String
    Dim xTR As String
    Dim xDe As New ADODB.Recordset
    Dim xR As New ADODB.Recordset
    Dim Acum As Long
        
On Error GoTo Recover
    
    Ventas.Clear
    
    Set Value = New Collection
    Set colHoras = New Collection
    
    xSql = "Set DateFormat DMY "
    xSql = xSql + " SELECT DISTINCT Transportador.IdTransportador, Transportador.DescripcionCorta"
    xSql = xSql + " FROM   Bascula INNER JOIN"
    xSql = xSql + "        Ventas ON Bascula.NumeroTransaccion = Ventas.IdVentas INNER JOIN"
    xSql = xSql + "        Terceros ON Ventas.IdCliente = Terceros.IdCliente INNER JOIN"
    xSql = xSql + "        Transportador ON Bascula.IdTransportador = Transportador.IdTransportador"
    xSql = xSql + " WHERE  (Bascula.TransaccionOrigen = 'DS'  ) AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND (YEAR(Bascula.FechaTurno) = " & Year(MyNow) & ") AND (MONTH(Bascula.FechaTurno)) = " & Month(MyNow) & " AND (DAY(Bascula.FechaTurno) = " & Day(MyNow) & ")"
    xSql = xSql + " GROUP BY Transportador.IdTransportador, Transportador.DescripcionCorta"
    xSql = xSql + " ORDER BY Transportador.IdTransportador"
       
    Set xDe = Conn.Execute(xSql)
    
    If xDe.EOF Then Exit Sub
    
    Do While Not xDe.EOF
        xTR = xTR + Str(xDe.Fields(0).Value) + ", "
        colHoras.Add xDe.Fields(1).Value
        xDe.MoveNext
    Loop
       
    xTR = Mid(xTR, 1, Len(xTR) - 1)
    Ventas.AddAxisItems colHoras, , 305, 2
       
    xSql = "Set DateFormat DMY "
    xSql = xSql + " SELECT DISTINCT Terceros.Descripcion, Terceros.IdCliente"
    xSql = xSql + " FROM    Bascula INNER JOIN"
    xSql = xSql + "         Ventas ON Bascula.NumeroTransaccion = Ventas.IdVentas INNER JOIN"
    xSql = xSql + "         Terceros ON Ventas.IdCliente = Terceros.IdCliente INNER JOIN"
    xSql = xSql + "         Transportador ON Bascula.IdTransportador = Transportador.IdTransportador"
    xSql = xSql + " WHERE  (Bascula.TransaccionOrigen = 'DS' ) AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND (YEAR(Bascula.FechaTurno) =" & Year(MyNow) & ") AND (MONTH(Bascula.FechaTurno)) =" & Month(MyNow) & " AND (DAY(Bascula.FechaTurno) = " & Day(MyNow) & ")"
    xSql = xSql + " GROUP BY Terceros.Descripcion, Terceros.IdCliente"
    xSql = xSql + " Order By Terceros.IdCliente, Terceros.Descripcion"
    
    Set xDe = Conn.Execute(xSql)
                  
    j = 1
    Do While Not xDe.EOF
        xSql = "SELECT * FROM FT_DB_VentasPorTransportador(" & xDe!IdCliente & ",'" & xTR & "'," & Year(MyNow) & "," & Month(MyNow) & "," & Day(MyNow) & ")"
        Set xR = Conn.Execute(xSql)
        i = 0
        Acum = 0
        Do While Not xR.EOF
            Value.Add xR.Fields(1).Value
            Acum = Acum + xR.Fields(1).Value
            xR.MoveNext
            i = i + 1
         Loop
         Ventas.AddSerie xDe!Descripcion & " -" & Format(Acum, "##0") & " Ton.", cPalette(j), Value
         j = j + 1
         xDe.MoveNext
         Set Value = New Collection
         'If j = 7 Then Exit Do
    Loop

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error, VentasClientes" & vbCrLf & Err.Description
    MsgBox MSG, , "DASHBOARD"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub

Private Sub TrasladosAcopios()
    Dim Value As Collection
    Dim colHoras As Collection
    Dim i As Integer, j As Integer
    Dim xSql As String
    Dim xTR As String
    Dim xDe As New ADODB.Recordset
    Dim xR As New ADODB.Recordset
    Dim Acum As Long
        
On Error GoTo Recover
        
    Traslados.Clear
    
    Set Value = New Collection
    Set colHoras = New Collection

    xSql = "Set DateFormat DMY "
    xSql = xSql + " SELECT DISTINCT Bascula.IdTransportador, Transportador.Descripcion"
    xSql = xSql + " FROM     Bascula INNER JOIN"
    xSql = xSql + "                  Transportador ON Bascula.IdTransportador = Transportador.IdTransportador"
    xSql = xSql + " WHERE  (Bascula.TransaccionOrigen = 'TR'  ) AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND (YEAR(Bascula.FechaTurno) =" & Year(MyNow) & ") AND (MONTH(Bascula.FechaTurno)) =" & Month(MyNow) & " AND (DAY(Bascula.FechaTurno) = " & Day(MyNow) & ")"
    
    Set xDe = Conn.Execute(xSql)
    If xDe.EOF Then Exit Sub
    
    Do While Not xDe.EOF
        xTR = xTR + Str(xDe.Fields(0).Value) + ", "
        colHoras.Add xDe.Fields(1).Value
        xDe.MoveNext
    Loop
       
    xTR = Mid(xTR, 1, Len(xTR) - 1)
    Traslados.AddAxisItems colHoras, , 305, 2
    
    xSql = "Set DateFormat DMY "
    xSql = xSql + " SELECT  DISTINCT Acopios.IdAcopio, Acopios.Descripcion"
    xSql = xSql + " FROM      Bascula INNER JOIN"
    xSql = xSql + "                  Acopios INNER JOIN"
    xSql = xSql + "                  Pilas ON Acopios.IdAcopio = Pilas.IdAcopio INNER JOIN"
    xSql = xSql + "                  Traslados ON Pilas.IdPila = Traslados.PilaDestino ON Bascula.NumeroTransaccion = Traslados.IdTraslado"
    xSql = xSql + " WHERE  (Bascula.TransaccionOrigen = 'TR'  ) AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND (YEAR(Bascula.FechaTurno) =" & Year(MyNow) & ") AND (MONTH(Bascula.FechaTurno)) =" & Month(MyNow) & " AND (DAY(Bascula.FechaTurno) = " & Day(MyNow) & ")"
    
    Set xDe = Conn.Execute(xSql)
    
    j = 1
    Do While Not xDe.EOF
        xSql = "SELECT * FROM FT_DB_TrasladoPorTransportador(" & xDe!IdAcopio & ",'" & xTR & "'," & Year(MyNow) & "," & Month(MyNow) & " ," & Day(MyNow) & ")"
        Set xR = Conn.Execute(xSql)
        i = 0
        Acum = 0
        Do While Not xR.EOF
            Value.Add xR.Fields(1).Value
            Acum = Acum + xR.Fields(1).Value
            xR.MoveNext
            i = i + 1
         Loop
          Traslados.AddSerie xDe!Descripcion & "-" & Format(Acum, "#,###") & " Ton.", cPalette(j), Value
          j = j + 1
          xDe.MoveNext
          Set Value = New Collection
    Loop

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error, TrasladosAcopio" & vbCrLf & Err.Description
    MsgBox MSG, , "DASHBOARD"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
    
End Sub

Private Sub PorcentajeVentas()

Dim i As Integer
Dim xD As New ADODB.Recordset
Dim xSql As String
On Error Resume Next

xSql = "Set DateFormat DMY "
 xSql = xSql + " SELECT ISNULL( ROUND(SUM(Bascula.PesoLleno - Bascula.PesoVacio)/1000,0),0) AS Neto, Terceros.Descripcion"
 xSql = xSql + " FROM    Bascula INNER JOIN"
 xSql = xSql + "                 Ventas ON Bascula.NumeroTransaccion = Ventas.IdVentas INNER JOIN"
 xSql = xSql + "                 Terceros ON Ventas.IdCliente = Terceros.IdCliente"
 xSql = xSql + " WHERE  (Bascula.TransaccionOrigen = 'DS') AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND YEAR(Bascula.FechaTurno) = " & Year(MyNow) & " AND MONTH(Bascula.FechaTurno)=" & Month(MyNow) & " AND DAY(Bascula.FechaTurno)=" & Day(MyNow)
 xSql = xSql + " GROUP BY Terceros.Descripcion"

Set xD = Conn.Execute(xSql)

PieVentas.Clear

i = 1
Do While Not xD.EOF
    PieVentas.AddItem xD.Fields(1).Value, xD.Fields(0).Value, cPalette(i), False
    xD.MoveNext
    i = i + 1
Loop
xD.Close

End Sub

Private Sub VentasporDia()
    Dim Value As Collection
    Dim Value2 As Collection
    Dim colHoras As Collection
    Dim i As Integer, UltimoDia As Integer, j As Integer
    Dim xSql As String
    Dim xDe As New ADODB.Recordset
    Dim xR As New ADODB.Recordset
    Dim xNeto As Single
    Dim Mes As Integer, Anio As Integer
    
    VentasDia.Clear
    
    Set Value = New Collection
    Set Value2 = New Collection
    UltimoDia = Conn.Execute("Select DAY(EOMONTH(GetDate())) AS DaysInMonth").Fields(0)
    
    For i = 1 To UltimoDia
            Value2.Add (i)
    Next i
    VentasDia.AddAxisItems Value2
    
    Mes = Month(MyNow)
    Anio = Year(MyNow)
    
    xSql = "    SELECT DISTINCT Terceros.Descripcion,Terceros.IdCliente"
    xSql = xSql + " FROM       Bascula INNER JOIN"
    xSql = xSql + "                   Ventas ON Bascula.NumeroTransaccion = Ventas.IdVentas INNER JOIN"
    xSql = xSql + "                   Terceros ON Ventas.IdCliente = Terceros.IdCliente INNER JOIN"
    xSql = xSql + "                   Transportador ON Bascula.IdTransportador = Transportador.IdTransportador"
    xSql = xSql + " WHERE  (Bascula.TransaccionOrigen = 'DS'  ) AND (Bascula.Estado = 'AC') AND (Bascula.IdMaterial = 1) AND (YEAR(Bascula.FechaTurno) = " & Anio & ") AND (MONTH(Bascula.FechaTurno)) = " & Mes & "  "
    xSql = xSql + " GROUP BY Terceros.Descripcion, Terceros.IdCliente"
    xSql = xSql + " Order By Terceros.IdCliente, Terceros.Descripcion"
        
    Set xR = Conn.Execute(xSql)
    j = j + 1
    Do While Not xR.EOF
        Set Value = New Collection
        xSql = "Select * From FT_DB_VentasClienteMes(" & xR!IdCliente & ", " & Anio & ", " & Mes & ")"
        Set xDe = Conn.Execute(xSql)
        For i = 1 To UltimoDia
            Value.Add xDe.Fields(i).Value
        Next
        VentasDia.AddLineSeries xR!Descripcion, Value, cPalette(j)
        j = j + 1
        xR.MoveNext
    Loop
End Sub

Private Sub ProduccionMes()

    Dim Value   As Collection
    Dim colHoras As Collection
    Dim i As Integer, j As Integer
    Dim xSql As String
    Dim xDe As New ADODB.Recordset
    Dim Turno As Integer, oFFset As Integer
    Dim FechaTurno As Date
    Dim xOp As New ADODB.Recordset
    Dim Acum As Long
    
    On Error GoTo Recover
    
    ChartBar4.Clear
    ChartBar4.Title = "Producción x Operador mes de " + MonthName(Month(MyNow))
    
    Set colHoras = New Collection
    j = 6
    oFFset = 0
    
    Turno = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(MyNow, "dd/MM/yyyy hh:mm") & "',1").Fields(0)
    FechaTurno = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(MyNow, "dd/MM/yyyy hh:mm") & "',2").Fields(0)

    For i = 0 To 30
        colHoras.Add Format(i + 1, "00")
    Next
    
    ChartBar4.AddAxisItems colHoras, , 305, 2
    
    Set xOp = Conn.Execute("Select * From OperadoresMineros")
    Do While Not xOp.EOF
        xSql = "Set DateFormat DMY Select * From FT_DB_RecepcionHora_Cantidad('" & Format(FechaTurno, "dd/MM/yyyy") & "', 'LT'" & "," & xOp!IdOperador & ")"
        Set xDe = Conn.Execute(xSql)
        Acum = 0
        Set Value = New Collection
        For i = 1 To 31
            Value.Add xDe.Fields(i).Value
            Acum = Acum + xDe.Fields(i).Value
        Next

        ChartBar4.AddSerie xOp!Descripcion & " -" & Format(Acum, "#,###") & " Ton.", cPalette(j), Value
        j = j + 1
        xOp.MoveNext
    Loop
    xOp.Close

Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error, ProduccionMes() " & vbCrLf & Err.Description
    MsgBox MSG, , "DASHBOARD"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
    
End Sub


Private Function NewCollection(ParamArray vArgList() As Variant) As Collection
    Dim Value As Variant
    Set NewCollection = New Collection
    For Each Value In vArgList
        NewCollection.Add Value
    Next
End Function

Private Sub Form_Resize()
   ' cResizer.ResizeControls Me
End Sub

Private Function Random(ByVal Min!, ByVal Max!)
    Random = Int((Max - Min + 1) * Rnd + Min)
End Function

Private Sub Timer1_Timer()
MyNow = Now 'CDate(Format(#8/31/2022 5:30:00 PM#, "dd/MM/yyyy hh:mm"))

Pasadas = Pasadas + 1
LabelPlus1.Caption = Format(Pasadas, "###,###")

Select Case DisplayIndex
Case 0
        Call CamionesHora
Case 1
       Call ProducionCarbonTurnos
Case 2
        Call VentasCliente
Case 3
        Call TrasladosAcopios
Case 4
'        Call PorcentajeVentas
Case 5
    Call VentasporDia
Case 6
    Call ProduccionMes
Case 7
    Call DespachosDiaChar
Case 8
    Call DespachosHora
End Select

DisplayIndex = DisplayIndex + 1
If DisplayIndex > 8 Then DisplayIndex = 0

End Sub



 
