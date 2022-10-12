VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OPERATIVO"
   ClientHeight    =   13260
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   27315
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   13260
   ScaleWidth      =   27315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   6720
      Top             =   1380
   End
   Begin VB.Label DatoPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   7
      Left            =   17580
      TabIndex        =   101
      Top             =   6000
      Width           =   675
   End
   Begin VB.Label DatoPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 2"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   6
      Left            =   17580
      TabIndex        =   100
      Top             =   5640
      Width           =   675
   End
   Begin VB.Label DatoPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 3"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   5
      Left            =   17580
      TabIndex        =   99
      Top             =   5340
      Width           =   675
   End
   Begin VB.Label DatoPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 4"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   4
      Left            =   17580
      TabIndex        =   98
      Top             =   5040
      Width           =   675
   End
   Begin VB.Label LabelPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   7
      Left            =   15900
      TabIndex        =   97
      Top             =   5940
      Width           =   1635
   End
   Begin VB.Label LabelPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 2"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   6
      Left            =   15900
      TabIndex        =   96
      Top             =   5640
      Width           =   1635
   End
   Begin VB.Label LabelPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 3"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   5
      Left            =   15900
      TabIndex        =   95
      Top             =   5340
      Width           =   1635
   End
   Begin VB.Label LabelPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 2"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   4
      Left            =   15900
      TabIndex        =   94
      Top             =   5040
      Width           =   1635
   End
   Begin DashBoard.LabelPlus PowerScreen 
      Height          =   1875
      Index           =   1
      Left            =   14340
      TabIndex        =   93
      Top             =   4680
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   3307
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      Caption         =   "Form1.frx":0D02
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      GradientColor1  =   16776960
      PicturePaddingY =   20
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      HotLine         =   -1  'True
      HotLineColor    =   16761087
      HotLineWidth    =   20
      HotLinePosition =   1
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":0D6A
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   18
      Left            =   5940
      TabIndex        =   91
      Top             =   11220
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":4F2C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   18
      Left            =   4680
      TabIndex        =   90
      Top             =   10860
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":4F4C
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":4F96
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   13
      Left            =   11040
      TabIndex        =   89
      Top             =   2820
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":77C4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   13
      Left            =   9720
      TabIndex        =   88
      Top             =   2460
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":77E4
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":7842
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   12
      Left            =   5940
      TabIndex        =   87
      Top             =   6300
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":A208
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   12
      Left            =   4620
      TabIndex        =   86
      Top             =   5940
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":A22A
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":A288
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   11
      Left            =   6000
      TabIndex        =   85
      Top             =   5100
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":CC4E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   11
      Left            =   4620
      TabIndex        =   84
      Top             =   4800
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":CC6E
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":CCCC
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   8
      Left            =   1860
      TabIndex        =   83
      Top             =   11940
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":F692
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   8
      Left            =   600
      TabIndex        =   82
      Top             =   11580
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":F6B2
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":F710
   End
   Begin DashBoard.ucChartArea Tipo3 
      Height          =   3495
      Left            =   19020
      TabIndex        =   81
      Top             =   9300
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   6165
      Title           =   "Produccion T3"
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
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
      BorderRound     =   5
      BorderColor     =   -2147483638
   End
   Begin DashBoard.ucChartArea Tipo2 
      Height          =   3495
      Left            =   18900
      TabIndex        =   80
      Top             =   5580
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   6165
      Title           =   "Produccion T2"
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
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
      BorderRound     =   5
      BorderColor     =   -2147483638
   End
   Begin DashBoard.ucChartArea Tipo1 
      Height          =   3375
      Left            =   18840
      TabIndex        =   79
      Top             =   2100
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   5953
      Title           =   "Produccion T1"
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
         Name            =   "Segoe UI"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
      BorderRound     =   5
      BorderColor     =   -2147483638
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   17
      Left            =   6000
      TabIndex        =   78
      Top             =   7980
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":120D6
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   17
      Left            =   4680
      TabIndex        =   77
      Top             =   7620
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":120F6
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":12154
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   16
      Left            =   5940
      TabIndex        =   76
      Top             =   9060
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":14982
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   16
      Left            =   4680
      TabIndex        =   75
      Top             =   8700
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":149A2
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":14A00
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   15
      Left            =   6000
      TabIndex        =   74
      Top             =   10140
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":1722E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   15
      Left            =   4680
      TabIndex        =   73
      Top             =   9780
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":1724E
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":172AC
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   14
      Left            =   11160
      TabIndex        =   72
      Top             =   5100
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":19ADA
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   14
      Left            =   9840
      TabIndex        =   71
      Top             =   4740
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":19AFA
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":19B58
   End
   Begin DashBoard.LabelPlus Acopio 
      Height          =   2655
      Index           =   2
      Left            =   9360
      TabIndex        =   70
      Top             =   4260
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   4683
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":1C51E
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      HotLineColor    =   33023
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
   Begin DashBoard.LabelPlus Acopio 
      Height          =   1875
      Index           =   1
      Left            =   9360
      TabIndex        =   69
      Top             =   2100
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   3307
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":1C564
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      HotLineColor    =   16711935
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
   Begin VB.Label DatoPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 4"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   3
      Left            =   17580
      TabIndex        =   68
      Top             =   3900
      Width           =   675
   End
   Begin VB.Label DatoPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 3"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   2
      Left            =   17580
      TabIndex        =   67
      Top             =   3600
      Width           =   675
   End
   Begin VB.Label DatoPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 2"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   1
      Left            =   17580
      TabIndex        =   66
      Top             =   3300
      Width           =   675
   End
   Begin VB.Label DatoPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   0
      Left            =   17580
      TabIndex        =   65
      Top             =   3000
      Width           =   675
   End
   Begin VB.Label LabelPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 2"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   3
      Left            =   15900
      TabIndex        =   64
      Top             =   3840
      Width           =   1635
   End
   Begin VB.Label LabelPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 3"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   2
      Left            =   15900
      TabIndex        =   63
      Top             =   3540
      Width           =   1635
   End
   Begin VB.Label LabelPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 2"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   1
      Left            =   15900
      TabIndex        =   62
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label LabelPS 
      BackColor       =   &H8000000E&
      Caption         =   "DATO 1"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   0
      Left            =   15900
      TabIndex        =   61
      Top             =   2940
      Width           =   1695
   End
   Begin DashBoard.LabelPlus PowerScreen 
      Height          =   1875
      Index           =   0
      Left            =   14280
      TabIndex        =   60
      Top             =   2520
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   3307
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      Caption         =   "Form1.frx":1C5AA
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      GradientColor1  =   16776960
      PicturePaddingY =   20
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      HotLine         =   -1  'True
      HotLineColor    =   16761087
      HotLineWidth    =   20
      HotLinePosition =   1
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":1C612
   End
   Begin DashBoard.LabelPlus Clasificadoras 
      Height          =   4815
      Index           =   0
      Left            =   14100
      TabIndex        =   59
      Top             =   2100
      Width           =   4515
      _ExtentX        =   7964
      _ExtentY        =   8493
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":207D4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      HotLineColor    =   65280
      HotLineWidth    =   20
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
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   10
      Left            =   6060
      TabIndex        =   58
      Top             =   3960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":20810
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   10
      Left            =   4620
      TabIndex        =   57
      Top             =   3600
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":20830
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":20870
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   9
      Left            =   6000
      TabIndex        =   56
      Top             =   2880
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":23236
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   9
      Left            =   4560
      TabIndex        =   55
      Top             =   2520
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":23256
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":2329E
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   7
      Left            =   1980
      TabIndex        =   54
      Top             =   10740
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":25C64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   7
      Left            =   540
      TabIndex        =   53
      Top             =   10380
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":25C84
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":25CE2
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   6
      Left            =   1980
      TabIndex        =   52
      Top             =   9600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":286A8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   6
      Left            =   540
      TabIndex        =   51
      Top             =   9240
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":286C8
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":28726
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   5
      Left            =   2040
      TabIndex        =   50
      Top             =   8520
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":2B0EC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   5
      Left            =   600
      TabIndex        =   49
      Top             =   8160
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":2B10C
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":2B16A
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   4
      Left            =   2040
      TabIndex        =   48
      Top             =   7440
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":2DB30
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   4
      Left            =   600
      TabIndex        =   47
      Top             =   7080
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":2DB50
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":2DBAE
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   3
      Left            =   1980
      TabIndex        =   46
      Top             =   6300
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":30574
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   3
      Left            =   600
      TabIndex        =   45
      Top             =   5940
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":30594
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":305F2
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   2
      Left            =   1980
      TabIndex        =   44
      Top             =   5160
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":32E20
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   2
      Left            =   600
      TabIndex        =   43
      Top             =   4800
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":32E40
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":32E9E
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   1
      Left            =   1920
      TabIndex        =   42
      Top             =   4020
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":356CC
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   1
      Left            =   540
      TabIndex        =   41
      Top             =   3660
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":356EC
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":3574A
   End
   Begin DashBoard.LabelPlus SaldoPila 
      Height          =   435
      Index           =   0
      Left            =   1920
      TabIndex        =   40
      Top             =   2880
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   767
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form1.frx":37F78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
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
      PictureArr      =   0
   End
   Begin DashBoard.LabelPlus Pila 
      Height          =   975
      Index           =   0
      Left            =   540
      TabIndex        =   39
      Top             =   2520
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   1720
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":37F98
      CaptionPaddingY =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16576
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":37FF6
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   315
      Index           =   1
      Left            =   2220
      TabIndex        =   37
      Top             =   1140
      Width           =   3435
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
      Left            =   17520
      TabIndex        =   36
      Top             =   1260
      Width           =   1515
   End
   Begin DashBoard.ucProgressCircular DespachosTurno1 
      Height          =   975
      Left            =   17700
      TabIndex        =   35
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
      Left            =   17460
      TabIndex        =   34
      Top             =   120
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
      Caption         =   "Form1.frx":3A824
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
      Left            =   19380
      TabIndex        =   33
      Top             =   1260
      Width           =   1515
   End
   Begin DashBoard.ucProgressCircular DespachosTurno2 
      Height          =   975
      Left            =   19560
      TabIndex        =   32
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
      Left            =   19320
      TabIndex        =   31
      Top             =   120
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
      Caption         =   "Form1.frx":3A858
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
      Left            =   19620
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
      Left            =   19620
      TabIndex        =   29
      Top             =   1320
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
      TabIndex        =   28
      Top             =   1260
      Width           =   1215
   End
   Begin DashBoard.ucProgressCircular CarbonTurno2 
      Height          =   975
      Left            =   10140
      TabIndex        =   27
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
      TabIndex        =   26
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
      Caption         =   "Form1.frx":3A88C
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
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
      TabIndex        =   25
      Top             =   1260
      Width           =   1215
   End
   Begin DashBoard.ucProgressCircular CarbonTurno1 
      Height          =   975
      Left            =   8640
      TabIndex        =   24
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
      TabIndex        =   23
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
      Caption         =   "Form1.frx":3A8C0
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
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
      TabIndex        =   22
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
      TabIndex        =   21
      Top             =   1260
      Width           =   1215
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
      Left            =   21180
      TabIndex        =   20
      Top             =   1260
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
      Left            =   22800
      TabIndex        =   19
      Top             =   1260
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
      Left            =   24420
      TabIndex        =   18
      Top             =   1260
      Width           =   1215
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
      TabIndex        =   17
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
      TabIndex        =   16
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
      TabIndex        =   15
      Top             =   1260
      Width           =   1215
   End
   Begin DashBoard.ucProgressCircular CarbonDia 
      Height          =   975
      Left            =   11700
      TabIndex        =   14
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
      TabIndex        =   12
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
      Caption         =   "Form1.frx":3A8F4
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
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
      Caption         =   "Form1.frx":3A928
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
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
      TabIndex        =   9
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
      Caption         =   "Form1.frx":3A95E
      CaptionPaddingY =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI Semibold"
         Size            =   8.25
         Charset         =   0
         Weight          =   600
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
      Left            =   6600
      Picture         =   "Form1.frx":3A996
      Top             =   720
      Width           =   1545
   End
   Begin DashBoard.ucProgressCircular DespachosAnio 
      Height          =   975
      Left            =   24300
      TabIndex        =   8
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
      Value           =   30000
      DisplayInPercent=   0   'False
      PF_ForeColor    =   133614
      AnimationInterval=   100
   End
   Begin DashBoard.LabelPlus LabelPlus6 
      Height          =   1755
      Left            =   24300
      TabIndex        =   7
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
      Caption         =   "Form1.frx":3E0B0
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
      Left            =   22680
      TabIndex        =   6
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
      Left            =   21180
      TabIndex        =   5
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
      Left            =   22740
      TabIndex        =   4
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
      Caption         =   "Form1.frx":3E0E8
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
      Left            =   21120
      TabIndex        =   3
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
      Caption         =   "Form1.frx":3E11E
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
      Caption         =   "DASHBOAR PRODUCCION LOGINEXT"
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
      Left            =   900
      TabIndex        =   2
      Top             =   240
      Width           =   6855
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
      TabIndex        =   1
      Top             =   720
      Width           =   4875
   End
   Begin DashBoard.LabelPlus LabelPlus1 
      Height          =   1815
      Left            =   180
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
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
      Caption         =   "Form1.frx":3E152
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PicturePaddingX =   -8
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
      PicturePresent  =   -1  'True
      PictureArr      =   "Form1.frx":3E172
   End
   Begin DashBoard.LabelPlus Acopio 
      Height          =   4875
      Index           =   3
      Left            =   4500
      TabIndex        =   92
      Top             =   7140
      Width           =   4035
      _ExtentX        =   7117
      _ExtentY        =   8599
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderColor     =   4210752
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":4237C
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      HotLineColor    =   65280
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
   Begin DashBoard.LabelPlus Acopio 
      Height          =   10755
      Index           =   0
      Left            =   180
      TabIndex        =   38
      Top             =   2100
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   18971
      BackShadow      =   0   'False
      Border          =   -1  'True
      BorderCornerLeftTop=   5
      BorderCornerRightTop=   5
      BorderCornerBottomRight=   5
      BorderCornerBottomLeft=   5
      BorderWidth     =   1
      CaptionAlignmentH=   1
      Caption         =   "Form1.frx":423C2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      PicturePaddingX =   5
      PicturePaddingY =   3
      ShadowSize      =   2
      ShadowOffsetX   =   2
      ShadowOffsetY   =   2
      PictureShadow   =   -1  'True
      HotLine         =   -1  'True
      HotLineColor    =   33023
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
    Set cPalette = NewCollection(&H4744E3, &H50C187, &HABA56C, &H48BDBF, &H4D91F4, &H7450, &H3DB0EF, vbBlue, vbGreen, vbRed, vbYellow, vbCyan)
        
    'Extrae el Nombre del Servidor
'    xFile = App.path & "\LOGICOAL.TXT"
'
'    Open xFile For Input As #1
'    Do While Not EOF(1)
'       Input #1, xA
'       Select Case xA
'       Case "[SERVER]"
'            Input #1, xSr
'       Case "[DATABASE]"
'            Input #1, xDb
'       Case "[MASTER]"
'            Input #1, xPs
'       End Select
'    Loop
'    Close #1

    Conn.Provider = "SQLOLEDB"
    'Conn.Properties("Integrated Security") = SSPI
    Conn.Properties("Data Source") = "CIFSRV0001"
    'Conn.Properties("Data Source") = "Localhost"
    Conn.Properties("Initial Catalog") = "Tracer"
    Conn.Properties("user ID") = "sisma_app"
    Conn.Properties("password") = "V$123bcd"
    Conn.CursorLocation = adUseServer
    Conn.CommandTimeout = 0
    Conn.Open

    MyNow = Now 'CDate(Format(#8/31/2022 5:30:00 PM#, "dd/MM/yyyy hh:mm"))
    Label3(1).Caption = Format(MyNow, "dd-MMMM-yyyy")
    
    DisplayIndex = 0
    Pasadas = 0
    Call Display
    
Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error abriendo Base de Datos," & vbCrLf & Err.Description
    MsgBox MSG, , "DASHBOARD DB"
    Err.Clear   ' Borra campos del objeto Err
    Exit Sub
End If
    
End Sub

Private Sub Display()
Dim xMin As Integer
Dim i As Long

    Call ProducionCarbonTurnos
    Call SaldosAcopiosRom
    Call ClasificadorasPS
    Call ProduccionT1
    Call ProduccionT2
    Call ProduccionT3
End Sub

Private Sub ProduccionT3()
    Dim Value As Collection
    Dim Value2 As Collection
    Dim colHoras As Collection
    Dim i As Integer, UltimoDia As Integer, j As Integer
    Dim xSql As String, FecTes As String
    Dim xDe As New ADODB.Recordset
    Dim xR As New ADODB.Recordset
    Dim xNeto As Single
    Dim Mes As Integer, Anio As Integer
    Dim FecIni As Date, FecFin As Date
    Dim AcumReal As Long, AcumPre
    
On Error GoTo Recover
   
    Tipo3.Clear
    
    Set Value = New Collection
    Set Value2 = New Collection
    UltimoDia = Conn.Execute("Select DAY(EOMONTH(GetDate())) AS DaysInMonth").Fields(0)
    
    For i = 1 To UltimoDia
        Value2.Add (i)
    Next i
    Tipo3.AddAxisItems Value2
    
    Mes = Month(MyNow)
    Anio = Year(MyNow)
    
    FecTes = "01/" & Format(Mes, "00") & "/" & Format(Anio, "0000") & " 06:00"
    
    FecIni = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(CDate(FecTes), "dd/MM/yyyy hh:mm") & "',3").Fields(0)
    FecFin = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(MyNow, "dd/MM/yyyy hh:mm") & "',4").Fields(0)
    
    xSql = "Set DateFormat DMY"
    xSql = xSql + " Select Dia,"
    xSql = xSql + "    (SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad"
    xSql = xSql + "    FROM Bascula INNER JOIN  Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote"
    xSql = xSql + "    WHERE (Bascula.TransaccionOrigen = 'LT') AND Lotes.IdTipoCarbon=3 AND (FechaTurno >='" & Format(FecIni, "dd/MM/yyyy hh:mm") & "' AND FechaTurno <='" & Format(FecFin, "dd/MM/yyyy hh:mm") & "') AND DAY(FechaTurno)=Dias.dia) AS Cantidad,"
    xSql = xSql + "    (SELECT ISNULL(SUM(Cantidad),0) AS Presupuesto FROM PlanesMinerosDetalle WHERE TipoCarbon=3 AND Mes=" & Mes & " AND anio=" & Anio & " AND Dia=Dias.Dia) AS Presupuesto"
    xSql = xSql + " From Dias Where Dia<=" & UltimoDia
        
    Set xR = Conn.Execute(xSql)
    
    For j = 1 To 4
        Select Case j
        Case 1
            Set Value = New Collection
            Do While Not xR.EOF
                AcumPre = AcumPre + xR.Fields(2).Value
                Value.Add AcumPre
                xR.MoveNext
            Loop
            Tipo3.AddLineSeries "Acum.Prto", Value, cPalette(j)
            xR.MoveFirst
        Case 2
            Set Value = New Collection
            Do While Not xR.EOF
                AcumReal = AcumReal + xR.Fields(1).Value
                Value.Add AcumReal
                xR.MoveNext
            Loop
            Tipo3.AddLineSeries "Acum.Real", Value, cPalette(j)
            xR.MoveFirst
        Case 3
            Set Value = New Collection
            Do While Not xR.EOF
                Value.Add xR.Fields(2).Value
                xR.MoveNext
            Loop
            Tipo3.AddLineSeries "Plan", Value, vbBlue
            xR.MoveFirst
        Case 4
            Set Value = New Collection
            Do While Not xR.EOF
                Value.Add xR.Fields(1).Value
                xR.MoveNext
            Loop
            Tipo3.AddLineSeries "Real", Value, vbYellow
            xR.MoveFirst
        
        End Select
    Next
   
    xR.Close
    
Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error Tipo3," & vbCrLf & Err.Description
    MsgBox MSG, , "ProduccionT3()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub

Private Sub ProduccionT2()
    Dim Value As Collection
    Dim Value2 As Collection
    Dim colHoras As Collection
    Dim i As Integer, UltimoDia As Integer, j As Integer
    Dim xSql As String, FecTes As String
    Dim xDe As New ADODB.Recordset
    Dim xR As New ADODB.Recordset
    Dim xNeto As Single
    Dim Mes As Integer, Anio As Integer
    Dim FecIni As Date, FecFin As Date
    Dim AcumReal As Long, AcumPre
    
On Error GoTo Recover
   
    Tipo2.Clear
    
    Set Value = New Collection
    Set Value2 = New Collection
    UltimoDia = Conn.Execute("Select DAY(EOMONTH(GetDate())) AS DaysInMonth").Fields(0)
    
    For i = 1 To UltimoDia
        Value2.Add (i)
    Next i
    Tipo2.AddAxisItems Value2
    
    Mes = Month(MyNow)
    Anio = Year(MyNow)
    
    FecTes = "01/" & Format(Mes, "00") & "/" & Format(Anio, "0000") & " 06:00"
    
    FecIni = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(CDate(FecTes), "dd/MM/yyyy hh:mm") & "',3").Fields(0)
    FecFin = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(MyNow, "dd/MM/yyyy hh:mm") & "',4").Fields(0)
    
    xSql = "Set DateFormat DMY"
    xSql = xSql + " Select Dia,"
    xSql = xSql + "    (SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad"
    xSql = xSql + "    FROM Bascula INNER JOIN  Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote"
    xSql = xSql + "    WHERE (Bascula.TransaccionOrigen = 'LT') AND Lotes.IdTipoCarbon=2 AND (FechaTurno >='" & Format(FecIni, "dd/MM/yyyy hh:mm") & "' AND FechaTurno <='" & Format(FecFin, "dd/MM/yyyy hh:mm") & "') AND DAY(FechaTurno)=Dias.dia) AS Cantidad,"
    xSql = xSql + "    (SELECT ISNULL(SUM(Cantidad),0) AS Presupuesto FROM PlanesMinerosDetalle WHERE TipoCarbon=2 AND Mes=" & Mes & " AND anio=" & Anio & " AND Dia=Dias.Dia) AS Presupuesto"
    xSql = xSql + " From Dias Where Dia<=" & UltimoDia
        
    Set xR = Conn.Execute(xSql)
    
    For j = 1 To 4
        Select Case j
        Case 1
            Set Value = New Collection
            Do While Not xR.EOF
                AcumPre = AcumPre + xR.Fields(2).Value
                Value.Add AcumPre
                xR.MoveNext
            Loop
            Tipo2.AddLineSeries "Acum.Prto", Value, cPalette(j)
            xR.MoveFirst
        Case 2
            Set Value = New Collection
            Do While Not xR.EOF
                AcumReal = AcumReal + xR.Fields(1).Value
                Value.Add AcumReal
                xR.MoveNext
            Loop
            Tipo2.AddLineSeries "Acum.Real", Value, cPalette(j)
            xR.MoveFirst
        Case 3
            Set Value = New Collection
            Do While Not xR.EOF
                Value.Add xR.Fields(2).Value
                xR.MoveNext
            Loop
            Tipo2.AddLineSeries "Plan", Value, vbBlue
            xR.MoveFirst
        Case 4
            Set Value = New Collection
            Do While Not xR.EOF
                Value.Add xR.Fields(1).Value
                xR.MoveNext
            Loop
            Tipo2.AddLineSeries "Real", Value, vbYellow
            xR.MoveFirst
        
        End Select
    Next
   
    xR.Close
    
Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error Tipo2," & vbCrLf & Err.Description
    MsgBox MSG, , "ProduccionT2()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub


Private Sub ProduccionT1()
Dim Value As Collection
Dim Value2 As Collection
Dim colHoras As Collection
Dim i As Integer, UltimoDia As Integer, j As Integer
Dim xSql As String, FecTes As String
Dim xDe As New ADODB.Recordset
Dim xR As New ADODB.Recordset
Dim xNeto As Single
Dim Mes As Integer, Anio As Integer
Dim FecIni As Date, FecFin As Date
Dim AcumReal As Long, AcumPre
    
On Error GoTo Recover
   
    Tipo1.Clear
    
    Set Value = New Collection
    Set Value2 = New Collection
    UltimoDia = Conn.Execute("Select DAY(EOMONTH(GetDate())) AS DaysInMonth").Fields(0)
    
    For i = 1 To UltimoDia
        Value2.Add (i)
    Next i
    Tipo1.AddAxisItems Value2
    
    Mes = Month(MyNow)
    Anio = Year(MyNow)
    
    FecTes = "01/" & Format(Mes, "00") & "/" & Format(Anio, "0000") & " 06:00"
    
    FecIni = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(CDate(FecTes), "dd/MM/yyyy hh:mm") & "',3").Fields(0)
    FecFin = Conn.Execute("SET DateFormat DMY EXEC PA_FechaTurno '" & Format(MyNow, "dd/MM/yyyy hh:mm") & "',4").Fields(0)
    
    xSql = "Set DateFormat DMY"
    xSql = xSql + " Select Dia,"
    xSql = xSql + "    (SELECT  Isnull(SUM(Bascula.PesoLleno - Bascula.PesoVacio),0)/1000 AS Cantidad"
    xSql = xSql + "    FROM Bascula INNER JOIN  Lotes ON Bascula.NumeroTransaccion = Lotes.IdLote"
    xSql = xSql + "    WHERE (Bascula.TransaccionOrigen = 'LT') AND Lotes.IdTipoCarbon=1 AND (FechaTurno >='" & Format(FecIni, "dd/MM/yyyy hh:mm") & "' AND FechaTurno <='" & Format(FecFin, "dd/MM/yyyy hh:mm") & "') AND DAY(FechaTurno)=Dias.dia) AS Cantidad,"
    xSql = xSql + "    (SELECT ISNULL(SUM(Cantidad),0) AS Presupuesto FROM PlanesMinerosDetalle WHERE TipoCarbon=1 AND Mes=" & Mes & " AND anio=" & Anio & " AND Dia=Dias.Dia) AS Presupuesto"
    xSql = xSql + " From Dias Where Dia<=" & UltimoDia
        
    Set xR = Conn.Execute(xSql)
    
    For j = 1 To 4
        Select Case j
        Case 1
            Set Value = New Collection
            Do While Not xR.EOF
                AcumPre = AcumPre + xR.Fields(2).Value
                Value.Add AcumPre
                xR.MoveNext
            Loop
            Tipo1.AddLineSeries "Acum.Prto", Value, cPalette(j)
            xR.MoveFirst
        Case 2
            Set Value = New Collection
            Do While Not xR.EOF
                AcumReal = AcumReal + xR.Fields(1).Value
                Value.Add AcumReal
                xR.MoveNext
            Loop
            Tipo1.AddLineSeries "Acum.Real", Value, cPalette(j)
            xR.MoveFirst
        Case 3
            Set Value = New Collection
            Do While Not xR.EOF
                Value.Add xR.Fields(2).Value
                xR.MoveNext
            Loop
            Tipo1.AddLineSeries "Plan", Value, vbBlue
            xR.MoveFirst
        Case 4
            Set Value = New Collection
            Do While Not xR.EOF
                Value.Add xR.Fields(1).Value
                xR.MoveNext
            Loop
            Tipo1.AddLineSeries "Real", Value, vbYellow
            xR.MoveFirst
        
        End Select
    Next
   
    xR.Close
    
Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error T1," & vbCrLf & Err.Description
    MsgBox MSG, , "ProduccionT1()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub

Private Sub ClasificadorasPS()
Dim xSql As String
Dim xA As New ADODB.Recordset
Dim xR As New ADODB.Recordset
Dim i As Integer, j As Integer, NoGrafico As Integer
Dim FecIni As Date, FecFin As Date

' On Error GoTo Recover

FecIni = CDate("01" & "/" & Trim(Str(Month(MyNow))) & "/" & Trim(Str(Year(MyNow))))
FecFin = MyNow
i = 0
j = 0       ' Numero de Pasadas
NoGrafico = 2

xSql = "SET DATEFORMAT DMY"
xSql = xSql + " SELECT   Trituradoras.Descripcion, SUM(Trituracion.Cantidad) Cantidad, SUM(Trituracion.HorasEfectivas) HE, SUM(DATEDIFF(MINUTE, TrituradoraParos.FechaInicio, TrituradoraParos.FechaFin))/60 AS HorasParos, "
xSql = xSql + " DATEDIFF(MINUTE,'" & Format(FecIni, "dd/MM/yyyy hh:mm") & "','" & Format(FecFin, "dd/MM/yyyy hh:mm") & "')/60 AS HorasTurno"
xSql = xSql + " FROM     TrituradoraParos INNER JOIN"
xSql = xSql + "          Trituracion ON TrituradoraParos.IdTrituracion = Trituracion.IdTrituracion INNER JOIN"
xSql = xSql + "          Trituradoras ON Trituracion.IdTrituradora = Trituradoras.IdTrituradora"
xSql = xSql + " WHERE    trituracion.FechaInicio>='" & Format(FecIni, "dd/MM/yyyy hh:mm") & "' AND trituracion.FechaCierre<='" & Format(FecFin, "dd/MM/yyyy hh:mm") & "'  AND Trituradoras.IdTrituradora<=2"
xSql = xSql + " GROUP BY Trituradoras.Descripcion"
xSql = xSql + " ORDER BY Trituradoras.Descripcion"

Set xR = Conn.Execute(xSql)
Do While Not xR.EOF
    j = j + 1
    If j > NoGrafico Then
        MsgBox "El Indice " & Trim(Str(j)) & " Sobrepasa la capacidad Gráfica " & Trim(Str(NoGrafico))
        Exit Sub
    End If
    LabelPS(i) = "Horas Turno"
    DatoPS(i) = Format(xR!HorasTurno, "##0")
    LabelPS(i + 1) = "Horas Paro"
    DatoPS(i + 1) = Format(xR!HorasParos, "##0.0")
    LabelPS(i + 2) = "Horas Efectivas"
    DatoPS(i + 2) = Format(xR!HE, "##0.0")
    LabelPS(i + 3) = "Procesado Ton."
    DatoPS(i + 3) = Format(xR!Cantidad, "###,##0")
    xR.MoveNext
    i = i + 4
Loop
xR.Close

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error," & vbCrLf & Err.Description
    MsgBox MSG, , "ClasificadorasPS()"
    Err.Clear   ' Borra campos del objeto Err
    Exit Sub
End If
End Sub

Private Sub SaldosAcopiosRom()
Dim xSql As String
Dim xA As New ADODB.Recordset
Dim xR As New ADODB.Recordset
Dim i As Integer, j As Integer, CountAcopio As Integer, NoGrafico As Integer

On Error GoTo Recover

i = 0
NoGrafico = 18
j = 0
CountAcopio = 0

xSql = "SELECT  DISTINCT  Acopios.Descripcion, Acopios.IdAcopio"
xSql = xSql + " FROM      Pilas INNER JOIN"
xSql = xSql + "           PilasFisicas ON Pilas.IdPilaFisica = PilasFisicas.IdPilaFisica INNER JOIN"
xSql = xSql + "           Acopios ON Pilas.IdAcopio = Acopios.IdAcopio"
xSql = xSql + " WHERE     pilas.Estado='IN'"
xSql = xSql + " ORDER BY  IdAcopio"
Set xA = Conn.Execute(xSql)

Do While Not xA.EOF
    Acopio(CountAcopio).Caption = xA!Descripcion
    xSql = "SELECT  Acopios.Descripcion, PilasFisicas.Descripcion AS DesPila, Pilas.Cantidad"
    xSql = xSql + " FROM  Pilas INNER JOIN"
    xSql = xSql + "       PilasFisicas ON Pilas.IdPilaFisica = PilasFisicas.IdPilaFisica INNER JOIN"
    xSql = xSql + "       Acopios ON Pilas.IdAcopio = Acopios.IdAcopio"
    xSql = xSql + " WHERE pilas.Estado='IN' AND Acopios.IdAcopio=" & xA!IdAcopio
    CountAcopio = CountAcopio + 1
    
    Set xR = Conn.Execute(xSql)
    Do While Not xR.EOF
        j = j + 1
        If j > NoGrafico Then
            MsgBox "El Indice " & Trim(Str(j)) & " Sobrepasa la capacidad Gráfica " & Trim(Str(NoGrafico))
            Exit Sub
        End If
        
        Pila(i).Caption = "     " & xR!DesPila
        SaldoPila(i).Caption = "Saldo " & Format(xR!Cantidad, "#,###,##0") & " Ton."
        i = i + 1
        xR.MoveNext
    Loop
    xR.Close
    xA.MoveNext
Loop

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error," & vbCrLf & Err.Description
    MsgBox MSG, , "SaldosAcopiosRom()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
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
    Exit Sub
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
MyNow = Now ' CDate(Format(#8/31/2022 5:30:00 PM#, "dd/MM/yyyy hh:mm"))

Pasadas = Pasadas + 1
LabelPlus1.Caption = Format(Pasadas, "###,###")

Select Case DisplayIndex
Case 0
    Call ProducionCarbonTurnos
Case 1
    Call SaldosAcopiosRom
Case 2
    Call ClasificadorasPS
Case 3
 
Case 4

Case 5

Case 6

End Select

DisplayIndex = DisplayIndex + 1
If DisplayIndex > 6 Then DisplayIndex = 0

End Sub



