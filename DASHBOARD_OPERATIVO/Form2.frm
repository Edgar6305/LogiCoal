VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00E5E5E5&
   Caption         =   "Form2"
   ClientHeight    =   10170
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   16560
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   10170
   ScaleWidth      =   16560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Dona"
      Height          =   615
      Left            =   1500
      TabIndex        =   20
      Top             =   8220
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   435
      Left            =   5400
      TabIndex        =   18
      Top             =   7800
      Width           =   555
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   1500
      TabIndex        =   16
      Top             =   7380
      Width           =   1695
   End
   Begin Proyecto1.ucPieChart PiePala 
      Height          =   3255
      Left            =   12360
      TabIndex        =   19
      Top             =   6540
      Width           =   3435
      _ExtentX        =   6059
      _ExtentY        =   5741
      Title           =   "Producci�n Palas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelsVisible   =   -1  'True
      ChartStyle      =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
      LabelsPositions =   2
   End
   Begin Proyecto1.ucChartBar ChartBar3 
      Height          =   2355
      Left            =   6120
      TabIndex        =   17
      Top             =   6480
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   4154
      Title           =   "Camiones Hora"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Border          =   -1  'True
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   16.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
      LabelsPositions =   0
   End
   Begin Proyecto1.ucProgressCircular Combustible 
      Height          =   1275
      Left            =   3900
      TabIndex        =   15
      Top             =   6840
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   2249
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -7
      Caption2        =   "Galones"
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   10
      Value           =   1
      DisplayInPercent=   0   'False
      AnimationInterval=   100
   End
   Begin Proyecto1.LabelPlus LabelPlus5 
      Height          =   315
      Left            =   1380
      TabIndex        =   14
      Top             =   6660
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   556
      BackColor       =   16777215
      BackShadow      =   0   'False
      Caption         =   "Form2.frx":0000
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   4210752
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin Proyecto1.LabelPlus LabelPlus4 
      Height          =   3675
      Left            =   120
      TabIndex        =   13
      Top             =   6240
      Width           =   15705
      _ExtentX        =   27702
      _ExtentY        =   6482
      BackColor       =   16777215
      Border          =   -1  'True
      BorderColor     =   0
      CaptionAlignmentH=   1
      Caption         =   "Form2.frx":002A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PictureAlignmentV=   1
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PicturePresent  =   -1  'True
      PictureArr      =   "Form2.frx":0064
   End
   Begin Proyecto1.LabelPlus LabelPlus3 
      Height          =   1215
      Left            =   10080
      TabIndex        =   12
      Top             =   2520
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   2143
      BackColorOpacity=   0
      BackShadow      =   0   'False
      CaptionAlignmentH=   1
      CaptionAlignmentV=   1
      Caption         =   "Form2.frx":1F02
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin Proyecto1.LabelPlus LabelPlus2 
      Height          =   1695
      Left            =   9960
      TabIndex        =   11
      Top             =   2280
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   2990
      BackColor       =   16777215
      BackShadow      =   0   'False
      Caption         =   "Form2.frx":1F38
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
   Begin Proyecto1.ucPieChart ucPieChart2 
      Height          =   2055
      Left            =   9960
      TabIndex        =   10
      Top             =   4080
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LabelsVisible   =   -1  'True
      ChartStyle      =   1
      LegendVisible   =   0   'False
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
      DonutWidth      =   10
      LabelsPositions =   2
   End
   Begin Proyecto1.ucChartBar ucChartBar2 
      Height          =   2175
      Left            =   120
      TabIndex        =   9
      Top             =   3960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3836
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ChartStyle      =   1
      ChartOrientation=   1
      LegendVisible   =   0   'False
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
   End
   Begin Proyecto1.ucChartArea ucChartArea2 
      Height          =   1935
      Left            =   4680
      TabIndex        =   8
      Top             =   4200
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3413
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LinesCurve      =   -1  'True
      LegendAlign     =   1
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
      BorderColor     =   15132391
   End
   Begin Proyecto1.ucTreeMaps ucTreeMaps1 
      Height          =   2055
      Left            =   4680
      TabIndex        =   7
      Top             =   2040
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   9359529
      LabelsVisible   =   0   'False
      LegendVisible   =   0   'False
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Proyecto1.ucChartArea ucChartArea1 
      Height          =   1935
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3413
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FillOpacity     =   0
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
      BorderColor     =   15132391
   End
   Begin Proyecto1.ucChartBar ucChartBar1 
      Height          =   1815
      Left            =   4680
      TabIndex        =   5
      Top             =   120
      Width           =   5175
      _ExtentX        =   9128
      _ExtentY        =   3201
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LegendVisible   =   0   'False
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
   End
   Begin Proyecto1.ucPieChart ucPieChart1 
      Height          =   2055
      Left            =   9960
      TabIndex        =   4
      Top             =   120
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   3625
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      LegendVisible   =   0   'False
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TitleForeColor  =   0
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1335
      Index           =   1
      Left            =   1680
      TabIndex        =   3
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2355
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -7
      Caption2        =   "Visitors"
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   10
      PB_Color1       =   13553360
      Value           =   30
      PF_ForeColor    =   11892015
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1335
      Index           =   0
      Left            =   240
      TabIndex        =   2
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2355
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -7
      Caption2        =   "Band"
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   10
      PB_Color1       =   13553360
      PB_Border       =   -1  'True
      PF_ForeColor    =   9359529
      AnimationInterval=   100
   End
   Begin Proyecto1.ucProgressCircular ucProgressCircular1 
      Height          =   1335
      Index           =   2
      Left            =   3120
      TabIndex        =   1
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   2355
      BeginProperty Caption1_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption1_OffsetY=   -7
      Caption2        =   "Volume"
      BeginProperty Caption2_Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption2_OffsetY=   10
      PB_Color1       =   13553360
      Value           =   70
      PF_ForeColor    =   5460735
      AnimationInterval=   100
   End
   Begin Proyecto1.LabelPlus LabelPlus1 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   2990
      BackColor       =   16777215
      BackShadow      =   0   'False
      Caption         =   "Form2.frx":1F70
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShadowColorOpacity=   0
      CallOutAlign    =   0
      CallOutWidth    =   0
      CallOutLen      =   0
      MousePointer    =   0
      BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IconForeColor   =   0
      IconOpacity     =   0
      PictureArr      =   0
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cResizer As ClsResizer

Private Sub Command1_Click()
Combustible.Value = Combustible.Value + 10
If Combustible.Value >= 100 Then Combustible.Value = Combustible.Value - 100
End Sub

Private Sub Command2_Click()
    Dim Value As Collection
    Dim Value2 As Collection
    Dim colDate As Collection
    Dim i As Integer
    
    ChartBar3.Clear
    
    Set Value = New Collection
    Set Value2 = New Collection
    Set colDate = New Collection
    For i = 1 To 12
        Value.Add Random(10, 60)
        colDate.Add Format(i + 5, "00")
    Next
    For i = 1 To 12
        Value2.Add Random(1, 20)
    Next
    
    ChartBar3.AddAxisItems colDate, , 305, 2
    ChartBar3.AddSerie "Camiones", vbRed, Value
    ChartBar3.AddSerie "Carbon", vbBlue, Value2
        
End Sub

Private Sub Command3_Click()
Dim i As Integer
Dim cPalette As Collection


Set cPalette = NewCollection(&H4744E3, &H50C187, &HABA56C, &H48BDBF, &H4D91F4, &H7450, &H3DB0EF)

PiePala.Clear
For i = 1 To 5
    PiePala.AddItem "Pala " & i, Random(10, 100), cPalette(i)
Next

End Sub

Private Sub Form_Load()
    Dim cPalette As Collection
    Dim i As Long, j As Long
    Dim Value As Collection
    Dim Lables As Collection
    Dim Icons As Collection
    Dim keys As Collection
    Dim CustomColors As Collection
    Dim colDate As Collection
    
    
    Set cResizer = New ClsResizer
    With cResizer
        .AddControlFont "ucProgressCircular", "Caption1_Font", "Caption2_Font"
    
        .AddControlFont "ucChartArea", "Font", "TitleFont"
        .AddControlFont "ucChartBar", "Font", "TitleFont"
        .AddControlFont "ucPieChart", "Font", "TitleFont"
        .AddControlFont "ucTreeMaps", "Font", "TitleFont"
        .AddControlFont "LabelPlus", "Font"
    
        .AddControlProperty "ucProgressCircular", "PB_Width", "PF_Width", "Caption1_OffsetY", "Caption2_OffsetY"
        .AddControlProperty "LabelPlus", "HotLineWidth"
        .AddControlProperty "ucChartArea", "LinesWidth"
        .AddControlProperty "ucPieChart", "SeparatorLineWidth", "DonutWidth"
        
        .SaveControlsPositions Me
    End With
    
    Set cPalette = NewCollection(&H4744E3, &H50C187, &HABA56C, &H48BDBF, &H4D91F4, &H7450, &H3DB0EF)

    
    Randomize Timer
    For i = 0 To 2
        ucPieChart1.AddItem "2000" + i, Random(10, 30), CLng(cPalette(i + 1))
    Next
    
    Set Value = New Collection
    For i = 0 To 6
        Value.Add Random(10, 10 * i)
    Next

    Set Lables = New Collection
    With Lables
        .Add "Facebook"
        .Add "Intagram"
        .Add "Wikipedia"
        .Add "Pinterest"
        .Add "WhatsApp"
        .Add "Twiter"
        .Add "Youtube"
    End With
    ucTreeMaps1.AddLineSeries vbNullString, vbBlue, Value, Lables
    
    Set Value = New Collection
    Set colDate = New Collection
    For i = 1 To 12
        Value.Add Random(10, 300)
        colDate.Add Format(DateSerial(2020, i, 1), "mmm")
    Next
    
    ucChartBar1.AddAxisItems colDate, , 305, 2
    ucChartBar1.AddSerie "ASDF", vbRed, Value
    
    Set Value = New Collection
    For i = 1 To cPalette.Count
        Value.Add Random(0, 100 * i)
    Next
    
    ucChartBar2.AddAxisItems Lables, True, , 2
    ucChartBar2.AddSerie "Serie 1", vbRed, Value, cPalette
    
    For j = 1 To 3
        Set Value = New Collection
        For i = 1 To cPalette.Count
            Value.Add Random(10 * i, 100 * i)
        Next
        ucChartArea1.AddLineSeries "Serie " & j, Value, cPalette(j)
    Next
    
    
    For j = 1 To 2
        Set Value = New Collection
        For i = 1 To cPalette.Count
            Value.Add Random(10 * i, 100 * i)
        Next
        ucChartArea2.AddLineSeries "Serie " & j, Value, cPalette(j + 5)
    Next
    
    For i = 1 To 5
        ucPieChart2.AddItem "Serie " & i, Random(10, 100), cPalette(i)
    Next

End Sub

Private Function NewCollection(ParamArray vArgList() As Variant) As Collection
    Dim Value As Variant
    Set NewCollection = New Collection
    For Each Value In vArgList
        NewCollection.Add Value
    Next
End Function

Private Function Random(ByVal Min!, ByVal Max!)
    Random = Int((Max - Min + 1) * Rnd + Min)
End Function

Private Sub Form_Resize()
    cResizer.ResizeControls Me
End Sub

