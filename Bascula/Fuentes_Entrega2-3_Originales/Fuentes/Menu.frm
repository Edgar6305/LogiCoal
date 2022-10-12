VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Menu 
   Caption         =   "Menu Tracer"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   18195
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   18195
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Image1 
      BackColor       =   &H00FFFFFF&
      Height          =   1275
      Left            =   0
      Picture         =   "Menu.frx":112A
      ScaleHeight     =   1215
      ScaleWidth      =   19935
      TabIndex        =   2
      Top             =   0
      Width           =   19995
      Begin VB.PictureBox Picture4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   555
         Left            =   12720
         Picture         =   "Menu.frx":61C0
         ScaleHeight     =   555
         ScaleWidth      =   2895
         TabIndex        =   5
         Top             =   480
         Width           =   2895
      End
      Begin VB.PictureBox Menu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   15840
         Picture         =   "Menu.frx":7581
         ScaleHeight     =   615
         ScaleWidth      =   615
         TabIndex        =   3
         Tag             =   "0"
         Top             =   480
         Width           =   615
      End
   End
   Begin MBSplit.Splitter VSplitter 
      Height          =   7875
      Left            =   20
      TabIndex        =   0
      Top             =   1320
      Width           =   20025
      _ExtentX        =   35322
      _ExtentY        =   13891
      SplitterPos     =   0
      Begin ComctlLib.TreeView oDest 
         Height          =   7215
         Left            =   150
         TabIndex        =   4
         Top             =   60
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   12726
         _Version        =   327682
         Indentation     =   882
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
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7155
         Left            =   4020
         ScaleHeight     =   7095
         ScaleWidth      =   15915
         TabIndex        =   1
         Top             =   120
         Width           =   15975
         Begin VB.CommandButton Command1 
            Caption         =   "Command1"
            Height          =   615
            Left            =   10800
            TabIndex        =   7
            Top             =   4140
            Visible         =   0   'False
            Width           =   1935
         End
         Begin VB.Timer Timer1 
            Interval        =   60000
            Left            =   10560
            Top             =   1620
         End
         Begin Crystal.CrystalReport oCr 
            Left            =   10560
            Top             =   2160
            _ExtentX        =   741
            _ExtentY        =   741
            _Version        =   348160
            WindowControlBox=   -1  'True
            WindowMaxButton =   -1  'True
            WindowMinButton =   -1  'True
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
         Begin Tracer.LabelPlus oFavorito 
            Height          =   1215
            Index           =   0
            Left            =   360
            TabIndex        =   6
            Top             =   360
            Visible         =   0   'False
            Width           =   2955
            _ExtentX        =   5212
            _ExtentY        =   2143
            BackColor       =   16777215
            BackShadow      =   0   'False
            Border          =   -1  'True
            BorderColor     =   8421504
            BorderCornerLeftTop=   2
            BorderCornerRightTop=   2
            BorderCornerBottomRight=   2
            BorderCornerBottomLeft=   2
            BorderWidth     =   1
            CaptionAlignmentH=   1
            Caption         =   "Menu.frx":7A80
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PicturePaddingX =   5
            PicturePaddingY =   25
            ShadowSize      =   3
            ShadowOffsetX   =   2
            ShadowOffsetY   =   2
            HotLine         =   -1  'True
            HotLineColor    =   12648384
            HotLineWidth    =   20
            HotLinePosition =   1
            BeginProperty IconFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PicturePresent  =   -1  'True
            PictureArr      =   "Menu.frx":7AA0
         End
         Begin ComctlLib.ImageList ImageList1 
            Left            =   10440
            Top             =   2760
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            UseMaskColor    =   0   'False
            _Version        =   327682
            BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
               NumListImages   =   62
               BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":8676
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":89C8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":964A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":999C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":9CEE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":9EC8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":A21A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":A56C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":A8BE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":AC10
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":AF62
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":B2B4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":B606
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":B958
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":BCAA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":BFFC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":C34E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":CD60
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":D0B2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":D28C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage21 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":DC0E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage22 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":DF60
                  Key             =   ""
               EndProperty
               BeginProperty ListImage23 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":E27A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage24 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":EDC4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage25 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":F836
                  Key             =   ""
               EndProperty
               BeginProperty ListImage26 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":102A8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage27 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":10C42
                  Key             =   ""
               EndProperty
               BeginProperty ListImage28 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":116B4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage29 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":11F8E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage30 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":12A60
                  Key             =   ""
               EndProperty
               BeginProperty ListImage31 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":133E2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage32 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":13EA0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage33 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":14DE6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage34 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":15608
                  Key             =   ""
               EndProperty
               BeginProperty ListImage35 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":15CE2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage36 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":16394
                  Key             =   ""
               EndProperty
               BeginProperty ListImage37 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":1656E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage38 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":1D0D4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage39 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":31D2E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage40 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":31F08
                  Key             =   ""
               EndProperty
               BeginProperty ListImage41 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":3225A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage42 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":39028
                  Key             =   ""
               EndProperty
               BeginProperty ListImage43 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":3989A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage44 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":3A21C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage45 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":451EE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage46 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":473F0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage47 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":48042
                  Key             =   ""
               EndProperty
               BeginProperty ListImage48 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":49F0C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage49 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":4BBBA
                  Key             =   ""
               EndProperty
               BeginProperty ListImage50 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":4C4AC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage51 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":4D602
                  Key             =   ""
               EndProperty
               BeginProperty ListImage52 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":4E6D4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage53 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":4FBDE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage54 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":50CF4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage55 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":51F32
                  Key             =   ""
               EndProperty
               BeginProperty ListImage56 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":53218
                  Key             =   ""
               EndProperty
               BeginProperty ListImage57 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":53B0A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage58 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":545FC
                  Key             =   ""
               EndProperty
               BeginProperty ListImage59 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":5495A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage60 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":55984
                  Key             =   ""
               EndProperty
               BeginProperty ListImage61 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":562C6
                  Key             =   ""
               EndProperty
               BeginProperty ListImage62 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                  Picture         =   "Menu.frx":56AF8
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim mFont2 As StdFont
Dim cPalette As Collection
Dim Flag_Menu As Boolean
Const COLOR_HOT As Long = &HFFF3E5
Const COLOR_SELECTED As Long = &HFFE8CC
Dim ok As Boolean

Private Sub Command1_Click()
'Movimientos.Show
Notas.Show
End Sub

Private Sub Form_Activate()
If Not ok Then
    ok = True
End If
End Sub

Private Sub Form_Load()
    Dim i As Integer
    
    'set the child controls for the horizontal splitter
    Set VSplitter.LeftOrTopCtl = oDest
    Set VSplitter.RightOrBottomCtl = Picture2
    
    Set mFont2 = New StdFont
    mFont2.Name = "Segoe UI"
    mFont2.Size = 7
    
    Set cPalette = NewCollection(vbBlue, vbGreen, vbRed, &H4744E3, &H50C187, &HABA56C, &H48BDBF, &H4D91F4, &H7450, vbYellow, &H3DB0EF, vbCyan)
        
    Set PassWord.dControl = Menu
    PassWord.Show 1
    
    Me.Caption = "Usuario " & Susuario
    Me.Caption = Me.Caption & "      Version:  " & App.Major & "." & App.Minor & "." & App.Revision
    Timer1.Enabled = True
    If Menu.Tag = 1 Then
        Call ArmaArbol
    End If
End Sub

Private Sub Form_Resize()
    Dim lSplitHeight As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    Image1.Width = Me.Width
    lSplitHeight = Me.ScaleHeight - Image1.Height - 100
    VSplitter.Width = Me.ScaleWidth - 100
    VSplitter.Height = IIf(lSplitHeight < 0, 0, lSplitHeight)
    oDest.Height = Me.Height - Picture4.Height
    Menu.Left = Me.ScaleWidth - Menu.Width - 1100
    Picture4.Left = Me.ScaleWidth - (Menu.Width + 1100 + 3000)
    VSplitter.SplitterPos = 0

End Sub

Private Sub Form_Unload(Cancel As Integer)
Conn.Close
   Dim v
   For Each v In Forms
      Unload v
   Next
End Sub

Private Sub LabelPlus12_MouseOver()
    LabelPlus12.BackColor = vbRed
End Sub

Private Sub Menu_Click()
Flag_Menu = Not Flag_Menu
If Flag_Menu Then
    For i = 1 To 35 Step 5
        VSplitter.SplitterPos = i
    Next i
Else
    For i = 35 To 0 Step -5
        VSplitter.SplitterPos = i
    Next i
End If
End Sub

Private Sub ArmaArbol()
Dim i As Integer
Dim mIcono As Integer
Dim xDe As New ADODB.Recordset
Dim xR As New ADODB.Recordset
Dim xSql As String

On Error GoTo Recover

Set Nodx = oDest.Nodes.Add(, , "*" & "0", "Trazabilidad del Carbon", 53)
xSql = "Select * from FT_ArbolMenu('" & Susuario & "') WHERE Acceso=1 Order By Consecutivo"
Set xDe = Conn.Execute(xSql)
While Not xDe.EOF
     mIcono = xDe!iCono
     If xDe!VBPrograma = "" Then
        Set Nodx = oDest.Nodes.Add("*" & xDe!MenuDependiente, tvwChild, "*" & xDe!Consecutivo, xDe!CodigoMenu, mIcono)
     Else
        Set Nodx = oDest.Nodes.Add("*" & xDe!MenuDependiente, tvwChild, "*" & xDe!Consecutivo, xDe!CodigoMenu, mIcono)
     End If
     Nodx.Tag = xDe!VBPrograma
     xDe.MoveNext
Wend
xDe.Close

sUsuarioFavoritos = True
SmaxDataFavoritos = 10
SmaxViewFavoritos = 5
If sUsuarioFavoritos Then
    xSql = "Select TOP " & SmaxDataFavoritos & " SUM(Numero) AS Suma,Programa,Descripcion from favoritos where usuario= '" & Susuario & "' group by Programa,Descripcion ORDER BY Suma DESC"
    Set xR = Conn.Execute(xSql)
    i = 0
    j = 1

    Do While Not xR.EOF
       If i > 0 Then
            Load oFavorito(i)
            If i >= SmaxViewFavoritos Then
                If i = SmaxViewFavoritos Then
                    oFavorito(i).Top = oFavorito(i - SmaxViewFavoritos).Top
                Else
                    oFavorito(i).Top = oFavorito(i - SmaxViewFavoritos - 1).Top + (oFavorito(i - SmaxViewFavoritos - 1).Height + 150)
                End If

                oFavorito(i).Left = oFavorito(i - SmaxViewFavoritos).Left + (oFavorito(i - SmaxViewFavoritos).Width + 300)
            Else
                oFavorito(i).Top = oFavorito(i - 1).Top + (oFavorito(i).Height + 150)
            End If
       End If

       oFavorito(i).Visible = True
       j = Random(i + 1, 12)
       oFavorito(i).HotLineColor = cPalette(i + 1)
       oFavorito(i).Caption = xR!Descripcion
       oFavorito(i).Tag = xR!Programa
       
       i = i + 1
       xR.MoveNext
    Loop
    xR.Close
End If

Exit Sub

Recover:
If Err.Number <> 0 Then
    MSG = "Llave Anterior" & xDe!MenuDependiente & " Llave Posterior " & xDe!Consecutivo & vbCrLf & Err.Description
    MsgBox MSG, , "Muestra Arbol"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub

Private Sub oDest_Click()
Dim frm As Form

On Error GoTo Recover

If oDest.SelectedItem.Child Is Nothing Then
    sProgSISMA = oDest.SelectedItem.Tag
    If sProgSISMA <> Flag_xProg Then
        Flag_xProg = sProgSISMA
        If sProgSISMA <> "" Then
            VSplitter.SplitterPos = 0
            Set frm = Forms.Add(sProgSISMA)
            frm.Show 1
            VSplitter.SplitterPos = 35
            Conn.Execute ("Set DateFormat DMY INSERT INTO Favoritos VALUES('" & Susuario & "','" & sProgSISMA & "','" & oDest.SelectedItem.text & "',1,'" & Format(Now(), "dd/MM/yyyy hh:mm") & "')")
        End If
    End If
    Set frm = Nothing
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error Cargando Formualario," & vbCrLf & Err.Description
    MsgBox MSG, , "oDest_Click()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub

Private Sub oFavorito_Click(Index As Integer)
Dim frm As Form
On Error GoTo Recover

sProgSISMA = oFavorito(Index).Tag
Set frm = Forms.Add(oFavorito(Index).Tag)
frm.Show
Conn.Execute ("Set DateFormat DMY INSERT INTO Favoritos VALUES('" & Susuario & "','" & oFavorito(Index).Tag & "','" & oFavorito(Index).Caption & "',1,'" & Format(Now(), "dd/MM/yyyy hh:mm") & "')")

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error Cargando Formualario," & vbCrLf & Err.Description
    MsgBox MSG, , "oFavorito_Click"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub

Private Sub oFavorito_GotFocus(Index As Integer)
    'oFavorito(Index).BackColor = COLOR_HOT
End Sub

Private Sub oFavorito_MouseEnter(Index As Integer)
    oFavorito(Index).BackColor = COLOR_HOT
End Sub

Private Sub oFavorito_MouseLeave(Index As Integer)
    oFavorito(Index).BackColor = &HFFFFFF
End Sub

Private Sub oFavorito_PostPaint(Index As Integer, ByVal HDC As Long)
    Dim mTop As Long, TextHeight As Long
    Dim sTitle As String
    Dim bProtected As Boolean
    Dim sDescription As String
    Dim lWidth As Long
    Dim lMargin As Long

    With oFavorito(Index)
        
        mTop = 25   '100 - .BackColorOpacity / 1.5
        sDescription = Conn.Execute("Select Descripcion From Programas Where Programa='" & oFavorito(Index).Tag & "'").Fields(0)
         
        lMargin = 35 '* .GetWindowsDPI0
        lWidth = ((.Width / .GetWindowsDPI / Screen.TwipsPerPixelX) - lMargin * 2)    '100= aproximate height
                                                                 
        TextHeight = .DrawText(HDC, sTitle, lMargin, mTop, lWidth, 200, mFont2, vbWhite, 100, ccEnter, cTop, True)
         .DrawText HDC, sDescription, lMargin, mTop, lWidth, 200, mFont2, &H4F3E38, 100, ccEnter, cTop, True
    End With

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

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
                dControl.SelStart = 0
                dControl.SelLength = Len(dControl.text)
End If
End Sub

Private Sub Timer1_Timer()
Dim JSONSend    As Dictionary
Dim JSONResult As Dictionary
Dim JSONResultREV As Dictionary
Dim winH As WinHttp.WinHttpRequest
Dim xP As New ADODB.Recordset
Dim i As Integer

On Error GoTo Recover

'objWinHTTP.SetTimeouts 5000, 5000, 5000, 5000
'Resolve, Connect, Send and Receive
'If WinHttpReq.WaitForResponse(TimeOutSec) = 0 Then
'        errString = "ReqErr: TimeOut"
'        bolResult = False
'
'        Exit Function
'
'End If
    
SLastTime = SLastTime + 1
If SLastTime > SInterval Then
    Set JSONSend = New Dictionary
    Set winH = New WinHttp.WinHttpRequest

    JSONSend.Item("empresa") = Susuario

    winH.Open "post", "http://18.229.172.128:3000/api/v1/avisos", False
    winH.SetRequestHeader "Content-type", "application/json"
    winH.SetRequestHeader "charset", "UTF-8"
    winH.SetRequestHeader "token", sTokenSisma
    winH.SetTimeouts 500, 500, 1000, 1000

    winH.Send JsonConverter.ConvertToJson(JSONSend)
    Set JSONResult = JsonConverter.ParseJson(winH.ResponseText)

    If JSONResult("status") = "Succes" Then
        For i = 1 To JSONResult("aviso").Count
            Select Case JSONResult("aviso")(i)("Tipo")
            Case "AVISOS"
                    MsgBox JSONResult("aviso")(i)("Aviso"), vbCritical
            Case "COBRO"
                    MsgBox JSONResult("aviso")(i)("Aviso"), vbCritical
            Case "SUSPENSION"
                    MsgBox JSONResult("aviso")(i)("Aviso"), vbCritical
                    End
            Case "USUARIOS"
                    Dim UsuariosSisma As Integer, InactivosSisma As Integer
                    Dim LoginSisma As String
                    Dim xR As New ADODB.Recordset
                    
                    MsgBox JSONResult("aviso")(i)("Aviso"), vbCritical
                    Set JSONSend = New Dictionary
                    Set winH = New WinHttp.WinHttpRequest
                    
                    UsuariosSisma = Conn.Execute("Select Count(*) from Usuarios_t Where Estado=1").Fields(0)
                    InactivosSisma = Conn.Execute("Select Count(*) from Usuarios_t Where Estado=0").Fields(0)
                    
                    Set xR = Conn.Execute("Select Login from Usuarios_t Where Estado=1")
                    Do While Not xR.EOF
                        LoginSisma = LoginSisma + xR!Login + ","
                        xR.MoveNext
                    Loop
                    LoginSisma = Mid(LoginSisma, 1, Len(LoginSisma) - 1)
                    
                    JSONSend.Item("usuariosActivos") = UsuariosSisma
                    JSONSend.Item("usuariosInactivos") = InactivosSisma
                    JSONSend.Item("loginActivos") = LoginSisma
    
                    winH.Open "post", "http://18.229.172.128:3000/api/v1/usuarios", False
                    winH.SetRequestHeader "Content-type", "application/json"
                    winH.SetRequestHeader "charset", "UTF-8"
                    winH.SetRequestHeader "token", sTokenSisma
                    winH.SetTimeouts 500, 500, 1000, 1000
                    
                    winH.Send JsonConverter.ConvertToJson(JSONSend)
                    Set JSONResultREV = JsonConverter.ParseJson(winH.ResponseText)
                    If JSONResultREV("status") = "FAILED" Then
                        MsgBox " Error de TOKEN, la ejecucion de la Alplicacion será suspendida"
                        End
                    End If
            Case "REVISION"
                    MsgBox JSONResult("aviso")(i)("Aviso"), vbCritical
                    Set JSONSend = New Dictionary
                    Set winH = New WinHttp.WinHttpRequest
                    
                    JSONSend.Item("empresa") = SNitEmpresaSismaG
                    JSONSend.Item("aplicacion") = sAplicacionSisma
                    JSONSend.Item("serie") = Trim(Nserie)
    
                    winH.Open "post", "http://18.229.172.128:3000/api/v1/movimiento", False
                    winH.SetRequestHeader "Content-type", "application/json"
                    winH.SetRequestHeader "charset", "UTF-8"
                    winH.SetRequestHeader "token", sTokenSisma
                    winH.SetTimeouts 500, 500, 1000, 1000
                    
                    winH.Send JsonConverter.ConvertToJson(JSONSend)
                    Set JSONResultREV = JsonConverter.ParseJson(winH.ResponseText)
                    If JSONResultREV("status") = "FAILED" Then
                        MsgBox " Error de TOKEN, la ejecucion de la Alplicacion será suspendida"
                        End
                    End If
            End Select
        Next i
    End If
    SLastTime = 0
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    'Msg = "Se produjo un error en procedimiento Timer," & vbCrLf & Err.Description
    'MsgBox Msg, , "Timer1_Timer()"
    Err.Clear   ' Borra campos del objeto Err
    Exit Sub
End If
End Sub

