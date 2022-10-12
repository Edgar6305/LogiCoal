VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form BrowseAcopioPilas 
   Caption         =   "Browse Acopios Pilas"
   ClientHeight    =   7515
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5355
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
   ScaleHeight     =   7515
   ScaleWidth      =   5355
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton osale 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1740
      TabIndex        =   0
      Top             =   6960
      Width           =   1500
   End
   Begin ComctlLib.TreeView oDest 
      Height          =   6615
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   11668
      _Version        =   327682
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4200
      Top             =   6840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   11
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BrwseAcopioPilas.frx":0000
            Key             =   "Soporte"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BrwseAcopioPilas.frx":0A92
            Key             =   "Volquetas"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BrwseAcopioPilas.frx":14E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BrwseAcopioPilas.frx":1F92
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BrwseAcopioPilas.frx":2AD4
            Key             =   "Cargadores"
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BrwseAcopioPilas.frx":4FF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BrwseAcopioPilas.frx":8968
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BrwseAcopioPilas.frx":D4FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BrwseAcopioPilas.frx":D6D4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BrwseAcopioPilas.frx":E21E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "BrwseAcopioPilas.frx":F2F0
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "BrowseAcopioPilas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xDe As New ADODB.Recordset
Dim xGr As New ADODB.Recordset
Public dControl As Control, xA As Boolean, xCuantosIndex As Integer
Public xtabla As String, x_Tipo As Integer
Dim xSql As String

Private Sub Form_Activate()
Me.SetFocus
End Sub

Private Sub Form_Load()

xSql = "Select * From vPilasAcopiosGeneral WHERE Estado='IN' Order By Ubicacion,Desacopio"
Call ArmaDestinos(xSql)
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
dControl.SetFocus
End Sub

Private Sub oDest_Click()
Dim xLen As Integer
xLen = Len(oDest.SelectedItem.Key)
If Mid(oDest.SelectedItem.Key, 1, 1) = "D" Then
    dControl.text = Mid(oDest.SelectedItem.Key, 2, xLen)
    Unload Me
End If
End Sub

Private Sub osale_Click()
Unload Me
End Sub

Sub ArmaDestinos(xSql As String)
Dim i As Integer
Dim xPro As String, xFlo As String, xTipoFlo As String
Dim xMaq As String, xPic As Integer
'Dim Nodx As Node
On Error Resume Next

    Set xDe = Conn.Execute(xSql)
    Set Nodx = oDest.Nodes.Add(, , "*" & "Acopios", "Acopios", 11)
    xTipo = "" ' Acopio
    xFlo = ""    ' Ubicacion
    
    While Not xDe.EOF
            If xFlo <> xDe!Ubicacion Then
                Set Nodx = oDest.Nodes.Add("*" & "Acopios", tvwChild, "A" & Trim(xDe!Ubicacion), xDe!Ubicacion, 10)
                xFlo = Trim(xDe!Ubicacion)
           End If
            
            If xTipo <> xDe!Desacopio Then
                Set Nodx = oDest.Nodes.Add("A" & Trim(xDe!Ubicacion), tvwChild, "B" & Trim(xDe!Desacopio), xDe!Desacopio, 10)
                xTipo = Trim(xDe!Desacopio)
           End If
   
          Set Nodx = oDest.Nodes.Add("B" & Trim(xDe!Desacopio), tvwChild, "D" & xDe!IdPila, xDe!Despila, 9)
            
        xDe.MoveNext
    Wend
    xDe.Close

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer Transaccion," & vbCrLf & Err.Description
    MsgBox MSG, , "Armadestinos"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub


