VERSION 5.00
Begin VB.Form PassWord 
   BorderStyle     =   0  'None
   Caption         =   "PassWord"
   ClientHeight    =   2910
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5265
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   5265
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox SingUp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   0
      Picture         =   "PassWord.frx":0000
      ScaleHeight     =   2865
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.PictureBox Picture5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   3480
         Picture         =   "PassWord.frx":1A0AA
         ScaleHeight     =   615
         ScaleWidth      =   1455
         TabIndex        =   6
         Top             =   2160
         Width           =   1455
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   3420
         Picture         =   "PassWord.frx":1D24C
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   3
         Top             =   1320
         Visible         =   0   'False
         Width           =   255
      End
      Begin VB.TextBox SusuarioSisma 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   2340
         TabIndex        =   2
         Top             =   1320
         Width           =   975
      End
      Begin VB.TextBox sPassword 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         IMEMode         =   3  'DISABLE
         Left            =   2340
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1860
         Visible         =   0   'False
         Width           =   795
      End
      Begin VB.Label DataBase 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   7
         Top             =   2340
         Width           =   2535
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   1
         X1              =   900
         X2              =   3660
         Y1              =   2100
         Y2              =   2100
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   900
         TabIndex        =   5
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   900
         TabIndex        =   4
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00C0C0C0&
         Index           =   0
         X1              =   900
         X2              =   3660
         Y1              =   1620
         Y2              =   1620
      End
   End
End
Attribute VB_Name = "PassWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public dControl As Control

Private Sub Form_Activate()
    SusuarioSisma.SetFocus
End Sub

Private Sub Form_Resize()
SingUp.Height = Me.Height
SingUp.Width = Me.Width
End Sub

Private Sub Picture5_Click()
End
End Sub

Private Sub sPassword_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyF4
            Case vbKeyUp
            Case vbKeyDown, vbKeyReturn
                Call sPassword_LostFocus
    End Select
End Sub

Private Sub sPassword_LostFocus()
Dim xR As New ADODB.Recordset
Dim i As Integer

If sPassword <> "" Then
    Set xR = Conn.Execute("SELECT * FROM Usuarios Where Login='" & SusuarioSisma & "'")
    If xR.EOF Then
        MsgBox "Error de Archivo EOF, Verifique"
        SusuarioSisma.SetFocus
    Else
        If sPassword = Enc_Leer(xR!PassWord) Then
            SingUp.Visible = False
            Susuario = SusuarioSisma
            dControl.Tag = 1
            Unload Me
        Else
            MsgBox "Clave Errada, Verifique, Verifique"
            sPassword.SetFocus
            Exit Sub
        End If
    End If
    xR.Close
End If
    
End Sub

Private Sub SusuarioSisma_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
            Case vbKeyF4
            Case vbKeyUp
            Case vbKeyDown, vbKeyReturn
                 Call SusuarioSisma_LostFocus
                 sPassword.Visible = True
    End Select
End Sub

Private Sub SusuarioSisma_LostFocus()
Dim xR As New ADODB.Recordset

If SusuarioSisma <> "" Then
    Set xR = Conn.Execute("SELECT * FROM Usuarios_T Where Login='" & SusuarioSisma & "'")
    
    If xR.EOF Then
        MsgBox "Usuario NO Registrado, Verifique"
        SusuarioSisma.SetFocus
    Else
        Picture3.Visible = True
        sPassword.Visible = True
        sPassword.SetFocus
    End If
End If

End Sub

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
                dControl.SelStart = 0
                dControl.SelLength = Len(dControl.text)
End If
End Sub
