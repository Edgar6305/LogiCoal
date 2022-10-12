VERSION 5.00
Begin VB.Form EnConstruccion 
   BorderStyle     =   0  'None
   ClientHeight    =   2940
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5220
   LinkTopic       =   "Form1"
   ScaleHeight     =   2940
   ScaleWidth      =   5220
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Pic1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      FillColor       =   &H00FF0000&
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   0
      Picture         =   "EnConstruccion.frx":0000
      ScaleHeight     =   2865
      ScaleWidth      =   5145
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin VB.PictureBox Pic2 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1275
         Left            =   660
         Picture         =   "EnConstruccion.frx":1A0AA
         ScaleHeight     =   1275
         ScaleWidth      =   3615
         TabIndex        =   1
         Top             =   1140
         Width           =   3615
      End
   End
End
Attribute VB_Name = "EnConstruccion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Activate()
Me.SetFocus
End Sub

Private Sub Form_Resize()
    Pic1.Height = Me.Height
    Pic1.Width = Me.Width
End Sub

Private Sub Pic2_Click()
Unload Me
End Sub
