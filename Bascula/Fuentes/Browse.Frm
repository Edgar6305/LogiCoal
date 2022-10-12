VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form BrowseCatalogo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar"
   ClientHeight    =   7035
   ClientLeft      =   1410
   ClientTop       =   1410
   ClientWidth     =   7410
   Icon            =   "Browse.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   7035
   ScaleWidth      =   7410
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox Check1 
      Caption         =   "Búsqueda Avanzada"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5580
      TabIndex        =   6
      Top             =   300
      Width           =   1635
   End
   Begin VB.ComboBox oIndices 
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
      Left            =   2055
      TabIndex        =   0
      Text            =   "Descripcion"
      ToolTipText     =   "Se escoje el Campo por el cual se desea realizar la Busqueda"
      Top             =   360
      Width           =   3255
   End
   Begin VB.CommandButton osale 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   5520
      TabIndex        =   3
      Top             =   840
      Width           =   1500
   End
   Begin VB.TextBox ollave 
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2055
      TabIndex        =   1
      ToolTipText     =   "Se escribe la Descripcion o el codigo por el cual se va a realizar la busqueda"
      Top             =   840
      Width           =   3255
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4785
      Left            =   180
      TabIndex        =   7
      Top             =   1500
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   8440
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16776938
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Caption         =   "Escoja Filtro de Buqueda ==>"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label otabla 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   6360
      Width           =   7140
   End
   Begin VB.Label Label1 
      Caption         =   "Digite el Criterio de Búsqueda ==>"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Index           =   0
      Left            =   360
      TabIndex        =   2
      Top             =   795
      Width           =   1650
   End
End
Attribute VB_Name = "BrowseCatalogo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim BrCa As New ADODB.Recordset

Dim itemx As ListItem
Public dControl As Control, xA As Boolean, xCuantosIndex As Integer
Public xtabla As String, x_Tipo As Integer
Private Sub Mark()
    ActiveControl.SelStart = 0
    ActiveControl.SelLength = Len(ActiveControl.text)
End Sub
Public Sub Muestra()
Dim i As Integer, j As Integer
ListView1.ListItems.Clear
i = 0

Do While i <= 300 And Not BrCa.EOF()
    i = i + 1
    Set iTmx = ListView1.ListItems.Add()
    j = 0
    For Each fl In BrCa.Fields
        If j = 0 Then
           iTmx.text = fl.Value
        Else
           If Not IsNull(fl.Value) Then
              iTmx.SubItems(j) = fl.Value
          End If
        End If
        j = j + 1
    Next
    BrCa.MoveNext
    If BrCa.EOF Then Exit Do
Loop
End Sub

Private Sub Form_Activate()
Me.SetFocus
End Sub

Private Sub Form_Load()
'Dim clmX As ColumnHeader
Dim xDef As TableDef
Dim xIn As Index

If x_Tipo = 1 Then
   BrCa.Open xtabla, Conn
Else
   BrCa.Open xtabla, Conn
End If

If Not BrCa.EOF Then
    otabla.Caption = xtabla
    
    For Each fl In BrCa.Fields
       If fl.type = 200 Or fl.type = 3 Then oIndices.AddItem fl.Name
    Next
    oIndices.text = oIndices.List(0)
     
    For Each fl In BrCa.Fields
        Set clmX = ListView1.ColumnHeaders.Add(, , fl.Name, fl.DefinedSize * 100)
    Next
    Call Muestra
    ollave.text = ListView1.SelectedItem.text
    Check1.Value = 1
Else
    MsgBox "No Hay Datos Que Mostrar "
    Call ListView1_DblClick
End If
End Sub

Private Sub Form_Resize()
On Error Resume Next
ListView1.Width = Me.Width - 500
ListView1.Height = Me.Height - 3000
otabla.Top = Me.Height - 1000
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
dControl.text = ollave.text
dControl.SetFocus
BrCa.Close

End Sub

Private Sub ListView1_Click()
ollave.text = ListView1.SelectedItem.text
End Sub

Private Sub ListView1_DblClick()
On Error Resume Next
Dim xItem As ListItem
Set xItem = ListView1.SelectedItem
dControl.text = xItem.text
dControl.SetFocus
Unload Me
End Sub

Private Sub ListView1_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
 Select Case KeyCode
    Case vbKeyReturn
    ollave.text = ListView1.SelectedItem.text
    Call osale_Click
 End Select
End Sub

Private Sub ListView1_KeyPress(KeyAscii As Integer)
On Error Resume Next
 Select Case KeyAscii
    Case vbKeyReturn
        Dim xItem As ListItem
        Set xItem = ListView1.SelectedItem
        dControl.text = xItem.text
        dControl.SetFocus
        BrCa.Close
 End Select
Unload Me
End Sub

Private Sub oIndices_Click()
ollave.SetFocus
ollave = ""
End Sub

Private Sub ollave_change()
If x_Tipo = 1 Then
    If Check1.Value = 0 Then
        If Not MySeekB(BrCa, Conn, xtabla, oIndices.text & ">='" & ollave & "' ORDER BY " & oIndices.text) Then
           Call Muestra
        End If
        
    Else
       If Not MySeekB(BrCa, Conn, xtabla, oIndices.text & " LIKE'%" & ollave & "%' ORDER BY " & oIndices.text) Then
          Call Muestra
       End If
    End If
End If
End Sub

Private Sub ollave_GotFocus()
Call Mark
End Sub

Private Sub ollave_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call osale_Click
    End If
End Sub
Private Sub osale_Click()
Unload Me
End Sub
Private Function MySeekB(xO As Object, xC As Object, xFile, xWhere As String) As Boolean
Dim xSql As String
On Error Resume Next
xO.Close
xO.Open "Select Top 200 * From " & xFile & " Where " & xWhere, xC, 2, 3, 1
If xO.EOF Then
   xO.Close
   xO.Open "Select * From " & xFile, xC, 2, 3, 1
   MySeekB = True
Else
   MySeekB = False
End If
End Function




