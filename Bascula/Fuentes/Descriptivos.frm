VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0BA686C6-F7D3-101A-993E-0000C0EF6F5E}#1.0#0"; "THREED32.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Begin VB.Form Descriptivos 
   Caption         =   "Descripciones Gnerales"
   ClientHeight    =   9285
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13965
   LinkTopic       =   "Form1"
   ScaleHeight     =   9285
   ScaleWidth      =   13965
   StartUpPosition =   1  'CenterOwner
   Begin MBSplit.Splitter VSplitter 
      Height          =   8175
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   13635
      _ExtentX        =   24051
      _ExtentY        =   14420
      SplitterPos     =   40
      Begin MSComctlLib.TreeView oTree 
         Height          =   8115
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   14314
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
      Begin VB.Frame Body 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   7935
         Left            =   5880
         TabIndex        =   1
         Top             =   180
         Width           =   7695
         Begin VB.ComboBox Combo1 
            BackColor       =   &H00FFFEEA&
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
            Left            =   480
            TabIndex        =   11
            Top             =   720
            Width           =   6495
         End
         Begin Threed.SSFrame oMarco 
            Height          =   6510
            Left            =   420
            TabIndex        =   2
            Top             =   1200
            Width           =   6555
            _Version        =   65536
            _ExtentX        =   11562
            _ExtentY        =   11483
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
            Enabled         =   -1  'True
            Begin VB.VScrollBar oMove 
               Height          =   6300
               Left            =   6240
               TabIndex        =   4
               Top             =   120
               Width           =   255
            End
            Begin VB.TextBox MemVar_5 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFF4&
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
               Left            =   60
               MaxLength       =   125
               TabIndex        =   3
               Top             =   450
               Width           =   6165
            End
            Begin VB.Label SubLabel3 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFEEA&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "Descripción"
               BeginProperty Font 
                  Name            =   "Segoe UI"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   285
               Index           =   0
               Left            =   60
               TabIndex        =   5
               Top             =   120
               Width           =   6165
            End
         End
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   6900
            Top             =   240
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
                  Picture         =   "Descriptivos.frx":0000
                  Key             =   ""
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":015A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":02B4
                  Key             =   ""
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":040E
                  Key             =   ""
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":0568
                  Key             =   ""
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":06C2
                  Key             =   ""
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":081C
                  Key             =   ""
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":0976
                  Key             =   ""
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":0AD0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":0C2A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":0D84
                  Key             =   ""
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":0EDE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":18F0
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":1E8A
                  Key             =   ""
               EndProperty
               BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":2424
                  Key             =   ""
               EndProperty
               BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":29BE
                  Key             =   ""
               EndProperty
               BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":3EC8
                  Key             =   ""
               EndProperty
               BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":4A12
                  Key             =   ""
               EndProperty
               BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":4FCB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":60E1
                  Key             =   ""
               EndProperty
               BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":68B3
                  Key             =   ""
               EndProperty
               BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":7985
                  Key             =   ""
               EndProperty
               BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":84CF
                  Key             =   ""
               EndProperty
               BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "Descriptivos.frx":9625
                  Key             =   ""
               EndProperty
            EndProperty
         End
         Begin VB.Label Label1 
            Caption         =   "Descriptivo General"
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   480
            TabIndex        =   6
            Top             =   360
            Width           =   2445
         End
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   780
      Left            =   0
      TabIndex        =   8
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
      Width4          =   1785
      NewRow4         =   0   'False
      Child5          =   "MuestraProduccion"
      MinHeight5      =   315
      Width5          =   2325
      NewRow5         =   0   'False
      MinHeight6      =   360
      Width6          =   2400
      NewRow6         =   0   'False
      MinHeight7      =   360
      Width7          =   1500
      NewRow7         =   0   'False
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
         ItemData        =   "Descriptivos.frx":9C3B
         Left            =   825
         List            =   "Descriptivos.frx":9C4B
         TabIndex        =   10
         Text            =   "TODAS"
         Top             =   405
         Width           =   2355
      End
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   330
         Left            =   165
         TabIndex        =   9
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
               Object.Visible         =   0   'False
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Top"
               Object.ToolTipText     =   "Primer Registro"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Previo"
               Object.ToolTipText     =   "Registro Anterior"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Key             =   "Proximo"
               Object.ToolTipText     =   "Próximo Registro"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
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
               Object.Visible         =   0   'False
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
End
Attribute VB_Name = "Descriptivos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim okUnload As Boolean
Const CONTSTOP = 60
Const maxView = 19
Const SpaceY = 300
Dim IsNewRecord As Boolean
Dim MaxUser As Integer
Dim OkOpen As Boolean
Dim R As New ADODB.Recordset

Private Sub Combo1_Click()
Call LoadValores
End Sub

Private Sub Form_Activate()
If Not OkOpen Then
        MaxUser = CONTSTOP
        Call LoadControls
        Call wShow
        If MemVar_1 <> "" Then
           'Set oT = Conn.Execute("Select * From PagosCuentas Where Codigo='" & MemVar_1 & "'")
        End If
        Call LoadData
        OkOpen = True
        Me.SetFocus
        Call MuestraArbol
End If
End Sub

Private Sub Form_Load()
Dim X As New ADODB.Recordset
    Set VSplitter.LeftOrTopCtl = oTree
    Set VSplitter.RightOrBottomCtl = Body
    
    xSql = "Select * From Descriptivos "
    Set xR = Conn.Execute(xSql)
    Combo1.Clear
    
    If Not xR.EOF Then
        Combo1 = Format(xR!IdTipo, "00") + " " + xR!Descripcion
        Do While Not xR.EOF
            Combo1.AddItem Format(xR!IdTipo, "00") + " " + xR!Descripcion
            xR.MoveNext
        Loop
    End If
    
    xSql = "Select * From DescriptivosDetalle"
    R.Open xSql, Conn, 2, 3, 1

End Sub

Private Sub LoadControls()
Dim j As Integer
    For j = 1 To MaxUser
            Load MemVar_5(j)
            If j Mod 2 = 0 Then MemVar_5(j).BackColor = &HFFFEEA   '&HDAFEFB
    Next j
    MemVar_5(0).BackColor = &HFFFEEA
End Sub

Private Sub AjustaMover()
        oMove.Max = MaxUser
        oMove.SmallChange = maxView
        oMove.LargeChange = maxView
        If MaxUser < maxView Then
                oMove.Visible = False
        Else
                oMove.Visible = True
        End If
End Sub

Private Sub wShow()
Dim i As Long
Dim j As Long
        For i = 0 To MaxUser
                If i >= oMove.Value And i <= oMove.Value + maxView Then
                        MemVar_5(i).Top = j * SpaceY + MemVar_5(0).Top
                        MemVar_5(i).Visible = True
                        j = j + 1
                Else
                        MemVar_5(i).Visible = False
                End If
        Next i
End Sub

Private Sub DeleteRow(N As Integer)
Dim i As Integer
    For i = N To MaxUser - 1
            MemVar_5(i).text = MemVar_5(i + 1).text
    Next i
End Sub

Private Sub InsertRow(N As Integer)
Dim i As Integer
    For i = MaxUser To N + 1 Step -1
            MemVar_5(i).text = MemVar_5(i - 1).text
    Next i
    MemVar_5(N).text = ""
End Sub

Private Sub Limpia(ByVal i As Integer)
        MemVar_5(i).text = ""
End Sub

Public Sub LoadValores()
Dim i As Integer, j As Integer
Dim xR As New ADODB.Recordset
Dim xSql As String

For i = 0 To MaxUser
  Call Limpia(i)
  MemVar_5(i).Locked = False
Next i
i = 0
oMove.Value = 0

xSql = "SELECT * From DescriptivosDetalle Where IdTipo=" & Val(Mid(Combo1, 1, 2))
Set xR = Conn.Execute(xSql)

Do While Not xR.EOF
    MemVar_5(i).text = xR!Descripcion
    MemVar_5(i).Locked = True
    i = i + 1
    xR.MoveNext
Loop

Call AjustaMover
Call wShow

End Sub

Private Sub MemVar_5_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
On Error Resume Next

Select Case KeyCode
        Case vbKeyF4
        Case vbKeyReturn
                If Index < MaxUser Then
                        Call revisa(Index)
                        MemVar_5(Index + 1).SetFocus
        End If
        Case vbKeyLeft
                If (Shift And 2) = 2 Then
                        MemVar_5(Index).SetFocus
                End If
        Case vbKeyRight
                If (Shift And 2) = 2 Then
                        If Index < MaxUser Then
                                Call revisa(Index)
                                MemVar_5(Index + 1).SetFocus
                        End If
                End If
        Case vbKeyDown
                If Index < MaxUser Then
                        Call revisa(Index)
                        MemVar_5(Index + 1).SetFocus
                End If
        Case vbKeyUp
                If Index > 0 Then
                        Call rev(Index)
                        MemVar_5(Index - 1).SetFocus
                End If
End Select

End Sub

Private Sub oMove_Change()
        Call wShow
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
        Call omenu(Button.Key)
End Sub

Private Sub omenu(ByVal Key As String)
On Error GoTo Recover
Select Case Key
        Case "Grabar"
               Call SaveValores
               Call MuestraArbol
        Case "Salida"
            Unload Me
        Case "Browse"
        Case "Imprime"
        Case "Borrar"
            If MsgBox("Esta seguro de Borrar La Orden de Traslado", vbYesNo, "Borrado de Registro") = vbYes Then
            End If
        Case "Top"
        Case "Bottom"
        Case "Proximo"
        Case "Previo"
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
Dim xR As New ADODB.Recordset

On Error GoTo Recover

Call LoadValores

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer Transaccion," & vbCrLf & Err.Description
    MsgBox MSG, , "LoadData"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
    
End Sub

Public Sub SaveValores()
Dim i As Long
Dim j As Integer

'    CREATE PROCEDURE [dbo].[PA_Descriptivos]
'    @IdTipo As Int,
'    @Descripcion As Varchar(125),
'    @Usuario As Varchar(15)

On Error GoTo Recover

For i = 0 To MaxUser
    If MySeek(R, Conn, "DescriptivosDetalle", "Descripcion='" & MemVar_5(i) & "' AND IdTipo=" & Val(Mid(Combo1, 2))) Then
        If MemVar_5(i).text <> "" Then
           xSql = "EXEC PA_Descriptivos " & Val(Mid(Combo1, 2)) & ",'" & MemVar_5(i) & "','" & Susuario & "'"
    
           Set xR = Conn.Execute(xSql)
           Res = xR!Res
           
           If Res <> "OK" Then
              MsgBox "Error al Grabar Descriptivo , " & MemVar_5(i) & "Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
           Else
              MemVar_5(i).Locked = True
           End If
           j = j + 1
        
        End If
    End If
Next i

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

Private Function Validate(X As String, Y As Byte) As Boolean
        Validate = True
End Function

Private Sub MuestraArbol()
Dim Em As New ADODB.Recordset
Dim xSql As String

'On Error GoTo Recover

oTree.Nodes.Clear
Set Nodx = oTree.Nodes.Add(, , "A" & "Descriptivos", "Descriptivos", 13)

xSql = "SELECT  * FROM  Descriptivos"
Set Em = Conn.Execute(xSql)

If Em.EOF Then Exit Sub

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("A" & "Descriptivos", tvwChild, "B" & Em!IdTipo, Em!Descripcion, 14)
    Em.MoveNext
Wend

xSql = "SELECT  DescriptivosDetalle.IDTipoDetalle, DescriptivosDetalle.IdTipo, DescriptivosDetalle.Descripcion"
xSql = xSql + " FROM    Descriptivos INNER JOIN"
xSql = xSql + "         DescriptivosDetalle ON Descriptivos.IDTipo = DescriptivosDetalle.IdTipo"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("B" & Em!IdTipo, tvwChild, "C" & Em!IDTipoDetalle, Em!Descripcion, 15)
    Em.MoveNext
Wend
    
For Each loNode In oTree.Nodes
    'If loNode.children = 1 Then
        loNode.Expanded = True
    'End If
Next
            
Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Cargar Orden desde el Arbol," & vbCrLf & Err.Description
    MsgBox MSG, , "MuestraArbol()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
End Sub

Private Sub oTree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer

Select Case Mid(Node.Key, 1, 1)
Case "C"

End Select

End Sub

Private Sub revisa(Index As Integer)
        If Index + 1 >= maxView + oMove.Value Then
                oMove.Value = oMove.Value + oMove.SmallChange
        End If
End Sub

Private Sub rev(Index As Integer)
        If Index - 1 < oMove.Value And Index > 0 Then
                If oMove.Value - oMove.SmallChange >= 0 Then
                        oMove.Value = oMove.Value - oMove.SmallChange
                Else
                        oMove.Value = 0
                End If
        End If
End Sub


