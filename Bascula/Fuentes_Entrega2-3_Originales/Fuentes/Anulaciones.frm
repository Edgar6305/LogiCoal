VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Anulaciones 
   Caption         =   "Anulaciones"
   ClientHeight    =   3570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6975
   LinkTopic       =   "Form1"
   ScaleHeight     =   3570
   ScaleWidth      =   6975
   StartUpPosition =   2  'CenterScreen
   Begin MSComCtl2.DTPicker MemVar_3 
      Height          =   285
      Left            =   1740
      TabIndex        =   11
      Top             =   1170
      Width           =   1410
      _ExtentX        =   2487
      _ExtentY        =   503
      _Version        =   393216
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   109117441
      CurrentDate     =   40925
   End
   Begin VB.TextBox MemVar_4 
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1515
      Left            =   1740
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1560
      Width           =   4635
   End
   Begin VB.TextBox MemVar_2 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1740
      MaxLength       =   10
      TabIndex        =   4
      Top             =   840
      Width           =   1080
   End
   Begin VB.TextBox MemVar_1 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1740
      MaxLength       =   2
      TabIndex        =   3
      Top             =   480
      Width           =   1080
   End
   Begin VB.PictureBox okFind 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2880
      Picture         =   "Anulaciones.frx":0000
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.PictureBox okNew 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   2880
      Picture         =   "Anulaciones.frx":0102
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   240
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3420
      Top             =   765
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Anulaciones.frx":0204
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Anulaciones.frx":035E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Anulaciones.frx":04B8
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Anulaciones.frx":0612
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Anulaciones.frx":076C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Anulaciones.frx":08C6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Anulaciones.frx":0A20
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Anulaciones.frx":0B7A
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Anulaciones.frx":0CD4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Anulaciones.frx":0E2E
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Anulaciones.frx":0F88
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Anulaciones.frx":10E2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   390
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   688
      BandCount       =   1
      _CBWidth        =   13335
      _CBHeight       =   390
      _Version        =   "6.0.8169"
      Child1          =   "Tbar"
      MinHeight1      =   330
      Width1          =   495
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   330
         Left            =   30
         TabIndex        =   10
         Top             =   30
         Width           =   13215
         _ExtentX        =   23310
         _ExtentY        =   582
         ButtonWidth     =   609
         ButtonHeight    =   582
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Grabar"
               Object.ToolTipText     =   "Grabar Registro Actual"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
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
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Salida"
               Object.ToolTipText     =   "Salir"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Debe escribir minimo 35 Caracreres en la descripcion"
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
      Left            =   1740
      TabIndex        =   12
      Top             =   3180
      Width           =   4635
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Observación"
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
      Left            =   180
      TabIndex        =   8
      Top             =   1620
      Width           =   960
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha"
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
      Left            =   180
      TabIndex        =   7
      Top             =   1260
      Width           =   450
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Número Doc."
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
      Left            =   180
      TabIndex        =   6
      Top             =   900
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Transacción"
      Height          =   195
      Left            =   180
      TabIndex        =   5
      Top             =   540
      Width           =   885
   End
End
Attribute VB_Name = "Anulaciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oT As New ADODB.Recordset
Dim okUnload As Boolean
Const CONTSTOP = -1
Const maxView = 0
Const SpaceY = 300
Dim IsNewRecord As Boolean
Dim MaxUser As Integer
Dim OkOpen As Boolean
Public CodigoItem As Long ' Parametro del codigo del item cuando applica
Public dControl As Control
Dim xBorrar As Boolean

Private Sub Form_Activate()
    If Not OkOpen Then
            MaxUser = CONTSTOP
            OkOpen = True
            Call MemVar_2_LostFocus
            Tbar.Buttons("Grabar").Enabled = False
            xBorrar = False
    End If
    Me.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
    oT.Close
    OkOpen = False
    dControl.Tag = xBorrar
    Unload Me
End Sub

Private Sub Form_Load()
   oT.Open "Anulaciones", Conn, 2, 3, 512
End Sub

Private Sub Memvar_4_KeyPress(KeyAscii As Integer)
   Tbar.Buttons("Grabar").Enabled = IIf(Len(Trim(MemVar_4)) < 40, False, True)
End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
        Call omenu(Button.Key)
End Sub

Private Sub omenu(ByVal Key As String)
        Select Case Key
                Case "Grabar"
                      Call SaveData
                Case "Salida"
                      Unload Me
                Case "Browse"
                        Select Case ActiveControl.Name
                        Case "MemVar_1"
                                BrowseCatalogo.x_Tipo = 1
                                BrowseCatalogo.xtabla = "Iva"
                                Set BrowseCatalogo.dControl = ActiveControl
                                BrowseCatalogo.Show
                        End Select
                   Case "Imprimir"
                Case "Borrar"
                    On Error GoTo BorraRec
                    If MsgBox("Esta seguro de Borrar el Registro " + Chr(13) + Chr(10) + "Las relaciones si las hay seran borradas", vbYesNo, "Borrado de Registro") = vbYes Then
                    oT.Delete
                    oT.MoveFirst
                    Call LoadData
                    End If
                    GoTo Salida
BorraRec:
                    MsgBox "Error al borrar el registro, existe relaciones con otros archivos las cuales " + Chr(13) + Chr(10) + "             no permiten que se borre este registro, verifique"
                On Error GoTo 0
                Case "Top"
                    oT.Close
                    oT.Open "Select Top 1 * From Iva Order By CodigoIva"
                    Call LoadData
                Case "Bottom"
                    oT.Close
                    oT.Open "Select Top 1 * From Iva Order By CodigoIva DESC"
                    Call LoadData
                Case "Proximo"
                    oT.Close
                    oT.Open "Select Top 1 * From Iva Where CodigoIva>'" & MemVar_1 & "' Order By CodigoIva"
                    If oT.EOF Then
                       oT.Close
                       oT.Open "Select Top 1 * From Iva Where CodigoIva>='" & MemVar_1 & "' Order By CodigoIva"
                    End If
                    Call LoadData
                Case "Previo"
                    oT.Close
                    oT.Open "Select Top 1 * From Iva Where CodigoIva<'" & MemVar_1 & "' Order By CodigoIva DESC"
                    If oT.EOF Then
                       oT.Close
                       oT.Open "Select Top 1 * From Iva Where CodigoIva<='" & MemVar_1 & "' Order By CodigoIva"
                    End If
                    Call LoadData
                
Salida:
        End Select
End Sub
Private Sub ExeBrow(oCod As String)
' Codigo Por desarrollar
End Sub

Private Sub LoadData()
        If oT.EOF And oT.BOF Then
                okNew.Visible = True
                IsNewRecord = True
                Exit Sub
        End If
        okNew.Visible = False
        okFind.Visible = True
        MemVar_1.Text = oT!Transaccion
        MemVar_2.Text = oT!NumeroDoc
        MemVar_3 = oT!Fecha
        MemVar_4.Text = oT!Observacion
End Sub

Private Sub SaveData()
Dim ok As Boolean
On Error GoTo SalvaGraba
ok = False
        If MemVar_1.Text = "" Then ok = True
        If ok Then Exit Sub
        If IsNewRecord Then
           oT.AddNew
        End If
        oT!Transaccion = MemVar_1.Text
        oT!NumeroDoc = MemVar_2.Text
        oT!Fecha = MemVar_3
        oT!Observacion = Mid(MemVar_4.Text, 1, 254)
        oT!Usuario = Susuario
        oT.Update
        xBorrar = True
        If IsNewRecord Then
                If Not MySeek(oT, Conn, "Anulaciones", "Transaccion='" & MemVar_1 & "' And Numerodoc=" & MemVar_2) Then
                okFind.Visible = True
                okNew.Visible = False
            End If
        End If
        IsNewRecord = False
        GoTo Salida
SalvaGraba:
        If Err.Number <> 0 Then
        Msg = Err.Description
        MsgBox Msg, , "Error de Datos", Err.HelpFile, Err.HelpContext
        End If
        On Error GoTo 0
Salida:

End Sub
Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
                dControl.SelStart = 0
                dControl.SelLength = Len(dControl.Text)
End If
End Sub
Private Function Validate(X As String, Y As Byte) As Boolean
        Validate = True
End Function
Private Sub MemVar_1_GotFocus()
        Call Mark(MemVar_1)
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
'Private Sub MemVar_1_Change()
        'MemVar_1.DataChanged = True
'End Sub
Private Sub MemVar_1_LostFocus()
        If Not Validate(MemVar_1.Text, 10) Then
                        MemVar_1.SetFocus
        Else
        End If
End Sub
Private Sub MemVar_2_GotFocus()
        Call Mark(MemVar_2)
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
'Private Sub MemVar_2_Change()
        'MemVar_2.DataChanged = True
'End Sub
Private Sub MemVar_2_LostFocus()
        If Not Validate(MemVar_2.Text, 10) Then
                        MemVar_2.SetFocus
        Else
                If Not MySeek(oT, Conn, "Anulaciones", "Transaccion='" & MemVar_1 & "' And Numerodoc=" & MemVar_2) Then
                        Call LoadData
                        okFind.Visible = True
                        okNew.Visible = False
                        IsNewRecord = False
                Else
                        Dim i As Integer
                        okFind.Visible = False
                        okNew.Visible = True
                        IsNewRecord = True
                End If
                MemVar_4.SetFocus
        End If
End Sub
Private Sub MemVar_3_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                Case vbKeyUp
                        MemVar_2.SetFocus
                Case vbKeyDown, vbKeyReturn
                        MemVar_4.SetFocus
End Select
End Sub
Private Sub MemVar_3_LostFocus()
        If Not Validate(MemVar_3, 6) Then
                        MemVar_3.SetFocus
        Else
        End If
End Sub
Private Sub MemVar_4_GotFocus()
        Call Mark(MemVar_4)
End Sub
Private Sub MemVar_4_KeyDown(KeyCode As Integer, Shift As Integer)
        Select Case KeyCode
                Case vbKeyF4
                Case vbKeyUp
                Case vbKeyDown, vbKeyReturn
                     If Tbar.Buttons("Grabar").Visible Then Call omenu("Grabar")
End Select
End Sub

Private Sub MemVar_4_LostFocus()
        If Len(MemVar_4) < 25 Then
           MsgBox "El texto debe tener almenos 25 Caracteres", vbExclamation
           MemVar_4.SetFocus
           Tbar.Buttons("Grabar").Enabled = False
        Else
           Tbar.Buttons("Grabar").Enabled = True
        End If
End Sub
Private Sub ShowMsg(frm As Form, ctl As Control, ttl As String, flgs As Integer, Optional oControlAux As Variant)
        Call OWNMSG.SetPos(frm, ctl, ttl, flgs, oControlAux)
End Sub



