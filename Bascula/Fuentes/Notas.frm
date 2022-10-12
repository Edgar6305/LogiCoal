VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Notas 
   Caption         =   "Notas Debito y Credito"
   ClientHeight    =   6315
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12900
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Notas.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6315
   ScaleWidth      =   12900
   StartUpPosition =   1  'CenterOwner
   Begin MBSplit.Splitter VSplitter 
      Height          =   4935
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   8705
      SplitterPos     =   40
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4635
         Left            =   5640
         TabIndex        =   2
         Top             =   120
         Width           =   6735
         Begin MSComCtl2.DTPicker FechaNota 
            Height          =   315
            Left            =   1800
            TabIndex        =   21
            Top             =   840
            Width           =   1755
            _ExtentX        =   3096
            _ExtentY        =   556
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
            CustomFormat    =   "dd/MM/yyyy hh:mm"
            Format          =   115146755
            CurrentDate     =   44769
         End
         Begin MSMask.MaskEdBox MemVar_3 
            Height          =   315
            Left            =   1800
            TabIndex        =   17
            Top             =   1620
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   "#,##0"
            PromptChar      =   "_"
         End
         Begin VB.TextBox MemVar_4 
            BackColor       =   &H00FFFEEA&
            Height          =   1515
            Left            =   1800
            MaxLength       =   255
            MultiLine       =   -1  'True
            TabIndex        =   16
            Top             =   2400
            Width           =   4215
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            ItemData        =   "Notas.frx":058A
            Left            =   1800
            List            =   "Notas.frx":0594
            TabIndex        =   14
            Text            =   "DEBITO"
            Top             =   1980
            Width           =   1575
         End
         Begin VB.TextBox MemVar_2 
            Height          =   315
            Left            =   1800
            MaxLength       =   50
            TabIndex        =   4
            Top             =   1260
            Width           =   555
         End
         Begin VB.TextBox MemVar_1 
            Height          =   315
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   3
            Top             =   480
            Width           =   1080
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Fecha "
            Height          =   195
            Index           =   1
            Left            =   420
            TabIndex        =   20
            Top             =   900
            Width           =   1095
         End
         Begin VB.Label LabelPila 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   2520
            TabIndex        =   18
            Top             =   1260
            Width           =   3435
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Observaciones"
            Height          =   195
            Index           =   1
            Left            =   420
            TabIndex        =   15
            Top             =   2460
            Width           =   1185
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cantidad - Ton."
            Height          =   195
            Left            =   420
            TabIndex        =   8
            Top             =   1680
            Width           =   1200
         End
         Begin VB.Label Label11 
            Caption         =   "Pila"
            Height          =   195
            Left            =   420
            TabIndex        =   7
            Top             =   1320
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Numero Nota"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   6
            Top             =   540
            Width           =   1035
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Debito/Credito"
            Height          =   195
            Index           =   0
            Left            =   420
            TabIndex        =   5
            Top             =   2040
            Width           =   1155
         End
      End
      Begin MSComctlLib.TreeView oTree 
         Height          =   3735
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   6588
         _Version        =   393217
         Style           =   7
         ImageList       =   "ImageList1"
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
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   3300
         Top             =   3720
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   20
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":05A9
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":0703
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":085D
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":09B7
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":0B11
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":0C6B
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":0DC5
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":0F1F
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":1079
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":11D3
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":132D
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":1487
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":1E99
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":2433
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":29CD
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":2F67
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":3AB1
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":4B83
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":5CD9
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Notas.frx":62EF
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   750
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   19815
      _ExtentX        =   34951
      _ExtentY        =   1323
      BandCount       =   5
      _CBWidth        =   19815
      _CBHeight       =   750
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
      Child4          =   "Actualizar"
      MinHeight4      =   315
      Width4          =   1500
      NewRow4         =   0   'False
      MinHeight5      =   315
      Width5          =   1305
      NewRow5         =   0   'False
      Begin KewlButtonz.KewlButtons Actualizar 
         Height          =   315
         Left            =   4935
         TabIndex        =   19
         Top             =   390
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Actualizar"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Notas.frx":664D
         PICN            =   "Notas.frx":6669
         PICH            =   "Notas.frx":6C03
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin KewlButtonz.KewlButtons oNuevo 
         Height          =   315
         Left            =   3405
         TabIndex        =   12
         Top             =   390
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Nuevo"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   -1  'True
         BCOL            =   13160660
         BCOLO           =   13160660
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "Notas.frx":719D
         PICN            =   "Notas.frx":71B9
         PICH            =   "Notas.frx":7753
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.Toolbar Tbar 
         Height          =   330
         Left            =   165
         TabIndex        =   11
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
               Key             =   "Borrar"
               Object.ToolTipText     =   "Borrar Registro Actual"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Imprimir"
               Object.ToolTipText     =   "Imprimir"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Top"
               Object.ToolTipText     =   "Primer Registro"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Previo"
               Object.ToolTipText     =   "Registro Anterior"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Proximo"
               Object.ToolTipText     =   "Próximo Registro"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Bottom"
               Object.ToolTipText     =   "Ultmo Registro"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Foto"
               Object.ToolTipText     =   "Insetar Imagen"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Estado"
               Object.ToolTipText     =   "Salir"
               ImageIndex      =   32
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
         ItemData        =   "Notas.frx":7CED
         Left            =   825
         List            =   "Notas.frx":7CFD
         TabIndex        =   10
         Text            =   "TODAS"
         Top             =   390
         Width           =   2355
      End
   End
   Begin ComctlLib.StatusBar oBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   13
      Top             =   6000
      Width           =   12900
      _ExtentX        =   22754
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   3
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "USUARIO"
            TextSave        =   "USUARIO"
            Key             =   "Usuario"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   6174
            MinWidth        =   6174
            Text            =   "FECHA CREACION"
            TextSave        =   "FECHA CREACION"
            Key             =   "Ot"
            Object.Tag             =   ""
         EndProperty
         BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "ESTADO"
            TextSave        =   "ESTADO"
            Key             =   "Estado"
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Notas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim okUnload As Boolean
Dim IsNewRecord As Boolean
Dim OkOpen As Boolean
Dim oT As New ADODB.Recordset

Private Sub Actualizar_Click()
If MsgBox("Esta Seguro de Actualizar La Nota de Inventarios ", vbYesNo, "Notas de Inventarios") = vbYes Then
    xSql = "UPDATE Ajustes Set Estado = 'AC' Where IdNotas=" & MemVar_1
    Conn.Execute (xSql)
    Call MemVar_1_LostFocus
End If
End Sub

Private Sub Form_Activate()
If Not OkOpen Then
     Call LoadData
     OkOpen = True
    MemVar_1.SetFocus
End If
Me.SetFocus
Call MuestraArbol
End Sub

Private Sub Form_Resize()
    Dim lSplitHeight As Long
    If Me.WindowState = vbMinimized Then Exit Sub
    'set the position/size of the command buttons and the logo picture
    'set the position/size of the splitter
    lSplitHeight = Me.ScaleHeight - CoolBar1.Height - oBar.Height - 100
    VSplitter.Height = IIf(lSplitHeight < 0, 0, lSplitHeight)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    oT.Close
    OkOpen = False
    Unload Me
End Sub

Private Sub Form_Load()

xSql = "Select Top 1 * From Ajustes " & xFilter
oT.Open xSql, Conn, 2, 3, 1

Set VSplitter.LeftOrTopCtl = oTree
Set VSplitter.RightOrBottomCtl = Frame1

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
                If MemVar_1 = "" And Not IsNewRecord Then
                    MemVar_1.SetFocus
                Else
                    MemVar_3.SetFocus
                End If
End Select
End Sub

Private Sub MemVar_2_LostFocus()
Dim xC As New ADODB.Recordset

If MemVar_2 <> "" Then
    Set xC = Conn.Execute("Select * From vPilasAcopiosGeneral Where IdPila=" & MemVar_2)
    If Not xC.EOF Then
        LabelPila = xC!Despila & " (" & xC!Desacopio & ")"
    Else
        MsgBox "Pila NO localizada, Verifique", vbInformation
        MemVar_2.SetFocus
    End If
End If

End Sub

Private Sub MemVar_3_GotFocus()
        Call Mark(MemVar_3)
End Sub

Private Sub MemVar_3_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
                Call omenu("Browse")
        Case vbKeyUp
                MemVar_2.SetFocus
        Case vbKeyDown, vbKeyReturn
                MemVar_4.SetFocus
End Select
End Sub

Private Sub MemVar_3_LostFocus()
If MemVar_3 = "" Then
    MsgBox "Debe escribir el Cargo del Usuario, Verifique", vbInformation
    MemVar_2.SetFocus
End If

End Sub

Private Sub oNuevo_Click()
If oNuevo.Caption = "Nuevo" Then
    oNuevo.Caption = "Cancelar"
    IsNewRecord = True
    MemVar_1.text = ""
    MemVar_2.text = ""
    MemVar_3.text = ""
    Combo1.ListIndex = 0
    MemVar_7 = "IN"
    LabelPila = ""

    oBar.Panels("Usuario").text = "USUARIO: " & Susuario
    oBar.Panels("Ot").text = "FECHA CREACION: " & Now
    oBar.Panels("Estado").text = "ESTADO: " & "IN"
    
    Tbar.buttons("Grabar").Enabled = True
    
    MemVar_2.SetFocus
Else
    oNuevo.Caption = "Nuevo"
    IsNewRecord = False
    MemVar_1.SetFocus
End If

End Sub

Private Sub Tbar_ButtonClick(ByVal Button As MSComctlLib.Button)
        Call omenu(Button.Key)
End Sub

Private Sub omenu(ByVal Key As String)

On Error GoTo Recover
Select Case Key
    Case "Grabar"
           Call SaveData
           Call MuestraArbol
    Case "Salida"
        Unload Me
    Case "Browse"
            Select Case ActiveControl.Name
            Case "MemVar_1"
                    BrowseCatalogo.x_Tipo = 1
                    BrowseCatalogo.xtabla = "NotasDedCre"
                    Set BrowseCatalogo.dControl = ActiveControl
                    BrowseCatalogo.Show 1
            Case "MemVar_2"
                    BrowseAcopioPilas.x_Tipo = 1
                    BrowseAcopioPilas.xtabla = "vPilasAcopiosGeneral"
                    Set BrowseAcopioPilas.dControl = ActiveControl
                    BrowseAcopioPilas.Show 1
            End Select
            
    Case "Imprime"
    
    Case "Borrar"
    
    Case "Top"
        oT.Close
        oT.Open "Select Top 1 * From Ajustes Order By IdNotas"
        Call LoadData
    Case "Bottom"
        oT.Close
        oT.Open "Select Top 1 * From Ajustes Order By IdNotas DESC"
        Call LoadData
    Case "Proximo"
        oT.Close
        oT.Open "Select Top 1 * From Ajustes Where IdNotas>'" & MemVar_1 & "' Order By  IdNotas"
        Call LoadData
    Case "Previo"
        oT.Close
        oT.Open "Select Top 1 * From Ajustes Where IdNotas<'" & MemVar_1 & "' Order By IdNotas DESC"
        Call LoadData
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
On Error GoTo Recover

 If oT.EOF And oT.BOF Then Exit Sub
 
 If Not IsNewRecord Then
     MemVar_1 = oT!IdNotas
     MemVar_2 = oT!Pila
     MemVar_3 = oT!Cantidad
     FechaNota = oT!Fecha
     MemVar_4 = oT!Observaciones
     Combo1.ListIndex = IIf(oT!Tipo = "D", 0, 1)
     MemVar_7 = oT!Estado
     LabelPila = Conn.Execute("Select DesPila From vPilasAcopiosGeneral Where IdPila=" & MemVar_2).Fields(0)
 Else
 
 End If
    
 oBar.Panels("Usuario").text = "USUARIO: " & oT!Usuario
 oBar.Panels("Ot").text = "FECHA CREACION: " & oT!Fecha
 oBar.Panels("Estado").text = "ESTADO: " & oT!Estado

 Actualizar.Enabled = (oT!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','EstadoNotas'").Fields(0)) '==> Fields(0) Nos dice si tiene acceso al Objeto

 Tbar.buttons("Grabar").Enabled = (oT!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','" & sProgSISMA & "'").Fields(1)) '==> Fields(1) Nos dice puede hacer cambios al Objeto
 Tbar.buttons("Borrar").Enabled = (oT!Estado = "IN" And Conn.Execute("PA_AutorizacionObjetos '" & Susuario & " ','" & sProgSISMA & "'").Fields(1)) '==> Fields(1) Nos dice puede hacer cambios al Objeto

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al leer Transaccion," & vbCrLf & Err.Description
    MsgBox MSG, , "LoadData"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If
    
End Sub

Private Sub SaveData()
Dim xSql As String
Dim Res As String
Dim xR As New ADODB.Recordset

On Error GoTo Recover

'ALTER PROCEDURE [dbo].[PA_Ajustes]
'@Pila As Int,
'@Cantidad AS Int,
'@Tipo AS Varchar(1),
'@Observaciones As Varchar(MAX),
'@Usuario As Varchar(10)

 If IsNewRecord Then
     xSql = "EXEC PA_Ajustes " & MemVar_2 & "," & MemVar_3 & ",'" & Mid(Combo1.text, 1, 1) & "','" & MemVar_4 & "','" & Susuario & "'"
 
    Set xR = Conn.Execute(xSql)
    Res = xR!Res
    
    If Res <> "OK" Then
       MsgBox "Error al Grabar Nota , Verifique" & vbCrLf & Res, vbInformation, "Error de Grabación"
       Exit Sub
    End If
 Else
     xSql = "UPDATE Ajustes Set  IdPila = " & MemVar_2 & "," & Tipo = Mid(Combo1, 1, 1) & "," & Observaciones = MemVar_4 & "," & Cantidad = MemVar_3 & " WHERE IdNota=" & MemVar_1
 End If

 If IsNewRecord Then
    MemVar_1 = xR!Numero
    Call MuestraArbol
    Call MemVar_1_LostFocus
 End If
 oNuevo.Caption = "Nuevo"
 IsNewRecord = False
 xR.Close
 
Exit Sub

Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Salvar los Datos," & vbCrLf & Err.Description
    MsgBox MSG, , "Savedata()"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

End Sub

Private Sub MuestraArbol()
Dim Em As New ADODB.Recordset
Dim xSql As String

On Error GoTo Recover

oTree.Nodes.Clear
Set Nodx = oTree.Nodes.Add(, , "0" & "Notas", "Notas", 15)

xSql = "SELECT DISTINCT DesAcopio, Ubicacion FROM  vPilasAcopiosGeneral"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("0" & "Notas", tvwChild, "A" & Em!Desacopio, Em!Desacopio & " " & Em!Ubicacion, 18)
    Em.MoveNext
Wend

xSql = "SELECT * FROM  vPilasAcopiosGeneral"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("A" & Em!Desacopio, tvwChild, "B" & Str(Em!IdPila), Str(Em!IdPila) + " " + Em!Despila, 17)
    Em.MoveNext
Wend


xSql = "SELECT * FROM Ajustes"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("B" & Str(Em!Pila), tvwChild, "C" & Format(Em!IdNotas, "00000"), "Nota No " & Format(Em!IdNotas, "0000") & " Tipo " & Em!Tipo & " Cant. " & Format(Em!Cantidad, "###,###"), 20)
    Em.MoveNext
Wend
    
For Each loNode In oTree.Nodes
    If loNode.children >= 1 Then
        loNode.Expanded = True
    End If
Next
    
Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error en insercion de nodo," & vbCrLf & Err.Description
    MsgBox MSG, , "Muestra Arbol"
    Err.Clear   ' Borra campos del objeto Err
    Resume Next
End If

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

Private Sub MemVar_1_GotFocus()
        Call Mark(MemVar_1)
End Sub

Private Sub MemVar_1_LostFocus()
Dim xSql As String

On Error GoTo Recover

If MemVar_1 <> "" Then
    If Not IsNewRecord Then
         If Not MySeek(oT, Conn, "Ajustes", "IdNotas='" & MemVar_1 & "'") Then
                Call LoadData
         Else
                MsgBox "Nota de Ajuste NO Registrado, Verifique"
                MemVar_1.SetFocus
        End If
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga La Nota" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_1_LostFocus()"
    Err.Clear
    Resume Next
End If

End Sub

Private Sub oTree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer

Select Case Mid(Node.Key, 1, 1)
Case "C"
    MemVar_1 = Mid(Node.Key, 2, 10)
    Call MemVar_1_LostFocus
End Select

End Sub

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
    dControl.SelStart = 0
    dControl.SelLength = Len(dControl.text)
End If
End Sub



