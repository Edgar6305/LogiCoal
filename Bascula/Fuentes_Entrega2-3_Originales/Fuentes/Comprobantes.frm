VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Comprobantes 
   Caption         =   "Comprobantes SieSa"
   ClientHeight    =   3210
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6105
   Icon            =   "Comprobantes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3210
   ScaleWidth      =   6105
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   1995
      Left            =   180
      ScaleHeight     =   1965
      ScaleWidth      =   5685
      TabIndex        =   0
      Top             =   180
      Width           =   5715
      Begin VB.ComboBox Combo1 
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
         Left            =   1680
         TabIndex        =   1
         Top             =   1200
         Width           =   3795
      End
      Begin MSComCtl2.DTPicker xFecFin 
         Height          =   315
         Left            =   1680
         TabIndex        =   2
         Top             =   780
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111345665
         CurrentDate     =   44770
      End
      Begin MSComCtl2.DTPicker xFecIni 
         Height          =   315
         Left            =   1680
         TabIndex        =   3
         Top             =   360
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   111345665
         CurrentDate     =   44770
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "Comprobante"
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
         Index           =   2
         Left            =   180
         TabIndex        =   6
         Top             =   1260
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "Fecha Final"
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
         Index           =   1
         Left            =   180
         TabIndex        =   5
         ToolTipText     =   "Se refiere a la Fecha Turno Final"
         Top             =   780
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000004&
         Caption         =   "Fecha Inicial"
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
         Index           =   0
         Left            =   180
         TabIndex        =   4
         ToolTipText     =   "Se refiere a la Fecha Turno Inicial"
         Top             =   420
         Width           =   1215
      End
   End
   Begin KewlButtonz.KewlButtons Command2 
      Height          =   435
      Left            =   4140
      TabIndex        =   7
      Top             =   2460
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Cancelar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   16777152
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Comprobantes.frx":000C
      PICN            =   "Comprobantes.frx":0028
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons Command1 
      Height          =   435
      Left            =   180
      TabIndex        =   8
      Top             =   2460
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Ejecutar"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   -1  'True
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "Comprobantes.frx":05C2
      PICN            =   "Comprobantes.frx":05DE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
End
Attribute VB_Name = "Comprobantes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim xFile As String
Dim xR As New ADODB.Recordset
Dim xSql As String

Private Sub Command1_Click()
Dim xC As New ADODB.Recordset
Dim Data As String
Dim xNumero As Long
Dim xTipo As String

On Error GoTo Recover

xTipo = Mid(Combo1.text, 1, 3)
Select Case xTipo
Case "INV"
    xSql = "SET DATEFORMAT DMY"
    xSql = xSql + " SELECT IdTiquete, PesoLleno - PesoVacio AS PesoNeto"
    xSql = xSql + " From Bascula"
    xSql = xSql + " WHERE  Estado='AC' AND IdMaterial=1 AND TransaccionOrigen='LT' AND Procesado=0"
    
    Set xR = Conn.Execute(xSql)
    Set xC = Conn.Execute("Select * From NumerosComprobantes Where Tipo='" & xTipo & "'")
    xNumero = xC!Numero
        
        
    If Not xR.EOF Then
        xFile = sComprobantes & "INVENTARIOS.TXT"
        Open xFile For Output As #1
        
        Do While Not xR.EOF
            Data = Format(xNumero, "0000000")
            Data = Data + "04700124"    '==> Fijos
            Data = Data + "001"         '==> Compañia
            Data = Data + Format(xR!PesoNeto, "000000000000000.0000")
            Print #1, Data
            xR.MoveNext
            xNumero = xNumero + 1
        Loop
        xSql = "UPDATE NumerosComprobantes SET Numero=" & xNumero & " WHERE Tipo='" & xTipo & "'"
        Conn.Execute (xSql)
        Close #1
        xSql = "UPDATE Bascula SET Procesado = 1 WHERE Procesado = 0"
        Conn.Execute (xSql)
        MsgBox "Comprobante Finalizado"
    Else
        MsgBox "NO hay Datos Para Mostrar"
    End If
    xR.Close
    xC.Close
End Select

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error en Creación de Comprobante," & vbCrLf & Err.Description
    MsgBox MSG, , "Comprobantes"
    Err.Clear
    Resume Next
End If

End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    xSql = "Select * From NumerosComprobantes"
    Set xR = Conn.Execute(xSql)
    
    If Not xR.EOF Then Combo1.text = xR!Tipo & " " & xR!Descripcion
    
    Do While Not xR.EOF
        Combo1.AddItem xR!Tipo & " " & xR!Descripcion
        xR.MoveNext
    Loop
    xR.Close
    Combo1.ListIndex = 0
    xFecIni = CDate("01/" + Format(Month(Now), "00") + "/" + Format(Year(Now)))
    xFecFin = Now
End Sub
