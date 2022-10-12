VERSION 5.00
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Begin VB.Form EntradaMasivaDatos 
   Caption         =   "Entrada Masiva de Datos"
   ClientHeight    =   5985
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   ScaleHeight     =   5985
   ScaleWidth      =   6135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Selección de Archivos"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4755
      Left            =   300
      TabIndex        =   2
      Top             =   300
      Width           =   5595
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Segoe UI"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2565
         Left            =   240
         TabIndex        =   4
         Top             =   1020
         Width           =   5055
      End
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
         Left            =   300
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   420
         Width           =   2715
      End
      Begin KewlButtonz.KewlButtons PortaPapeles 
         Height          =   555
         Left            =   3840
         TabIndex        =   5
         Top             =   4020
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   979
         BTYPE           =   3
         TX              =   "ClipBoard"
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
         MICON           =   "EntradaMasivaDatos.frx":0000
         PICN            =   "EntradaMasivaDatos.frx":001C
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
   Begin KewlButtonz.KewlButtons Command1 
      Height          =   435
      Left            =   300
      TabIndex        =   0
      Top             =   5340
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   767
      BTYPE           =   3
      TX              =   "Revisar"
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
      MICON           =   "EntradaMasivaDatos.frx":0956
      PICN            =   "EntradaMasivaDatos.frx":0972
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin KewlButtonz.KewlButtons Command2 
      Height          =   435
      Left            =   4140
      TabIndex        =   1
      Top             =   5340
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
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "EntradaMasivaDatos.frx":0F0C
      PICN            =   "EntradaMasivaDatos.frx":0F28
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
Attribute VB_Name = "EntradaMasivaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sArchivo As String

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Form_Load()
MsgBox "Path Masivo " & sPathMasivo
sArchivo = Dir(sPathMasivo & "*.xlsx")
Combo1.text = sArchivo
Do While sArchivo <> ""
    Combo1.AddItem sArchivo
    sArchivo = Dir
Loop
End Sub

Private Sub Command1_Click()
Dim xSql As String, Msj As String, Res As String
Dim xRec As New ADODB.Recordset
Dim xMas As New ADODB.Recordset
Dim i As Integer, j As Integer, Conter As Integer

On Error GoTo Recover

List1.Clear
If Not Conn.Execute("Select * From ArchivosMasivos Where Nombre='" & Combo1.text & "'").EOF Then
    MsgBox "Archivo YA Procesado, Verifique"
    Exit Sub
End If
List1.AddItem "Espere mientras se cargan los archivos desde excell..."
List1.Refresh
i = 1

xSql = "SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.16.0','Excel 12.0;HDR=NO;Database=" & sPathMasivo & Combo1 & "','select * from [sisma$] WHERE ISDATE(F9)')"
Set xRec = Conn.Execute(xSql)
List1.Clear
Do While Not xRec.EOF
    If Not Transaccion(xRec!F1) Then
        List1.AddItem "Linea " & Format(i, "000") & " Transaccion Invalida     " & xRec!F1
    End If
    If Origen(xRec!F1, xRec!F2) Then
        List1.AddItem "Linea " & Format(i, "000") & " Origen Invalido          " & xRec!F1 & " " & xRec!F2
    End If
    If Transportador(xRec!F4) Then
        List1.AddItem "Linea " & Format(i, "000") & " Transportador Invalido   " & xRec!F4
    End If
    If Placas(xRec!F5) Then
        List1.AddItem "Linea " & Format(i, "000") & " Placas NO Localizadas    " & xRec!F5
    End If
    If Conductor(xRec!F6) Then
        List1.AddItem "Linea " & Format(i, "000") & " Conductor Invalido       " & xRec!F6
    End If
    
    If Not IsDate(xRec!F9) Then
        List1.AddItem "Linea " & Format(i, "000") & " Fecha Llegada Invalido   " & xRec!F9
    End If
    If Not IsDate(xRec!F10) Then
        List1.AddItem "Linea " & Format(i, "000") & " Hora Llegada Invalido   " & xRec!F10
    End If
    
    If Not IsNumeric(xRec!F12) Then
        If Not (xRec!F12 = 0 Or xRec!F12 = 1) Then
            List1.AddItem "Linea " & Format(i, "000") & " Campo Carpado Invalido   " & xRec!F12
        End If
    End If
    
    If Not IsNumeric(xRec!F13) Then
        List1.AddItem "Linea " & Format(i, "000") & " IDTipo de carbon Invalido " & IIf(IsNull(xRec!F13), "Null", xRec!F13)
    End If
     
    xRec.MoveNext
    i = i + 1
Loop

List1.Refresh
If List1.ListCount > 0 Then
    MsgBox "La Revisión Finalizó Con Errores, favor revisar"
    Exit Sub
End If

If MsgBox("Esta Seguro de Subir los registros a la Base de Datos de Bascula ", vbYesNo, "Cargue Masivo de Datos") = vbYes Then
    'CREATE PROCEDURE [dbo].[PA_CargueMasivo]
    '@IdTransaccion int,
    '@TransaccionOrigen varchar(2),
    '@NumeroTransaccion int,
    '@Documentoasociado Int,
    '@IdTransportador int,
    '@Placas varchar(6),
    '@Conductor varchar(15),
    '@PesoLleno Float,
    '@PesoVacio Float,
    '@FechaLleno Datetime,
    '@FechaVacio DateTime,
    '@Usuario Varchar(15),
    '@Observaciones Varchar(125),
    '@Carpado Int
    '@IdTipoCarbon
    
    Command1.Caption = "Ejecutar"
    
    xRec.MoveFirst
    i = 1
    List1.Clear
    Do While Not xRec.EOF
        If xRec!F1 = "LT" Then
           IdTran = 1
        Else
           IdTran = 2
        End If
        xSql = "SET DATEFORMAT DMY EXEC PA_CargueMasivo " & IdTran & "," & xRec!F1 & "," & xRec!F2 & "," & xRec!F3 & "," & xRec!F4 & ",'" & xRec!F5 & "','" & xRec!F6 & "',"
        xSql = xSql & xRec!F7 & "," & xRec!F8 & ",'" & Format(xRec!F9, "dd/MM/yyyy") & "','" & Format(xRec!F10, "dd/MM/yyyy") & "','" & Susuario & "','" & xRec!F11 & "'," & xRec!F12 & "," & xRec!F13
     
        Conn.Execute (xSql)
    '    Res = xMas!Res
    '    If Res <> "OK" Then
        List1.AddItem "Linea " & Format(i, "000") & " Registro Procesado " & xRec!F1 & " Numero " & xRec!F2 & " Placas " & xRec!F5
    '    End If
        xRec.MoveNext
        i = i + 1
    Loop
    
    Conn.Execute ("SET DATEFORMAT DMY INSERT INTO ArchivosMasivos VALUES('" & Combo1 & "','" & Susuario & "','" & Format(Now, "dd/MM/yyyy hh:mm") & "')")
    MsgBox "Cargue Finalizado Correctamente"
    Command1.Caption = "Revisar"
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error en Generación de Datos de Excell a SQL," & vbCrLf & Err.Description
    MsgBox MSG, , "Cargue de Datos"
    Err.Clear   ' Borra campos del objeto Err
    Exit Sub
End If
End Sub

Private Function Conductor(Id As String) As Boolean
Dim Res As Boolean
    Res = Conn.Execute("Select * From Conductores Where Cedula='" & Id & "'").EOF
    Conductor = Res
End Function

Private Function Placas(Id As String) As Boolean
Dim Res As Boolean
    Res = Conn.Execute("Select * From Placas Where Placas='" & Id & "'").EOF
    Placas = Res
End Function

Private Function Transportador(Id As Integer) As Boolean
Dim Res As Boolean
    Res = Conn.Execute("Select * From Transportador Where IdTransportador=" & Id).EOF
    Transportador = Res
End Function

Private Function Origen(Tran As String, Numero As Integer) As Boolean
Dim Res As Boolean

Select Case Tran
Case "LT"
    Res = Conn.Execute("Select * From Lotes Where IdLote=" & Numero).EOF
Case "DS"
    Res = Conn.Execute("Select * From Ventas Where IdVentas=" & Numero).EOF
Case "TR"
    Res = Conn.Execute("Select * From Traslados Where IdTraslado=" & Numero).EOF
End Select

Origen = Res
End Function

Private Function Transaccion(Tran As String) As Boolean
Dim Res As Boolean

If Tran = "LT" Or Tran = "DS" Or Tran = "TR" Then
    Res = True
Else
    Res = False
End If

Transaccion = Res
End Function

Private Sub PortaPapeles_Click()
Dim i As Integer, c As String
  For i = 0 To List1.ListCount - 1
    c = c & List1.List(i) & vbCrLf
  Next
  Clipboard.Clear
  Clipboard.SetText c
  MsgBox "Información Copiada en el ClipBoard"
End Sub
