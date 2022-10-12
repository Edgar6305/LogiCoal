VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{D18BBD1F-82BB-4385-BED3-E9D31A3E361E}#1.0#0"; "KewlButtonz.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{886257EB-E47C-11D3-8ED1-95743DE02879}#2.0#0"; "MBSplit.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Movimientos 
   Caption         =   "Cotejar Movimientos"
   ClientHeight    =   7770
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11505
   LinkTopic       =   "Form1"
   ScaleHeight     =   7770
   ScaleWidth      =   11505
   StartUpPosition =   1  'CenterOwner
   Begin MBSplit.Splitter VSplitter 
      Height          =   6315
      Left            =   60
      TabIndex        =   0
      Top             =   780
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   11139
      SplitterPos     =   35
      Begin VB.Frame Frame1 
         Height          =   6015
         Left            =   4260
         TabIndex        =   2
         Top             =   120
         Width           =   6975
         Begin VB.TextBox MemVar_3 
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
            Left            =   1800
            MaxLength       =   150
            TabIndex        =   18
            Top             =   1380
            Width           =   4740
         End
         Begin VB.TextBox MemVar_2 
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
            Left            =   1800
            MaxLength       =   150
            TabIndex        =   15
            Top             =   1020
            Width           =   4740
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
            Left            =   1800
            TabIndex        =   13
            Top             =   1920
            Width           =   5055
         End
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
            Left            =   1800
            TabIndex        =   12
            Top             =   2460
            Width           =   5055
         End
         Begin VB.TextBox MemVar_1 
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
            Left            =   1800
            MaxLength       =   15
            TabIndex        =   3
            Top             =   600
            Width           =   1080
         End
         Begin KewlButtonz.KewlButtons PortaPapeles 
            Height          =   555
            Left            =   5280
            TabIndex        =   14
            Top             =   5100
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
            MICON           =   "Movimientos.frx":0000
            PICN            =   "Movimientos.frx":001C
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
            Left            =   1800
            TabIndex        =   17
            Top             =   5100
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
            MICON           =   "Movimientos.frx":0956
            PICN            =   "Movimientos.frx":0972
            UMCOL           =   -1  'True
            SOFT            =   0   'False
            PICPOS          =   0
            NGREY           =   0   'False
            FX              =   0
            HAND            =   0   'False
            CHECK           =   0   'False
            VALUE           =   0   'False
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Sentencia"
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
            Left            =   420
            TabIndex        =   19
            Top             =   1440
            Width           =   750
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "DirName"
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
            Left            =   420
            TabIndex        =   16
            Top             =   1020
            Width           =   660
         End
         Begin VB.Label LabelCliente 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Segoe UI"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   3000
            TabIndex        =   11
            Top             =   600
            Width           =   3555
         End
         Begin VB.Label Label11 
            Caption         =   "Archivo"
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
            Left            =   420
            TabIndex        =   5
            Top             =   1980
            Width           =   1035
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
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
            Left            =   420
            TabIndex        =   4
            Top             =   660
            Width           =   540
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
            NumListImages   =   19
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":0F0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":1066
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":11C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":131A
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":1474
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":15CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":1728
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":1882
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":19DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":1B36
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":1C90
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":1DEA
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":27FC
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":2D96
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":3330
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":38CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":4414
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":54E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Movimientos.frx":663C
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Height          =   750
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   19815
      _ExtentX        =   34951
      _ExtentY        =   1323
      BandCount       =   4
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
      MinHeight4      =   315
      Width4          =   1695
      NewRow4         =   0   'False
      Begin KewlButtonz.KewlButtons oNuevo 
         Height          =   315
         Left            =   3405
         TabIndex        =   9
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
         MICON           =   "Movimientos.frx":6C52
         PICN            =   "Movimientos.frx":6C6E
         PICH            =   "Movimientos.frx":7208
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
         TabIndex        =   8
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
         ItemData        =   "Movimientos.frx":77A2
         Left            =   825
         List            =   "Movimientos.frx":77B2
         TabIndex        =   7
         Text            =   "TODAS"
         Top             =   390
         Width           =   2355
      End
   End
   Begin ComctlLib.StatusBar oBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   10
      Top             =   7455
      Width           =   11505
      _ExtentX        =   20294
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   2
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
Attribute VB_Name = "Movimientos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Mov As New ADODB.Recordset
Dim okUnload As Boolean
Dim IsNewRecord As Boolean
Dim OkOpen As Boolean
Dim oT As New ADODB.Recordset
Dim xT As New ADODB.Recordset
Dim sArchivo As String

Private Sub Command1_Click()
Dim xSql As String, Msj As String, Res As String
Dim xRec As New ADODB.Recordset
Dim i As Integer, j As Integer, Conter As Integer

On Error GoTo Recover

List1.Clear
If Not Conn.Execute("Select * From ArchivosMovimientos Where Nombre='" & Combo1.text & "'").EOF Then
    MsgBox "Archivo YA Procesado, Verifique"
    Exit Sub
End If
List1.AddItem "Espere mientras se cargan los archivos desde excell..."
List1.Refresh
i = 1
xSql = "SELECT * FROM OPENROWSET('Microsoft.ACE.OLEDB.16.0','Excel 12.0;HDR=NO;Database=" & MemVar_2 & Combo1 & "','" & MemVar_3 & "')"
Set xRec = Conn.Execute(xSql)

Do While Not xRec.EOF
    If i = 1 Then List1.Clear
    If Not IsNumeric(xRec!F2) Then
        List1.AddItem "Linea " & Format(i, "000") & " El Tiquete NO es un numero " & xRec!F20
    End If
    If Not IsDate(xRec!F16) Then
        List1.AddItem "Linea " & Format(i, "000") & " La Fecha NO es un Date " & xRec!F20
    End If
    If Not IsDate(xRec!F17) Then
        List1.AddItem "Linea " & Format(i, "000") & " La Fecha NO es un Date " & xRec!F20
    End If
    If Not IsNumeric(xRec!F20) Then
        List1.AddItem "Linea " & Format(i, "000") & " El peso NO es un numero " & xRec!F20
    End If
    xRec.MoveNext
    i = i + 1
Loop

If List1.ListCount > 0 Then
    MsgBox "La Revisión Finalizó Con Errores, favor revisar"
    Exit Sub
End If

If MsgBox("Esta Seguro de Subir los registros a la Base de Datos de Bascula ", vbYesNo, "Cargue Masivo de Datos") = vbYes Then
    Command1.Caption = "Ejecutar"
    xRec.MoveFirst
    xRec.MoveNext
    i = 1
    List1.Clear
    Do While Not xRec.EOF
        xSql = "SET DATEFORMAT DMY EXEC PA_CargueMovimientosDetalles " & MemVar_1.Tag & "," & xRec!F2 & ",'" & Format(xRec!F16, "dd/MM/yyyy hh:mm") & "','" & Format(xRec!F17, "dd/MM/yyyy hh:mm") & "'," & xRec!F20
        Conn.Execute (xSql)
        List1.AddItem "Linea " & Format(i, "000") & " Tiquete " & Format(xRec!F2, "######") & " Fecha " & xRec!F16 & " " & xRec!F17 & " Peso Neto " & xRec!F20
        xRec.MoveNext
        i = i + 1
    Loop
    
    Conn.Execute ("SET DATEFORMAT DMY INSERT INTO ArchivosMovimientos VALUES('" & Combo1 & "','" & Susuario & "','" & Format(Now, "dd/MM/yyyy hh:mm") & "')")
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
On Error Resume Next
    oT.Close
    xT.Close
    MenuNuevo.Flag_xProg = ""
    Unload Me
End Sub

Private Sub Form_Load()

xSql = "Select Top 1 * From Movimientos " & xFilter
oT.Open xSql, Conn, 2, 3, 1

xSql = "Select * From Terceros"
xT.Open xSql, Conn, 2, 3, 1

Set VSplitter.LeftOrTopCtl = oTree
Set VSplitter.RightOrBottomCtl = Frame1

End Sub
Private Sub MemVar_1_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
                Call omenu("Browse")
        Case vbKeyUp
        Case vbKeyDown, vbKeyReturn
                Combo1.SetFocus
End Select

End Sub

Private Sub MemVar_1_GotFocus()
        Call Mark(MemVar_1)
End Sub

Private Sub MemVar_1_LostFocus()
Dim xSql As String

On Error GoTo Recover

MemVar_1 = UCase(MemVar_1)
If Not IsNewRecord Then
     If Not MySeek(xT, Conn, "Terceros", "Identificacion='" & MemVar_1 & "'") Then
            LabelCliente.Caption = xT!Descripcion
            If Not MySeek(oT, Conn, "Movimientos", "Cliente='" & MemVar_1 & "'") Then
                MemVar_1.Tag = oT!IdMovimiento
                MemVar_2 = oT!Dirname
                MemVar_3 = oT!Sentencia
                Combo1.Clear
                sArchivo = Dir(oT!Dirname & "*.xlsx")
                Combo1.text = sArchivo
                Do While sArchivo <> ""
                    Combo1.AddItem sArchivo
                    sArchivo = Dir
                Loop
            Else
                MsgBox "El Cliente NO Registra Parametrización de Movimientos"
                MemVar_1.SetFocus
            End If
     Else
            MsgBox "Cliente NO Registrado, Verifique"
            MemVar_1.SetFocus
    End If
End If

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al  Carga el Cliente" & vbCrLf & Err.Description
    MsgBox MSG, , "MemVar_1_LostFocus()"
    Err.Clear
    Resume Next
End If

End Sub

Private Sub MemVar_2_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
        Case vbKeyF4
        Case vbKeyUp
                MemVar_1.SetFocus
        Case vbKeyDown, vbKeyReturn
                Combo1.SetFocus
End Select
End Sub

Private Sub oNuevo_Click()
If oNuevo.Caption = "Nuevo" Then
    oNuevo.Caption = "Cancelar"
    IsNewRecord = True
    MemVar_1 = ""
    MemVar_2 = ""
    MemVar_3 = ""
    Fecha = Now
   
    oBar.Panels("Usuario").text = "USUARIO: " & Susuario
    oBar.Panels("Ot").text = "FECHA CREACION: " & Now
    
Else
    oNuevo.Caption = "Nuevo"
    IsNewRecord = False
End If
MemVar_1.SetFocus
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
                    BrowseCatalogo.xtabla = "Select Identificacion, Descripcion From Terceros"
                    Set BrowseCatalogo.dControl = ActiveControl
                    BrowseCatalogo.Show 1
            End Select
    Case "Imprime"
    
    Case "Borrar"
    
    Case "Top"
        oT.Close
        oT.Open "Select Top 1 * From Movimientos Order By Cliente"
        Call LoadData
    Case "Bottom"
        oT.Close
        oT.Open "Select Top 1 * From Movimientos Order By Cliente DESC"
        Call LoadData
    Case "Proximo"
        oT.Close
        oT.Open "Select Top 1 * From Movimientos Where Cliente>'" & MemVar_1 & "' Order By  Cliente"
        Call LoadData
    Case "Previo"
        oT.Close
        oT.Open "Select Top 1 * From Movimientos Where Cliente<'" & MemVar_1 & "' Order By Cliente DESC"
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
     MemVar_1 = oT!Cliente
     MemVar_2 = oT!Dirname
     MemVar_3 = oT!Sentencia
'     MemVar_4 = oT!Sentencia
'     MemVar_5 = oT!Observaciones
     If Not MySeek(xT, Conn, "Terceros", "Identificacion='" & MemVar_1 & "'") Then
            LabelCliente.Caption = xT!Descripcion
     Else
            MsgBox "Cliente NO Registrado, Verifique"
            MemVar_1.SetFocus
            Exit Sub
    End If
 Else
 
 End If

 oBar.Panels("Usuario").text = "USUARIO: " & oT!Usuario
 oBar.Panels("Ot").text = "FECHA CREACION: " & oT!Fecha

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
Dim ok As Boolean

'On Error GoTo Recover

        ok = False
        If MemVar_1.text = "" Then ok = True
        If MemVar_2.text = "" Then ok = True
        If ok Then Exit Sub
        
        If IsNewRecord Then
            oT.AddNew
            oT!Cliente = MemVar_1
            oT!Dirname = MemVar_2
            oT!SheetName = MemVar_3
            oT!Sentencia = MemVar_4
            oT!Observaciones = MemVar_5
            oT!Fecha = Now
            oT!Usuario = Susuario
        Else
            oT!Cliente = MemVar_1
            oT!Dirname = MemVar_2
            oT!SheetName = MemVar_3
            oT!Sentencia = MemVar_4
            oT!Observaciones = MemVar_5
            oT!Fecha = Now
            oT!Usuario = Susuario
        End If
        
        oT.Update
        
        If IsNewRecord Then
            If Not MySeek(oT, Conn, "Usuarios_T", "Login=" & MemVar_1) Then
                     Call LoadData
            End If
        End If
        IsNewRecord = False
        oNuevo.Caption = "Nuevo"

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
Set Nodx = oTree.Nodes.Add(, , "0" & "Clientes", "Clientes", 13)

xSql = "SELECT DISTINCT Descripcion, Cliente FROM vMovimientoTercero"
Set Em = Conn.Execute(xSql)

While Not Em.EOF
    Set Nodx = oTree.Nodes.Add("0" & "Clientes", tvwChild, "A" & Em!Cliente, Em!Descripcion, 14)
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

Private Sub oTree_NodeClick(ByVal Node As MSComctlLib.Node)
Dim i As Integer

Select Case Mid(Node.Key, 1, 1)
Case "A"
    MemVar_1 = Mid(Node.Key, 2, 10)
    Call MemVar_1_LostFocus
Case "B"
End Select

End Sub

Private Sub Mark(dControl As Control)
If TypeOf ActiveControl Is TextBox Or TypeOf ActiveControl Is MaskEdBox Then
                dControl.SelStart = 0
                dControl.SelLength = Len(dControl.text)
End If
End Sub

Private Sub PortaPapeles_Click()
Dim i As Integer, c As String
  For i = 0 To List1.ListCount - 1
    c = c & List1.List(i) & vbCrLf
  Next
  Clipboard.Clear
  Clipboard.SetText c
End Sub



