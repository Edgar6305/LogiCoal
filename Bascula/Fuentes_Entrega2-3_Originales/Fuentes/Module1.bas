Attribute VB_Name = "Module1"
Option Explicit
Const LOCALE_SDECIMAL = &HE
Declare Function GetUserDefaultLangID Lib "kernel32" () As Integer
Declare Function GetUserDefaultLCID Lib "kernel32" () As Long
Declare Function GetLocaleInfo Lib "kernel32" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" _
    (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, _
    ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, _
    lpMaximumComponentLength As Long, lpFileSystemFlags As Long, _
    ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long

Global Conn As New ADODB.Connection
Global Susuario As String
Global SPerfilUsuario As String
Public sUsuarioFavoritos As Boolean
Public SmaxDataFavoritos As Integer
Public SmaxViewFavoritos As Integer
Public sDataReportPath  As String
Public sPuertoBascula1 As Integer
Public sPuertoBascula2 As Integer
Public sPuertoBascula4 As Integer
Public sPuertoBascula5 As Integer
Public sProgSISMA As String
Global Nserie As Long, SLastTime As Integer, SInterval As Integer
Global sTokenSisma As String, sAplicacionSisma As String, SNitEmpresaSismaG As String
Global sComprobantes As String

Sub Main()

Dim L As Integer, xA As String, xFile As String, xRp As String, xDb  As String
Dim Smaster As String, SSerieDisco As String, MSG  As String, xDs As String
Dim Unidad As String
Dim sistemaArchivos As String
Dim volumen As String
Dim retorno As Long
Dim xP As New ADODB.Recordset

Susuario = String(256, " ")
retorno = GetVolumeInformation(Unidad, volumen, Len(volumen), Nserie, 0, 0, sistemaArchivos, Len(sistemaArchivos))
L = GetUserName(Susuario, 256)
sTokenSisma = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJ1c2VyIjp7ImVtcHJlc2EiOiI4MDIwMjI2MjIiLCJhcGxpY2FjaW9uIjoiMDQtMDAxIiwic2VyaWUiOiIxNjM1ODg5MDgifSwiaWF0IjoxNjU1NzU0NDY5fQ.qqF4kbd-hAjYP5Nsr1H74D_XC6178VzU8n6cP0rODXA"
SLastTime = 10  ' Conteo de segundos
SInterval = 10      ' Inervalo de aviso de pago de factura
sAplicacionSisma = "04-001"

    On Error GoTo Recover
    
    'Extrae el Nombre del Servidor
    xFile = App.Path & "\LOGICOAL.TXT"
    Open xFile For Input As #1
    Do While Not EOF(1)
       Input #1, xA
       Select Case xA
       Case "[DATABASE]"
            Input #1, xDb
       Case "[REPORTE]"
            Input #1, xRp
       Case "[SERIE]"
            Input #1, SSerieDisco
       Case "[COMPROBANTES]"
            Input #1, sComprobantes
       Case "[MASTER]"
            Input #1, Smaster
       End Select
    Loop
    Close #1
    
     Conn.Provider = "SQLOLEDB"
    'Conn.Properties("Integrated Security") = SSPI
    Conn.Properties("Data Source") = xDb
    Conn.Properties("Initial Catalog") = "TRACER"
    Conn.Properties("user ID") = "sa"
    Conn.Properties("password") = Smaster
    Conn.CursorLocation = adUseServer
    Conn.CommandTimeout = 0
    Conn.Open

    sDataReportPath = xRp
    
    If Now() < #6/18/2022 1:00:00 PM# Then
       MsgBox "Serie del disco " + Str(Nserie)
    End If
    
    If xDs = "SI" Then
       MsgBox "SISMA esta Detenido, favor ponerse en contacto con Sistemas", vbInformation
       Exit Sub
    ElseIf Separador() <> "." Then
       MsgBox "El Separador de Decimales debe ser Punto(.) favor Ponerse en contato con Sistemas", vbInformation
        Exit Sub
    End If
    
    Set xP = Conn.Execute("Select * From Parametros")
        If xP.EOF Then
           MsgBox "NO se Localizo el Archivo de Parametrización, se detendra la Aplicación"
           End
        End If
        SNitEmpresaSismaG = xP!Nit
        sPuertoBascula1 = xP!REPesoLleno
        sPuertoBascula2 = xP!REPesoLleno
        sPuertoBascula4 = xP!REPesoLleno
        sPuertoBascula5 = xP!REPesoLleno
    xP.Close
    PassWord.DataBase = "Data Base " & xDb
    Menu.Show

Exit Sub
Recover:
If Err.Number <> 0 Then
    MSG = "Se produjo un error al Conectar Base de Datos," & vbCrLf & Err.Description
    MsgBox MSG, , "Sub Main()"
    Err.Clear   ' Borra campos del objeto Err
    End
End If

End Sub

Public Function Separador() As String
Dim Buffer As String, ret As Long
Buffer = String(255, " ")
ret = GetLocaleInfo(GetUserDefaultLCID, LOCALE_SDECIMAL, Buffer, 255)
Separador = Trim$(Replace$(Buffer, Chr(0), ""))
End Function

Public Function MySeek(xO As Object, xC As Object, xFile, xWhere As String) As Boolean
Dim xSql As String
On Error Resume Next

xO.Close
xO.Open "Select * From " & xFile & " Where " & xWhere, xC, 2, 3, 1

If xO.EOF Then
   xO.Close
   xO.Open "Select top 1 * From " & xFile, xC, 2, 3, 1
   MySeek = True
Else
   MySeek = False
End If
End Function

Public Sub BorraRpt(rpt As CrystalReport, N As Integer)
Dim i As Integer
For i = 0 To N
    rpt.Formulas(i) = ""
Next i
End Sub

Public Sub ExportaExcel(xFile As String)

Dim ObjExcel As Object
Dim Libro As Object
Dim Hoja As Object
Dim i As Long
Dim j As Long
Dim xSql As String
Dim xRec As New ADODB.Recordset
Dim iCol As Integer


On Error GoTo Error_Handler
Screen.MousePointer = vbHourglass

Set xRec = Conn.Execute("Select * From " & xFile)

If xRec.EOF Then
    MsgBox "No hay datos para exportar a excel": Exit Sub
Else
    Set Libro = Nothing
    Set ObjExcel = CreateObject("Excel.Application")
    'ObjExcel.DisplayAlerts = True
    'Creamos un nuevo libro
    Set Libro = ObjExcel.Workbooks.Add
    Set Hoja = Libro.Sheets(1)
    Libro.Sheets(1).Select
    Libro.Sheets(1).Name = "SISMA"
    
    j = 1
    i = 0
    Do While Not xRec.EOF
       If i = 0 Then
            For iCol = 1 To xRec.Fields.Count
                Hoja.cells(j, iCol) = xRec.Fields(iCol - 1).Name
            Next
            i = 1
       Else
            For iCol = 1 To xRec.Fields.Count
                Hoja.cells(j + 1, iCol) = xRec.Fields(iCol - 1).Value
            Next
            j = j + 1
            xRec.MoveNext
       End If
    Loop
      
    'Excel.Visible = True
    ObjExcel.Visible = True
    
    With Hoja
        .rows(1).Font.Bold = True
        .rows(1).Font.Color = vbRed
        .Columns("A:Z").AutoFit
    End With
End If

Set Hoja = Nothing
Set Libro = Nothing
Set ObjExcel = Nothing
  
Screen.MousePointer = vbDefault
    
Exit Sub

Error_Handler:

  MsgBox Err.Description, vbCritical
  On Error Resume Next

  Set Hoja = Nothing
  Set Libro = Nothing
  Set ObjExcel = Nothing
  Screen.MousePointer = vbDefault
End Sub

Public Function file(xFile As String) As Boolean
Dim i As Integer, xRes As Boolean, xT As New ADODB.Recordset, xSql As String
xRes = False
xSql = "SELECT DISTINCT db_name() AS TABLE_CATALOG, user_name(o.uid) AS TABLE_SCHEMA, o.name AS TABLE_NAME,CASE o.xtype WHEN 'U' THEN 'BASE TABLE' WHEN 'V' THEN 'VIEW' END As TABLE_TYPE"
xSql = xSql + " FROM sysobjects o"
xSql = xSql + " WHERE o.xtype IN ('U', 'V') AND permissions(o.id) != 0 And o.name='" & xFile & "'"
xT.Open xSql, Conn, 2, 3, 1
file = Not xT.EOF
xT.Close
End Function

Public Function F_Anulado(xTran, xNum As Integer) As String
Dim oT As New ADODB.Recordset

oT.Open "Anulaciones", Conn
If Not MySeek(oT, Conn, "Anulaciones", "Transaccion=" & xTran & " And Numerodoc=" & xNum) Then
   F_Anulado = RTrim(oT!Observacion) & vbCrLf & oT!Usuario & oT!Fecha
Else
   F_Anulado = "Sin Observaciones"
End If
oT.Close
End Function

Function Enc_Guardar(P_Nombre As String)
Dim Tlong, T_Tot, PDer, PIzq, Acum, Tt
Acum = ""
If IsEmpty(P_Nombre) Then
     P_Nombre = ""
Else
    P_Nombre = Trim(P_Nombre)
    T_Tot = Len(P_Nombre)
    For Tt = 1 To T_Tot
         Acum = Acum + Chr(Asc(Mid(P_Nombre, Tt, 1)) - 11)
    Next Tt
    P_Nombre = Acum
    If T_Tot > 1 Then
         P_Nombre = Trim(Chr(Asc(Trim(Mid(P_Nombre, 1, 1))) - T_Tot)) + _
                    Trim(Mid(P_Nombre, 2, T_Tot - 2)) + _
                    Trim(Chr(Asc(Trim(Mid(P_Nombre, T_Tot, 1))) + T_Tot))
    Else
         P_Nombre = Trim(Chr(Asc(Trim(Mid(P_Nombre, 1, 1))) - T_Tot))
    End If
    'P_Nombre = STRTRAN(P_Nombre, "'", "|")
End If
Enc_Guardar = P_Nombre
End Function

Function Enc_Leer(P_Encrip As String)
    Dim Acum, T_Tot, Tlong, PDer, PIzq, Tt
    Acum = ""
    
    'USER:ARODRIGUEZ
    'DATE:2010.06.02
    'OBJECT:Cuando es Nulo se debe salir
    If P_Encrip = "" Then Exit Function
    
    'P_Encrip = STRTRAN(P_Encrip, "|", "'")
    If IsEmpty(P_Encrip) Then
         P_Encrip = ""
    Else
    P_Encrip = Trim(P_Encrip)
    T_Tot = Len(P_Encrip)
    For Tt = 1 To T_Tot
         Acum = Acum + Chr(Asc(Mid(P_Encrip, Tt, 1)) + 11)
    Next Tt
    P_Encrip = Acum
    Tlong = Len(P_Encrip)
    If T_Tot > 1 Then
         PDer = Trim(Chr(Asc(Trim(Mid(P_Encrip, 1, 1))) + T_Tot)) _
                + Trim(Mid(P_Encrip, 2, T_Tot - 2)) _
                + Trim(Chr(Asc(Trim(Mid(P_Encrip, T_Tot, 1))) - T_Tot))
    Else
         PDer = Trim(Chr(Asc(Trim(Mid(P_Encrip, 1, 1))) + T_Tot))
    End If
    P_Encrip = PDer
    End If
    Enc_Leer = P_Encrip
End Function

Public Function Seguridad(xTool As Control, xF As String, xT As Integer) As Boolean
Dim oT As New ADODB.Recordset
Dim xR As Boolean

oT.Open "Perfiles", Conn, 2, 3, 512

If Not MySeek(oT, Conn, "Perfiles", "Usuario='" & SPerfilUsuario & "' And Programa='" & xF & "'") Then
    Select Case xT
    Case 2
       If oT!total Then
          xR = True
       Else
          xTool.buttons("Grabar").Enabled = False
          xTool.buttons("Borrar").Enabled = False
          xR = False
       End If
    Case 1
       If oT!Acceso Then
          xTool.Enabled = True
          xR = True
       Else
          xTool.Enabled = False
          xR = False
       End If
    End Select
Else
    Select Case xT
    Case 2
          xTool.buttons("Grabar").Enabled = False
          xTool.buttons("Borrar").Enabled = False
          xR = False
    Case 1
        xTool.Enabled = False
        xR = False
    End Select
End If
Seguridad = xR
oT.Close
End Function

