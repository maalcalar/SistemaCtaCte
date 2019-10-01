Attribute VB_Name = "Modulo1"
Option Explicit
Public fMainform As frmMenu
Public FileReport As String
Public filebloc As String
Public apli As Long

Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long


Public flogin As frmLogin
Public fimpresora2 As frmImpresora2
Public Db As Connection, Db2 As Connection
Public DbMaster As Connection
Public DbCb As Connection
Public DbOld As Connection
Public DbEvento As Connection

Public ADO1  As Recordset
Public ADO2  As Recordset
Public ADO3  As Recordset
Public ADO4  As Recordset
Public ADO5  As Recordset
Public ADO6  As Recordset
Public ADO7  As Recordset
Public ADO8  As Recordset
Public ADO3a  As Recordset
Public ADO4a  As Recordset
Public ADO5a  As Recordset
Public ADO6a  As Recordset
Public ADO7a  As Recordset
Public ADO8a  As Recordset
Public ADOx  As Recordset

Public ADOCb1 As Recordset
Public ADOCb2 As Recordset
Public ADOCb3 As Recordset

Public ADOMaster As Recordset
Public ADOMaster2 As Recordset
Public ADOMaster3 As Recordset

Public ADOOld1 As Recordset
Public ADOOld2 As Recordset
Public ADOOld3 As Recordset

Public ADOEvento1 As Recordset
Public ADOEvento2 As Recordset
Public ADOEvento3 As Recordset

Public xlista As String, _
       xseleccion As String, _
       xselecSocio As Integer, _
       xselecIns As Integer, _
       xselecCodofin As Long, _
       xseleccion2 As String, xseleTecla As String, xseleTipo As String
Public wnomcia As String
Public wcodcia As String * 2
Public wruccia As String * 11
Public walmcia As String * 4
Public wnomalm As String

Public wdircia As String
Public wtelcia As String
Public wnomlog As String

Public wdiacia As String
Public wmescia As String
Public wanocia As String
Public wnomusu As String
Public wcodusu As String

Public wrescia As String

Public SwAbrirBd As Boolean
Public SwAbrirCb As Boolean
Public SwAbrirMaster As Boolean
Public SwAbrirOld As Boolean
Public SwAbrirEvento As Boolean

Public cadMaster As String
Public cadInvent As String
Public cadCb As String
Public cadOld As String
Public cadEvento As String

Public xraiz As String
Public xraizDIECO As String
Public xraizCAJAMP As String
Public xraizBCP As String
Public DataBase As String
Public DataBase2 As String
Public xUser As String
Public xPwd As String
Public xUser2 As String
Public xPwd2 As String

Public ruta As String
Public rutaMaster As String
Public rutaCb As String
Public rutaOld As String
Public rutaEvento As String

Public arriba As String
Public wprinter As String * 1
'
Public MENUMAE As Boolean, MENUELE As Boolean, MENUGES As Boolean, MENUAPO As Boolean
Public MENUDIE As Boolean, MENUCAJ As Boolean, MENUTES As Boolean, MENUCON As Boolean, MENUSER As Boolean
Public MENUBCP As Boolean, MENURPT As Boolean, MENUCNT As Boolean
Public wporigv As Single

Public wswprint As String * 1
Public tipoprint As String * 1
Public todoprint As Boolean
Public desdeprint As Integer, hastaprint As Integer

Public wnewane As String
Public wnewart As String

Public V As Long, H As Integer
Public objExcel As Excel.Application

Public wayucliente As String

Public wayuclidoc As String * 11
Public wayutipdoc As String * 2
Public wayuserdoc As String * 4
Public wayunumdoc As String * 9
Public wayufecha As Date
Public wayuvcmto As Date
Public wayumoneda As String * 1
Public wayutotvta As Currency
Public wayusdovta As Currency
Public wayucanje As String * 6

Public xodbc As String
Public zMesTope As String
Public zFamSocio As Integer
Public zFamParie As String
Public zLinParie As String

Public zSocio As String

Public zDetaCambio As Boolean
Public zDetaCodSoc As Integer
Public zDetaTipDsc As String
Public zDetaAnoDsc As String
Public zDetaMesDsc As String
Public zDetaSw As Boolean

Public zSerCaj As String
Public zMonCaj As String

Public SUPERVISOR As Boolean

Public Sub CreaCn(zStoreProc As String, _
                  zUsu As String, zCia As String, zSub As String, _
                  zTmo As String, zSer As String, zGui As String, zCos As String)
   Dim cn As New ADODB.Connection
   Dim cmd As New ADODB.Command
   Dim prm As ADODB.Parameter
'   cn.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=" + ruta + ";Data Source=" + DataBase + ""
   
   cn.ConnectionString = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" + xUser + ";pwd=" + xPwd + ";Data Source=" + DataBase + ";Initial Catalog=" + ruta + ""
   cn.Open
   cmd.ActiveConnection = cn
   cmd.CommandType = adCmdStoredProc
   cmd.CommandTimeout = 720
   
   Set prm = cmd.CreateParameter("wUsu", adChar, adParamInput, 3, zUsu)
   cmd.Parameters.Append prm
   
   Set prm = cmd.CreateParameter("wCia", adChar, adParamInput, 2, zCia)
   cmd.Parameters.Append prm
   
   Set prm = cmd.CreateParameter("wSub", adChar, adParamInput, 1, zSub)
   cmd.Parameters.Append prm
   
   Set prm = cmd.CreateParameter("wTmo", adChar, adParamInput, 2, zTmo)
   cmd.Parameters.Append prm
   
   Set prm = cmd.CreateParameter("wSer", adChar, adParamInput, 4, zSer)
   cmd.Parameters.Append prm
   
   Set prm = cmd.CreateParameter("wGui", adChar, adParamInput, 9, zGui)
   cmd.Parameters.Append prm
   
   Set prm = cmd.CreateParameter("wCos", adChar, adParamInput, 1, zCos)
   cmd.Parameters.Append prm
   
   cmd.CommandText = zStoreProc
   cmd.Execute
      
   cmd.Parameters.Delete (0)
   cmd.Parameters.Delete (0)
   cmd.Parameters.Delete (0)
   cmd.Parameters.Delete (0)
   cmd.Parameters.Delete (0)
   cmd.Parameters.Delete (0)
      
   cn.Close
End Sub

Public Function Leerado(sql As String) As Long
    On Error GoTo salir
    Set ADO1 = New Recordset
    ADO1.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado = ADO1.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado = -1
End Function

Public Function Leerado2(sql As String) As Long
  On Error GoTo salir
    Set ADO2 = New Recordset
    ADO2.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado2 = ADO2.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado2 = -1
End Function

Public Function Leerado3(sql As String) As Long
  On Error GoTo salir
    Set ADO3 = New Recordset
    ADO3.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado3 = ADO3.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado3 = -1
End Function

Public Function Leerado4(sql As String) As Long
  On Error GoTo salir
    Set ADO4 = New Recordset
    ADO4.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado4 = ADO4.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado4 = -1
End Function

Public Function Leerado5(sql As String) As Long
  On Error GoTo salir
    Set ADO5 = New Recordset
    ADO5.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado5 = ADO5.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado5 = -1
End Function

Public Function Leerado6(sql As String) As Long
  On Error GoTo salir
    Set ADO6 = New Recordset
    ADO6.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado6 = ADO6.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado6 = -1
End Function

Public Function Leerado7(sql As String) As Long
  On Error GoTo salir
    Set ADO7 = New Recordset
    ADO7.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado7 = ADO7.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado7 = -1
End Function

Public Function Leerado8(sql As String) As Long
  On Error GoTo salir
    Set ADO8 = New Recordset
    ADO8.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado8 = ADO8.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado8 = -1
End Function

Public Function Leerado3a(sql As String) As Long
  On Error GoTo salir
    Set ADO3a = New Recordset
    ADO3a.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado3a = ADO3a.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado3a = -1
End Function

Public Function Leerado4a(sql As String) As Long
  On Error GoTo salir
    Set ADO4a = New Recordset
    ADO4a.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado4a = ADO4a.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado4a = -1
End Function

Public Function Leerado5a(sql As String) As Long
  On Error GoTo salir
    Set ADO5a = New Recordset
    ADO5a.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado5a = ADO5a.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado5a = -1
End Function

Public Function Leerado6a(sql As String) As Long
  On Error GoTo salir
    Set ADO6a = New Recordset
    ADO6a.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado6a = ADO6a.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado6a = -1
End Function

Public Function Leerado7a(sql As String) As Long
  On Error GoTo salir
    Set ADO7a = New Recordset
    ADO7a.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado7a = ADO7a.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado7a = -1
End Function

Public Function Leerado8a(sql As String) As Long
  On Error GoTo salir
    Set ADO8a = New Recordset
    ADO8a.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leerado8a = ADO8a.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leerado8a = -1
End Function

Public Function Leeradox(sql As String) As Long
  On Error GoTo salir
    Set ADOx = New Recordset
    ADOx.Open sql, Db, adOpenDynamic, adLockOptimistic
    Leeradox = ADOx.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   Leeradox = -1
End Function

Public Function LeeradoMaster(sql As String) As Long
  On Error GoTo salir
    Set ADOMaster = New Recordset
    ADOMaster.Open sql, DbMaster, adOpenDynamic, adLockOptimistic
    LeeradoMaster = ADOMaster.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   LeeradoMaster = -1
End Function

Public Function LeeradoMaster2(sql As String) As Long
  On Error GoTo salir
    Set ADOMaster2 = New Recordset
    ADOMaster2.Open sql, DbMaster, adOpenDynamic, adLockOptimistic
    LeeradoMaster2 = ADOMaster2.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   LeeradoMaster2 = -1
End Function

Public Function LeeradoMaster3(sql As String) As Long
  On Error GoTo salir
    Set ADOMaster3 = New Recordset
    ADOMaster3.Open sql, DbMaster, adOpenDynamic, adLockOptimistic
    LeeradoMaster3 = ADOMaster3.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   LeeradoMaster3 = -1
End Function

Public Function LeeradoCb1(sql As String) As Long
    On Error GoTo salir
    Set ADOCb1 = New Recordset
    ADOCb1.Open sql, DbCb, adOpenDynamic, adLockOptimistic
    LeeradoCb1 = ADOCb1.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   LeeradoCb1 = -1
End Function

Public Function LeeradoCb2(sql As String) As Long
    On Error GoTo salir
    Set ADOCb2 = New Recordset
    ADOCb2.Open sql, DbCb, adOpenDynamic, adLockOptimistic
    LeeradoCb2 = ADOCb2.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   LeeradoCb2 = -1
End Function

Public Function LeeradoCb3(sql As String) As Long
    On Error GoTo salir
    Set ADOCb3 = New Recordset
    ADOCb3.Open sql, DbCb, adOpenDynamic, adLockOptimistic
    LeeradoCb3 = ADOCb3.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   LeeradoCb3 = -1
End Function

Public Function LeeradoOld1(sql As String) As Long
    On Error GoTo salir
    Set ADOOld1 = New Recordset
    ADOOld1.Open sql, DbOld, adOpenDynamic, adLockOptimistic
    LeeradoOld1 = ADOOld1.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   LeeradoOld1 = -1
End Function

Public Function LeeradoOld2(sql As String) As Long
    On Error GoTo salir
    Set ADOOld2 = New Recordset
    ADOOld2.Open sql, DbOld, adOpenDynamic, adLockOptimistic
    LeeradoOld2 = ADOOld2.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   LeeradoOld2 = -1
End Function

Public Function LeeradoOld3(sql As String) As Long
    On Error GoTo salir
    Set ADOOld3 = New Recordset
    ADOOld3.Open sql, DbOld, adOpenDynamic, adLockOptimistic
    LeeradoOld3 = ADOOld3.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   LeeradoOld3 = -1
End Function

Public Function LeeradoEvento1(sql As String) As Long
    On Error GoTo salir
    Set ADOEvento1 = New Recordset
    ADOEvento1.Open sql, DbEvento, adOpenDynamic, adLockOptimistic
    LeeradoEvento1 = ADOEvento1.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   LeeradoEvento1 = -1
End Function

Public Function LeeradoEvento2(sql As String) As Long
    On Error GoTo salir
    Set ADOEvento2 = New Recordset
    ADOEvento2.Open sql, DbEvento, adOpenDynamic, adLockOptimistic
    LeeradoEvento2 = ADOEvento2.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   LeeradoEvento2 = -1
End Function

Public Function LeeradoEvento3(sql As String) As Long
    On Error GoTo salir
    Set ADOEvento3 = New Recordset
    ADOEvento3.Open sql, DbEvento, adOpenDynamic, adLockOptimistic
    LeeradoEvento3 = ADOEvento3.RecordCount
    Exit Function
salir:
   MsgBox err.Description, vbCritical
   LeeradoEvento3 = -1
End Function

Sub Main()
   wdircia = "CENTRO COMERCIAL MEGA PLAZA"
   wtelcia = "999-8877"
   wrescia = "CARLOS COLOMA OBREGON"
   wnomlog = "EBER MENA"
   
   Dim micadena As String, mipos As String, auxcadena As String
    'configura
   Open App.Path + "\CTACTE.ini" For Input As #1 ' Abre el archivo para recibir los datos.
    ' Repite el bucle hasta el final del archivo.
   Do While Not EOF(1)
      Line Input #1, micadena
      mipos = InStr(4, micadena, "=", 1)
      auxcadena = Mid(micadena, 1, mipos - 1)
      If UCase(auxcadena) = "RUTARAIZ" Then
         xraiz = Mid(micadena, mipos + 1)
      End If
      If UCase(auxcadena) = "RUTADIECO" Then
         xraizDIECO = Mid(micadena, mipos + 1)
      End If
      If UCase(auxcadena) = "RUTACAJAMP" Then
         xraizCAJAMP = Mid(micadena, mipos + 1)
      End If
      If UCase(auxcadena) = "RUTABCP" Then
         xraizBCP = Mid(micadena, mipos + 1)
      End If
      If UCase(auxcadena) = "DATABASE" Then
         DataBase = Mid(micadena, mipos + 1)
      End If
      If UCase(auxcadena) = "ARRIBA" Then
          arriba = Mid(micadena, mipos + 1)
      End If
      If UCase(auxcadena) = "ODBC" Then
          xodbc = Mid(micadena, mipos + 1)
      End If
      If UCase(auxcadena) = "USER" Then
         xUser = Mid(micadena, mipos + 1)
      End If
      If UCase(auxcadena) = "PWD" Then
         xPwd = Mid(micadena, mipos + 1)
      End If
   Loop
   Close #1
    
   Call abrirBD
   Call abrirMaster
    
   Set DbMaster = New Connection
   DbMaster.CursorLocation = adUseClient
   DbMaster.Open cadMaster
   ' Se Carga formulario de Login
   Set flogin = New frmLogin
   flogin.Show vbModal
   If Not flogin.OK Then
      End
   End If
   Unload flogin
        
   Call abrirBD
   Call abrirMaster
    
   Set fMainform = New frmMenu
   fMainform.Show
End Sub

Public Sub abrirMaster()
   On Error GoTo err
   rutaMaster = "AOPIP_MASTER"
'   cadMaster = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=" + rutaMaster + ";Data Source=" + DataBase + ""
   
   cadMaster = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" + xUser + ";pwd=" + xPwd + ";Initial Catalog=" + rutaMaster + ";Data Source=" + DataBase + ""
   Set DbMaster = New Connection
   DbMaster.CursorLocation = adUseClient
   DbMaster.Open cadMaster
   SwAbrirMaster = True
   Exit Sub
err:
   MsgBox "BD Master No Existe", vbExclamation
   SwAbrirMaster = False
End Sub
    
Public Sub abrirBD()
   On Error GoTo err
   Set Db = New Connection
   Db.CursorLocation = adUseClient
   Db.ConnectionTimeout = 0
   Db.CommandTimeout = 720
   
   ruta = "AOPIP_CTACTE"
'   cadInvent = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=" + ruta + ";Data Source=" + DataBase + ""
   
   cadInvent = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" + xUser + ";pwd=" + xPwd + ";Data Source=" + DataBase + ";Initial Catalog=" + ruta + ""
   Db.Open cadInvent
   SwAbrirBd = True
   Exit Sub
err:
   MsgBox "BD CTACTE No Existe", vbExclamation
   SwAbrirBd = False
End Sub

Public Sub abrirCB()
   On Error GoTo err
   rutaCb = "AOPIP_CONTAB" + wanocia
'   cadContab = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=" + rutaCb + ";Data Source=" + DataBase + ""
   
   cadCb = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" + xUser + ";pwd=" + xPwd + ";Initial Catalog=" + rutaCb + ";Data Source=" + DataBase + ""
   Set DbCb = New Connection
   DbCb.CursorLocation = adUseClient
   DbCb.Open cadCb
   SwAbrirCb = True
   Exit Sub
err:
   MsgBox "BD CONTAB No Existe", vbExclamation
   SwAbrirCb = False
End Sub

Public Sub abrirEVENTO()
   On Error GoTo err
   rutaEvento = "AOPIP_EVENTO"
'   cadEvento = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;User ID=sa;Initial Catalog=" + rutaEvento + ";Data Source=" + DataBase + ""
   
   cadEvento = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" + xUser + ";pwd=" + xPwd + ";Initial Catalog=" + rutaEvento + ";Data Source=" + DataBase + ""
   Set DbEvento = New Connection
   DbEvento.CursorLocation = adUseClient
   DbEvento.Open cadEvento
   SwAbrirEvento = True
   Exit Sub
err:
   MsgBox "BD EVENTO No Existe", vbExclamation
   SwAbrirEvento = False
End Sub

Public Sub abrirOLD()
   On Error GoTo err
    cadOld = "DSN=cta-cte"
'   cadOld = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" + rutaOld + ";Jet OLEDB:Database;"
    
   Set DbOld = New Connection
   DbOld.CursorLocation = adUseClient
   DbOld.Open cadOld
   SwAbrirOld = True
   Exit Sub
err:
   MsgBox "Base de Datos CTA-CTE No Existe", vbExclamation
   Resume Next
End Sub

Public Function IXY(AnchoX As Double, LargoY As Double, Letra As String) As Boolean
' Rutina para imprimir En Impresora De Matriz
On Error GoTo salir
   IXY = False
   Printer.CurrentX = Printer.ScaleY(AnchoX, vbCharacters, vbTwips)
   If arriba = "+0" Or arriba = "-0" Then
      Printer.CurrentY = Printer.ScaleY(LargoY, vbCharacters, vbTwips)
   Else
      If Mid(arriba, 1, 1) = "+" Then
         Printer.CurrentY = Printer.ScaleY(LargoY + CDbl(Mid(arriba, 2)), vbCharacters, vbTwips)
      Else
         Printer.CurrentY = Printer.ScaleY(LargoY - CDbl(Mid(arriba, 2)), vbCharacters, vbTwips)
      End If
   End If
   Printer.Print Letra
   IXY = True
   Exit Function
salir:
   IXY = False
   MsgBox err.Description
End Function

Public Function funnommes(wmmm As String) As String
Select Case wmmm
Case "00"
     funnommes = " APERTURA"
Case "01"
     funnommes = "  ENERO  "
Case "02"
     funnommes = " FEBRERO "
Case "03"
     funnommes = "  MARZO  "
Case "04"
     funnommes = "  ABRIL  "
Case "05"
     funnommes = "  MAYO   "
Case "06"
     funnommes = "  JUNIO  "
Case "07"
     funnommes = "  JULIO  "
Case "08"
     funnommes = "  AGOSTO "
Case "09"
     funnommes = "SETIEMBRE"
Case "10"
     funnommes = " OCTUBRE "
Case "11"
     funnommes = "NOVIEMBRE"
Case "12"
     funnommes = "DICIEMBRE"
End Select
End Function

Public Function fundiames(wmmm As String) As String
   Select Case wmmm
   Case "01", "03", "05", "07", "08", "10", "12"
        fundiames = "31"
   Case "02"
        If wanocia = "2008" Or wanocia = "2012" Or wanocia = "2016" Or wanocia = "2020" Or wanocia = "2024" Then
           fundiames = "29"
        Else
           fundiames = "28"
        End If
   Case "04", "06", "09", "11"
        fundiames = "30"
   End Select
End Function

Public Function nompart(wnnn, nro)
    Dim lin1 As Byte, lin2 As Byte, lin3 As Byte, wape1 As String, wape2 As String, wape3 As String
    lin1 = InStr(1, wnnn, " ")
    lin2 = InStr(lin1 + 1, wnnn, " ")
    lin3 = Len(Trim(wnnn))
    If lin1 = 0 Then
       wape1 = wnnn
       wape2 = "XX"
       wape3 = "XX"
    Else
       If lin2 = lin3 Then
          wape1 = wnnn
          wape2 = "XX"
          wape3 = "XX"
       Else
          If lin2 = 0 Then
            wape1 = Mid(wnnn, 1, lin1 - 1)
            wape2 = Mid(wnnn, lin1 + 1, lin3 - lin1)
            wape3 = "XX"
          Else
             wape1 = Mid(wnnn, 1, lin1 - 1)
             wape2 = Mid(wnnn, lin1 + 1, lin2 - lin1)
             wape3 = Mid(wnnn, lin2 + 1, lin3 - lin2)
          End If
       End If
    End If
    Select Case nro
    Case 1
         nompart = Left(wape1, 20)
    Case 2
         nompart = Left(wape2, 20)
    Case 3
         nompart = Left(wape3, 20)
    End Select
End Function

Public Function NumLetras(ByVal curNumero As Double, Optional blnO_Final As Boolean = True) As String
'Devuelve un número expresado en letras.
'El parámetro blnO_Final se utiliza en la recursión para saber si se debe colocar
'la "O" final cuando la palabra es UN(O)
    Dim dblCentavos As Double
    Dim lngContDec As Long
    Dim lngContCent As Long
    Dim lngContMil As Long
    Dim lngContMillon As Long
    Dim strNumLetras As String
    Dim strNumero As Variant
    Dim strDecenas As Variant
    Dim strCentenas As Variant
    Dim blnNegativo As Boolean
    Dim blnPlural As Boolean
    
    If Int(curNumero) = 0# Then
        strNumLetras = "CERO"
    End If
    
    strNumero = Array(vbNullString, "UN", "DOS", "TRES", "CUATRO", "CINCO", "SEIS", "SIETE", _
                   "OCHO", "NUEVE", "DIEZ", "ONCE", "DOCE", "TRECE", "CATORCE", _
                   "QUINCE", "DIECISEIS", "DIECISIETE", "DIECIOCHO", "DIECINUEVE", _
                   "VEINTE")

    strDecenas = Array(vbNullString, vbNullString, "VEINTI", "TREINTA", "CUARENTA", "CINCUENTA", "SESENTA", _
                    "SETENTA", "OCHENTA", "NOVENTA", "CIEN")

    strCentenas = Array(vbNullString, "CIENTO", "DOSCIENTOS", "TRESCIENTOS", _
                     "CUATROCIENTOS", "QUINIENTOS", "SEISCIENTOS", "SETECIENTOS", _
                     "OCHOCIENTOS", "NOVECIENTOS")

    If curNumero < 0# Then
        blnNegativo = True
        curNumero = Abs(curNumero)
    End If

    If Int(curNumero) <> curNumero Then
        dblCentavos = Abs(curNumero - Int(curNumero))
        curNumero = Int(curNumero)
    End If

    Do While curNumero >= 1000000#
        lngContMillon = lngContMillon + 1
        curNumero = curNumero - 1000000#
    Loop

    Do While curNumero >= 1000#
        lngContMil = lngContMil + 1
        curNumero = curNumero - 1000#
    Loop
    
    Do While curNumero >= 100#
        lngContCent = lngContCent + 1
        curNumero = curNumero - 100#
    Loop
    
    If Not (curNumero > 10# And curNumero <= 20#) Then
        Do While curNumero >= 10#
            lngContDec = lngContDec + 1
            curNumero = curNumero - 10#
        Loop
    End If
    
    If lngContMillon > 0 Then
        If lngContMillon >= 1 Then   'si el número es >1000000 usa recursividad
            strNumLetras = NumLetras(lngContMillon, False)
            If Not blnPlural Then blnPlural = (lngContMillon > 1)
            lngContMillon = 0
        End If
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMillon) & " MILLON" & _
                                                                    IIf(blnPlural, "ES ", " ")
    End If
    
    If lngContMil > 0 Then
        If lngContMil >= 1 Then   'si el número es >100000 usa recursividad
            strNumLetras = strNumLetras & NumLetras(lngContMil, False)
            lngContMil = 0
        End If
        strNumLetras = Trim(strNumLetras) & strNumero(lngContMil) & " MIL "
    End If
    
    If lngContCent > 0 Then
        If lngContCent = 1 And lngContDec = 0 And curNumero = 0# Then
            strNumLetras = strNumLetras & "CIEN"
        Else
            strNumLetras = strNumLetras & strCentenas(lngContCent) & " "
        End If
    End If
    
    If lngContDec >= 1 Then
        If lngContDec = 1 Then
            strNumLetras = strNumLetras & strNumero(10)
        Else
            strNumLetras = strNumLetras & strDecenas(lngContDec)
        End If
        
        If lngContDec >= 3 And curNumero > 0# Then
            strNumLetras = strNumLetras & " Y "
        End If
    Else
        If curNumero >= 0# And curNumero <= 20# Then
            strNumLetras = strNumLetras & strNumero(curNumero)
            If curNumero = 1# And blnO_Final Then
                strNumLetras = strNumLetras & "O"
            End If
            If dblCentavos > 0# Then
               If blnO_Final Then
                  strNumLetras = Trim(strNumLetras) & " CON " & Format$(CInt(dblCentavos * 100#), "00") & "/100"
               Else
                  strNumLetras = strNumLetras
               End If
            Else
               If blnO_Final Then
                  strNumLetras = Trim(strNumLetras) & " CON 00/100"
               Else
                  strNumLetras = strNumLetras
               End If
            End If
            NumLetras = strNumLetras
            Exit Function
        End If
    End If
    
    If curNumero > 0# Then
        strNumLetras = strNumLetras & strNumero(curNumero)
        If curNumero = 1# And blnO_Final Then
            strNumLetras = strNumLetras & "O"
        End If
    End If
    
    If dblCentavos > 0# Then
        strNumLetras = strNumLetras & " CON " + Format$(CInt(dblCentavos * 100#), "00") & "/100"
    Else
        If blnO_Final Then
        strNumLetras = strNumLetras & " CON 00/100"
        End If
    End If
    
    NumLetras = IIf(blnNegativo, "(" & strNumLetras & ")", strNumLetras)

End Function

Public Sub printfac(zTdo As String, zSer As String, zDoc As String)
    MsgBox "Prepare la Impresora"
    Dim a As String, numreg As Long, fila As Double
    
    On Error GoTo ErrManejo
    numreg = Leerado7("SELECT * FROM VTACAB WHERE TIPDOC='" + zTdo + "' AND " _
                                            & "   SERDOC='" + zSer + "' AND " _
                                            & "   NUMDOC='" + zDoc + "' ")
    If numreg = 0 Then
       MsgBox "Venta a Imprimir No Tiene Cabecera", vbInformation
       Exit Sub
    End If
    If ADO7!anulado Then
       MsgBox "Venta a Imprimir Esta Anulado", vbExclamation
       Exit Sub
    End If
    numreg = Leerado6("SELECT * FROM VTADET WHERE TIPDOC='" + zTdo + "' AND " _
                                            & "   SERDOC='" + zSer + "' AND " _
                                            & "   NUMDOC='" + zDoc + "' ORDER BY LINEA ")
    If numreg = 0 Then
       MsgBox "Venta a Imprimir No Tiene Detalle", vbInformation
       Exit Sub
    End If
    numreg = Leerado5("SELECT * FROM MAECLIENTE WHERE CODIGO='" + ADO7!cliente + "' ")
    If numreg = 0 Then
       MsgBox "Codigo de Cliente " + ADO7!cliente + " a Imprimir No Existe", vbInformation
       Exit Sub
    End If
       
    Printer.FontName = "Draft 12cpi"
    Printer.FontBold = True
    Printer.PrintQuality = -1
    
    a = IXY(33, 8.9, "FACTURA " + zSer + "-" + zDoc)
    
    a = IXY(5.2, 10.7, Format(Day(ADO7!fecha), "00"))
    a = IXY(11.6, 10.7, funnommes(Format(Month(ADO7!fecha), "00")))
    a = IXY(24.7, 10.7, Format(Year(ADO7!fecha), "####"))
    
    a = IXY(6.8, 11.9, ADO5!nombre)
    a = IXY(18.8, 13.3, ADO7!cliente)
    
    a = IXY(6.8, 14.4, IIf(IsNull(ADO5!direc), "", Trim(ADO5!direc)))
    a = IXY(35.2, 14.4, IIf(IsNull(ADO7!condvta), "", Trim(ADO7!condvta)))
    
    a = IXY(29.8, 15.2, Left(ADO7!guiarem, 3) + "-" + Right(ADO7!guiarem, 6))
    
    Printer.FontName = "Draft 17cpi"
    Printer.FontBold = True
    Printer.PrintQuality = -1
    
    fila = 18.9
    ADO6.MoveFirst
    Do While Not ADO6.EOF
       If Left(ADO6!articulo, 3) <> "999" Then
          a = IXY(4.1, fila, IIf(ADO6!cantidad > 0, Space(9 - Len(Trim(Format(ADO6!cantidad, "####0.000")))) + Format(ADO6!cantidad, "####0.000"), Space(9)) + " " + ADO6!um)
          a = IXY(6.9, fila, ADO6!articulo)
       Else
          a = IXY(4.1, fila, IIf(ADO6!cantidad > 0, Space(5 - Len(Trim(Format(ADO6!cantidad, "####0")))) + Format(ADO6!cantidad, "####0"), Space(5)))
       End If
       a = IXY(7.1, fila, Left(ADO6!nombre, 70))
       
       a = IXY(34.5, fila, IIf(ADO6!unitario > 0, Space(9 - Len(Trim(Format(ADO6!unitario, "##,##0.00")))) + Format(ADO6!unitario, "##,##0.00"), Space(9)))
       a = IXY(40.1, fila, IIf(ADO6!netvta > 0, Space(10 - Len(Trim(Format(ADO6!netvta, "###,##0.00")))) + Format(ADO6!netvta, "###,##0.00"), Space(10)))
       fila = fila + 1.38
       
       If Trim(ADO6!nombr2) <> "" Then
          a = IXY(7.1, fila, ADO6!nombr2)
          fila = fila + 1.38
       End If
       If Trim(ADO6!nombr3) <> "" Then
          a = IXY(7.1, fila, ADO6!nombr3)
          fila = fila + 1.38
       End If
       If Trim(ADO6!nombr4) <> "" Then
          a = IXY(7.1, fila, ADO6!nombr4)
          fila = fila + 1.38
       End If
       If Trim(ADO6!nombr5) <> "" Then
          a = IXY(7.1, fila, ADO6!nombr5)
          fila = fila + 1.38
       End If
       If Trim(ADO6!nombr6) <> "" Then
          a = IXY(7.1, fila, ADO6!nombr6)
          fila = fila + 1.38
       End If
       If Trim(ADO6!nombr7) <> "" Then
          a = IXY(7.1, fila, ADO6!nombr7)
          fila = fila + 1.38
       End If
       If Trim(ADO6!nombr8) <> "" Then
          a = IXY(7.1, fila, ADO6!nombr8)
          fila = fila + 1.38
       End If
       If Trim(ADO6!nombr9) <> "" Then
          a = IXY(7.1, fila, ADO6!nombr9)
          fila = fila + 1.38
       End If
       If Trim(ADO6!nombrA) <> "" Then
          a = IXY(7.1, fila, ADO6!nombrA)
          fila = fila + 1.38
       End If
       If Trim(ADO6!nombrB) <> "" Then
          a = IXY(7.1, fila, ADO6!nombrB)
          fila = fila + 1.38
       End If
       If Trim(ADO6!nombrC) <> "" Then
          a = IXY(7.1, fila, ADO6!nombrC)
          fila = fila + 1.38
       End If
       If Trim(ADO6!nombrD) <> "" Then
          a = IXY(7.1, fila, ADO6!nombrD)
          fila = fila + 1.38
       End If
       If Trim(ADO6!nombrF) <> "" Then
          a = IXY(7.1, fila, ADO6!nombrF)
          fila = fila + 1.38
       End If
       If Trim(ADO6!nombrG) <> "" Then
          a = IXY(7.1, fila, ADO6!nombrG)
          fila = fila + 1.38
       End If
       
       ADO6.MoveNext
    Loop
    a = IXY(6, 58.4, "SON " + NumLetras(Val(ADO7!totvta)) + IIf(ADO7!moneda = "D", " DOLARES AMERICANOS", " NUEVOS SOLES"))
     
    a = IXY(40.4, 61.8, IIf(ADO7!netvta > 0, IIf(ADO7!moneda = "D", "US$", "S/.") + Space(10 - Len(Trim(Format(ADO7!netvta, "###,##0.00")))) + Format(ADO7!netvta, "###,##0.00"), Space(13)))
    a = IXY(36.6, 63.4, Format(wporigv, "#0") + "%")
    a = IXY(40.4, 63.4, IIf(ADO7!igvvta > 0, IIf(ADO7!moneda = "D", "US$", "S/.") + Space(10 - Len(Trim(Format(ADO7!igvvta, "###,##0.00")))) + Format(ADO7!igvvta, "###,##0.00"), Space(13)))
    a = IXY(40.4, 64.9, IIf(ADO7!totvta > 0, IIf(ADO7!moneda = "D", "US$", "S/.") + Space(10 - Len(Trim(Format(ADO7!totvta, "###,##0.00")))) + Format(ADO7!totvta, "###,##0.00"), Space(13)))
    
    Printer.FontName = "Draft 12cpi"
    Printer.FontBold = True
    Printer.PrintQuality = -1
    
    Printer.EndDoc
    
    Set ADO7 = Nothing
    Set ADO6 = Nothing
    Set ADO5 = Nothing
    Exit Sub
ErrManejo:
'    MsgBox "Se Cancela La Opción de Impresión", vbExclamation
    Resume Next
'    Exit Sub
End Sub

Public Sub printbol(zTdo As String, zSer As String, zDoc As String)
    MsgBox "Prepare la Impresora"
    Dim a As String, numreg As Long, fila As Double
    Dim wpor1 As Currency, wpor2 As Currency, wpor3 As Currency
    Dim wdsc1 As Currency, wdsc2 As Currency, wdsc3 As Currency
    
    On Error GoTo ErrManejo
    numreg = Leerado7("SELECT * FROM VTACAB WHERE TIPDOC='" + zTdo + "' AND " _
                                            & "   SERDOC='" + zSer + "' AND " _
                                            & "   NUMDOC='" + zDoc + "' ")
    If numreg = 0 Then
       MsgBox "Venta a Imprimir No Tiene Cabecera", vbInformation
       Exit Sub
    End If
    If ADO7!anulado Then
       MsgBox "Venta a Imprimir Esta Anulado", vbExclamation
       Exit Sub
    End If
    
    numreg = Leerado6("SELECT * FROM VTADET WHERE TIPDOC='" + zTdo + "' AND " _
                                            & "   SERDOC='" + zSer + "' AND " _
                                            & "   NUMDOC='" + zDoc + "' ORDER BY LINEA ")
    If numreg = 0 Then
       MsgBox "Venta a Imprimir No Tiene Detalle", vbInformation
       Exit Sub
    End If
    
    numreg = Leerado5("SELECT * FROM MAECLIENTE WHERE CODIGO='" + ADO7!cliente + "'")
    If numreg = 0 Then
       MsgBox "Codigo de Cliente " + ADO7!cliente + " a Imprimir No Existe", vbInformation
       Exit Sub
    End If
       
    Printer.FontName = "Draft 12cpi"
    Printer.FontBold = True
    Printer.PrintQuality = -1
    
    a = IXY(32, 6.8, "BOLETA " + zSer + "-" + zDoc)
    
       a = IXY(3.7, 8.8, Format(Day(ADO7!fecha), "00"))
    a = IXY(8.3, 8.8, funnommes(Format(Month(ADO7!fecha), "00")))
    a = IXY(18.7, 8.8, Format(Year(ADO7!fecha), "####"))
   
    
    a = IXY(6.7, 10.3, ADO5!nombre)
    a = IXY(37.7, 10.3, ADO5!codigo)
    
    a = IXY(6.7, 11.5, Mid(IIf(IsNull(ADO5!direc), "", Trim(ADO5!direc)), 1, 63))
    
    a = IXY(3.7, 14.1, ADO5!codigo)
    a = IXY(15.7, 14.1, ADO7!vendedor)
    a = IXY(22.7, 14.1, ADO7!condvta)
    a = IXY(36.7, 14.1, IIf(ADO7!moneda = "S", "SOLES", "DOLARES"))
    
       
    Printer.FontName = "Draft 20cpi"
    Printer.FontBold = True
    Printer.PrintQuality = -1
    
    fila = 17.2
    ADO6.MoveFirst
    Do While Not ADO6.EOF
       
    a = IXY(0.5, fila, IIf(ADO6!cantidad > 0, Space(8 - Len(Trim(Format(ADO6!cantidad, "###0.000")))) + Format(ADO6!cantidad, "###0.000"), Space(8)) + " " + ADO6!um)
    a = IXY(6.5, fila, ADO6!articulo)
    a = IXY(10.5, fila, Mid(ADO6!nombre, 1, 50))
    a = IXY(36.5, fila, IIf(ADO6!unitatot > 0, Space(8 - Len(Trim(Format(ADO6!unitatot, "####0.00")))) + Format(ADO6!unitatot, "####0.00"), Space(8)))
    a = IXY(41.2, fila, IIf(ADO7!moneda = "D", "US$", "S/.") + IIf(ADO6!totvta > 0, Space(9 - Len(Trim(Format(ADO6!totvta, "##,##0.00")))) + Format(ADO6!totvta, "##,##0.00"), Space(9)))
       fila = fila + 1.15
       ADO6.MoveNext
    Loop
    
    Printer.FontName = "Draft 12cpi"
    Printer.FontBold = True
    Printer.PrintQuality = -1
    
    a = IXY(4, 27.6, NumLetras(Val(ADO7!totvta)) + IIf(ADO7!moneda = "D", " DOLARES AMERICANOS", " NUEVOS SOLES"))
    a = IXY(20, 28.6, "S.E.u O.")
    
    a = IXY(39.1, 30.2, IIf(ADO7!moneda = "D", "US$", "S/.") + IIf(ADO7!totvta > 0, Space(9 - Len(Trim(Format(ADO7!totvta, "##,##0.00")))) + Format(ADO7!totvta, "##,##0.00"), Space(9)))
    
    Printer.FontName = "Draft 12cpi"
    Printer.FontBold = True
    Printer.PrintQuality = -1
    
    Printer.EndDoc
    
    Set ADO7 = Nothing
    Set ADO6 = Nothing
    Set ADO5 = Nothing
    Exit Sub
ErrManejo:
'    MsgBox "Se Cancela La Opción de Impresión", vbExclamation
    Resume Next
'    Exit Sub
End Sub

Public Sub printgui(zTmo As String, zSgu As String, zGui As String)
    MsgBox "Prepare la Impresora"
    Dim a As String, numreg As Long, fila As Double, col As Double
    
    On Error GoTo ErrManejo
    numreg = Leerado7("SELECT * FROM GUICAB WHERE TIPGUI='" + zTmo + "' AND " _
                                            & "   SERGUI='" + zSgu + "' AND " _
                                            & "   NUMGUI='" + zGui + "' ")
    If numreg = 0 Then
       MsgBox "Guia de Remisión a Imprimir No Tiene Cabecera", vbInformation
       Exit Sub
    End If
    If ADO7!anulado Then
       MsgBox "Guia de Remisión a Imprimir Esta Anulada", vbExclamation
       Exit Sub
    End If
    
    numreg = Leerado6("SELECT * FROM GUIDET WHERE TIPGUI='" + zTmo + "' AND " _
                                            & "   SERGUI='" + zSgu + "' AND " _
                                            & "   NUMGUI='" + zGui + "' ORDER BY LINEA ")
    If numreg = 0 Then
       MsgBox "Guia de Remisión a Imprimir No Tiene Detalle", vbInformation
       Exit Sub
    End If
    
       
    Printer.FontName = "Draft 12cpi"
    Printer.FontBold = True
    Printer.PrintQuality = -1
    
    If zTmo = "01" Then
       Dim wNomPro As String
       wNomPro = Space(1)
       numreg = Leerado8("SELECT * FROM MAEPROVEED WHERE CODIGO='" + ADO7!proveed + "' ")
       If numreg > 0 Then
          wNomPro = ADO8!nombre
          a = IXY(32, 0.01, ADO8!nombre)
       End If
    Else
       numreg = Leerado5("SELECT * FROM MAECLIENTE WHERE CODIGO='" + ADO7!cliente + "' ")
       If numreg = 0 Then
          MsgBox "Codigo de Cliente " + ADO7!cliente + " a Imprimir No Existe", vbInformation
          Exit Sub
       End If
    End If
    
    If zTmo = "14" Or zTmo = "19" Or zTmo = "01" Then
       a = IXY(34, 6.7, "GUIA " + zSgu + "-" + zGui)
    Else
       If Not IsNull(ADO7!serref) And Not IsNull(ADO7!numref) Then
          a = IXY(34, 6.7, "GUIA " + ADO7!serref + "-" + ADO7!numref)
       End If
    End If
    
    a = IXY(5.2, 8.7, Format(Day(ADO7!fecgui), "00"))
    a = IXY(6.7, 8.7, Format(Month(ADO7!fecgui), "00"))
    a = IXY(8.2, 8.7, Format(Year(ADO7!fecgui), "0000"))
    
    a = IXY(20.2, 8.7, Format(Day(ADO7!fecini), "00"))
    a = IXY(21.7, 8.7, Format(Month(ADO7!fecini), "00"))
    a = IXY(23.2, 8.7, Format(Year(ADO7!fecini), "0000"))
    
    Printer.FontName = "Draft 17cpi"
    Printer.FontBold = True
'    Printer.PrintQuality = -1
    
    a = IXY(3.2, 11.2, IIf(IsNull(ADO7!glosa1), "", Left(ADO7!glosa1, 64)))
    a = IXY(27.2, 11.2, IIf(IsNull(ADO7!glosa2), "", Left(ADO7!glosa2, 75)))
    
    If zTmo = "01" Then
       a = IXY(26, 12.6, wnomcia)
       a = IXY(30, 15, wruccia)
    Else
       a = IXY(27.8, 12.6, IIf(IsNull(ADO5!nombre), "", ADO5!nombre))
       a = IXY(28.9, 15, IIf(IsNull(ADO5!codigo), "", ADO5!codigo))
    End If
    
    a = IXY(8, 16.5, IIf(IsNull(ADO7!vehmarca), "", ADO7!vehmarca))
    a = IXY(32, 17.7, IIf(IsNull(ADO7!transnom), "", ADO7!transnom))
    a = IXY(11, 17.7, IIf(IsNull(ADO7!vehcerti), "", ADO7!vehcerti))
    a = IXY(32, 18.9, IIf(IsNull(ADO7!transruc), "", ADO7!transruc))
    a = IXY(11, 18.9, IIf(IsNull(ADO7!vehlicen), "", ADO7!vehlicen))
   
    fila = 22.6
    ADO6.MoveFirst
    Do While Not ADO6.EOF
       If ADO6!cantidad > 0 Then
          numreg = Leerado4("SELECT * FROM MAEARTICULO WHERE CODIGO='" + ADO6!articulo + "' ")
       
          a = IXY(0.5, fila, ADO6!articulo)
          a = IXY(3.9, fila, Mid(ADO6!nombre, 1, 45))
          a = IXY(28.5, fila, ADO6!um)
          a = IXY(31.8, fila, IIf(ADO6!cantidad > 0, Space(8 - Len(Trim(Format(ADO6!cantidad, "####0.000")))) + Format(ADO6!cantidad, "####0.000"), Space(8)))
          fila = fila + 1.4
       End If
       ADO6.MoveNext
    Loop
    
    Printer.FontName = "Draft 12cpi"
    Printer.FontBold = True
    Printer.PrintQuality = -1
    
    Printer.EndDoc
    
    Set ADO7 = Nothing
    Set ADO6 = Nothing
    Set ADO5 = Nothing
    Set ADO4 = Nothing
    Exit Sub
ErrManejo:
'    MsgBox "Se Cancela La Opción de Impresión", vbExclamation
    Resume Next
End Sub

Public Sub printdeb(zTdo As String, zSer As String, zdeb As String)
    MsgBox "Prepare la Impresora"
    Dim a As String, numreg As Long, fila As Double
    
    On Error GoTo ErrManejo
    numreg = Leerado7("SELECT * FROM NOTCAB WHERE TIPDOC='" + zTdo + "' AND " _
                                            & "   SERDOC='" + zSer + "' AND " _
                                            & "   NUMDOC='" + zdeb + "' ")
    If numreg = 0 Then
       MsgBox "Nota Contable a Imprimir No Tiene Cabecera", vbInformation
       Exit Sub
    End If
    If ADO7!anulado Then
       MsgBox "Nota Contable a Imprimir Esta Anulada", vbExclamation
       Exit Sub
    End If
    
    numreg = Leerado6("SELECT * FROM NOTDET WHERE TIPDOC='" + zTdo + "' AND " _
                                            & "   SERDOC='" + zSer + "' AND " _
                                            & "   NUMDOC='" + zdeb + "' ORDER BY LINEA ")
    If numreg = 0 Then
       MsgBox "Nota Contable a Imprimir No Tiene Detalle", vbInformation
       Exit Sub
    End If
    
    numreg = Leerado5("SELECT * FROM MAECLIENTE WHERE CODIGO='" + ADO7!cliente + "'")
    If numreg = 0 Then
       MsgBox "Codigo de Cliente " + ADO7!cliente + " a Imprimir No Existe", vbInformation
       Exit Sub
    End If
       
    Printer.FontName = "Draft 12cpi"
    Printer.FontBold = True
    Printer.PrintQuality = -1
    a = IXY(28.5, 6.5, "NOTA DEBITO " + zSer + "-" + zdeb)
    a = IXY(2.5, 10.3, IIf(IsNull(ADO5!nombre), "", ADO5!nombre))
    If Len(Trim(ADO7!refnumdoc)) > 0 Then
       a = IXY(32, 10.3, IIf(ADO7!refnumdoc = "01", "FACTURA ", "BOLETA "))
    End If
    a = IXY(4.5, 11.5, ADO7!cliente)
    If Len(Trim(ADO7!refnumdoc)) > 0 Then
       a = IXY(32, 11.5, ADO7!refserdoc + "-" + ADO7!refnumdoc)
    End If
    a = IXY(4.5, 12.7, Format(ADO7!fecha, "dd/mm/yyyy"))
    If Len(Trim(ADO7!refnumdoc)) > 0 Then
       a = IXY(30, 12.7, Format(Day(ADO7!reffecha), "00"))
       a = IXY(33.9, 12.7, funnommes(Format(Month(ADO7!reffecha), "00")))
       a = IXY(40.5, 12.7, Right(Format(Year(ADO7!reffecha), "0000"), 1))
    End If
    If ADO6.RecordCount > 0 Then
       fila = 17.2
       ADO6.MoveFirst
       Do While Not ADO6.EOF
          If Left(ADO6!articulo, 3) <> "999" Then
             a = IXY(1.6, fila, IIf(ADO6!cantidad > 0, Space(6 - Len(Trim(Format(ADO6!cantidad, "##.##0")))) + Format(ADO6!cantidad, "##.##0"), Space(6)))
             a = IXY(9.8, fila, ADO6!um)
             a = IXY(12.8, fila, Mid(ADO6!nombre, 1, 50))
             a = IXY(33.3, fila, IIf(ADO7!moneda = "D", "US$", "S/.") + IIf(ADO6!unitario > 0, Space(7 - Len(Trim(Format(ADO6!unitario, "###0.00")))) + Format(ADO6!unitario, "###0.00"), Space(7)))
             a = IXY(39.3, fila, IIf(ADO7!moneda = "D", "US$", "S/.") + IIf(ADO6!valvta > 0, Space(9 - Len(Trim(Format(ADO6!valvta, "##,##0.00")))) + Format(ADO6!valvta, "##,##0.00"), Space(9)))
             fila = fila + 1.25
          End If
          ADO6.MoveNext
       Loop
       a = IXY(5, fila + 1.2, IIf(IsNull(ADO7!glosa1), "", ADO7!glosa1))
       a = IXY(5, fila + 2.4, IIf(IsNull(ADO7!glosa2), "", ADO7!glosa2))
       a = IXY(5, fila + 3.6, IIf(IsNull(ADO7!glosa3), "", ADO7!glosa3))
       a = IXY(5, fila + 4.8, IIf(IsNull(ADO7!glosa4), "", ADO7!glosa4))
    Else
       a = IXY(5, 20.2, IIf(IsNull(ADO7!glosa1), "", ADO7!glosa1))
       a = IXY(5, 21.4, IIf(IsNull(ADO7!glosa2), "", ADO7!glosa2))
       a = IXY(5, 22.6, IIf(IsNull(ADO7!glosa3), "", ADO7!glosa3))
       a = IXY(5, 23.8, IIf(IsNull(ADO7!glosa4), "", ADO7!glosa4))
    End If
    a = IXY(4.5, 27.5, "SON " + NumLetras(Val(ADO7!totvta)) + IIf(ADO7!moneda = "S", " NUEVOS SOLES", " DOLARES AMERICANOS"))
   
'    a = IXY(36.3, 28.9, IIf(ADO7!moneda = "D", "US$", "S/.") + IIf(ADO7!netvta > 0, Space(9 - Len(Trim(Format(ADO7!netvta, "##,##0.00")))) + Format(ADO7!netvta, "##,##0.00"), Space(9)))
    a = IXY(36.3, 28.9, IIf(ADO7!moneda = "D", "US$", "S/.") + IIf(ADO7!igvvta > 0, Space(9 - Len(Trim(Format(ADO7!igvvta, "##,##0.00")))) + Format(ADO7!igvvta, "##,##0.00"), Space(9)))
    a = IXY(36.3, 31.7, IIf(ADO7!moneda = "D", "US$", "S/.") + IIf(ADO7!totvta > 0, Space(9 - Len(Trim(Format(ADO7!totvta, "##,##0.00")))) + Format(ADO7!totvta, "##,##0.00"), Space(9)))
    
    Printer.EndDoc
    
    Set ADO7 = Nothing
    Set ADO6 = Nothing
    Set ADO5 = Nothing
    Exit Sub
ErrManejo:
    MsgBox "Se Cancela La Opción de Impresión", vbExclamation
    Exit Sub
End Sub

Public Sub printcre(zTdo As String, zSer As String, zCRe As String)
    MsgBox "Prepare la Impresora"
    Dim a As String, numreg As Long, fila As Double
    
    On Error GoTo ErrManejo
    numreg = Leerado7("SELECT * FROM NOTCAB WHERE TIPDOC='" + zTdo + "' AND " _
                                            & "   SERDOC='" + zSer + "' AND " _
                                            & "   NUMDOC='" + zCRe + "' ")
    If numreg = 0 Then
       MsgBox "Nota Contable a Imprimir No Tiene Cabecera", vbInformation
       Exit Sub
    End If
    If ADO7!anulado Then
       MsgBox "Nota Contable a Imprimir Esta Anulada", vbExclamation
       Exit Sub
    End If
    
    numreg = Leerado6("SELECT * FROM NOTDET WHERE TIPDOC='" + zTdo + "' AND " _
                                            & "   SERDOC='" + zSer + "' AND " _
                                            & "   NUMDOC='" + zCRe + "' ORDER BY LINEA ")
    If numreg = 0 Then
       MsgBox "Nota Contable a Imprimir No Tiene Detalle", vbInformation
       Exit Sub
    End If
    
    numreg = Leerado5("SELECT * FROM MAECLIENTE WHERE CODIGO='" + ADO7!cliente + "'")
    If numreg = 0 Then
       MsgBox "Codigo de Cliente " + ADO7!cliente + " a Imprimir No Existe", vbInformation
       Exit Sub
    End If
       
    Printer.FontName = "Draft 12cpi"
    Printer.FontBold = True
    Printer.PrintQuality = -1
    
    a = IXY(28.5, 6.5, "NOTA CREDITO " + zSer + "-" + zCRe)
    a = IXY(2.5, 10.3, IIf(IsNull(ADO5!nombre), "", ADO5!nombre))
    If Len(Trim(ADO7!refnumdoc)) > 0 Then
       a = IXY(32, 10.3, IIf(ADO7!refnumdoc = "01", "FACTURA ", "BOLETA "))
    End If
    a = IXY(4.5, 11.5, ADO7!cliente)
    If Len(Trim(ADO7!refnumdoc)) > 0 Then
       a = IXY(32, 11.5, ADO7!refserdoc + "-" + ADO7!refnumdoc)
    End If
    a = IXY(4.5, 12.7, Format(ADO7!fecha, "dd/mm/yyyy"))
    If Len(Trim(ADO7!refnumdoc)) > 0 Then
       a = IXY(30, 12.7, Format(Day(ADO7!reffecha), "00"))
       a = IXY(33.9, 12.7, funnommes(Format(Month(ADO7!reffecha), "00")))
       a = IXY(42.8, 12.7, Right(Format(Year(ADO7!reffecha), "0000"), 1))
    End If
    If ADO6.RecordCount > 0 Then
       fila = 17.2
       ADO6.MoveFirst
       Do While Not ADO6.EOF
          If Left(ADO6!articulo, 3) <> "999" Then
             a = IXY(1.6, fila, IIf(ADO6!cantidad > 0, Space(6 - Len(Trim(Format(ADO6!cantidad, "##.##0")))) + Format(ADO6!cantidad, "##.##0"), Space(6)))
             a = IXY(9.8, fila, ADO6!um)
             a = IXY(12.8, fila, Mid(ADO6!nombre, 1, 50))
             a = IXY(33.3, fila, IIf(ADO7!moneda = "D", "US$", "S/.") + IIf(ADO6!unitario > 0, Space(7 - Len(Trim(Format(ADO6!unitario, "###0.00")))) + Format(ADO6!unitario, "###0.00"), Space(7)))
             a = IXY(39.3, fila, IIf(ADO7!moneda = "D", "US$", "S/.") + IIf(ADO6!valvta > 0, Space(9 - Len(Trim(Format(ADO6!valvta, "##,##0.00")))) + Format(ADO6!valvta, "##,##0.00"), Space(9)))
             fila = fila + 1.25
          Else
             a = IXY(36.3, fila, IIf(ADO7!moneda = "D", "US$", "S/.") + IIf(ADO6!netvta > 0, Space(9 - Len(Trim(Format(ADO6!netvta, "##,##0.00")))) + Format(ADO6!netvta, "##,##0.00"), Space(9)))
          End If
          ADO6.MoveNext
       Loop
'       a = IXY(5, fila + 1.2, IIf(IsNull(ADO7!glosa1), "", ADO7!glosa1))
'       a = IXY(5, fila + 2.4, IIf(IsNull(ADO7!glosa2), "", ADO7!glosa2))
'       a = IXY(5, fila + 3.6, IIf(IsNull(ADO7!glosa3), "", ADO7!glosa3))
'       a = IXY(5, fila + 4.8, IIf(IsNull(ADO7!glosa4), "", ADO7!glosa4))
    Else
'       a = IXY(5, 20.2, IIf(IsNull(ADO7!glosa1), "", ADO7!glosa1))
'       a = IXY(5, 21.4, IIf(IsNull(ADO7!glosa2), "", ADO7!glosa2))
'       a = IXY(5, 22.6, IIf(IsNull(ADO7!glosa3), "", ADO7!glosa3))
'       a = IXY(5, 23.8, IIf(IsNull(ADO7!glosa4), "", ADO7!glosa4))
    End If
    a = IXY(36.3, 19.9, IIf(ADO7!moneda = "D", "US$", "S/.") + IIf(ADO7!igvvta > 0, Space(9 - Len(Trim(Format(ADO7!igvvta, "##,##0.00")))) + Format(ADO7!igvvta, "##,##0.00"), Space(9)))
    a = IXY(36.3, 21.7, IIf(ADO7!moneda = "D", "US$", "S/.") + IIf(ADO7!totvta > 0, Space(9 - Len(Trim(Format(ADO7!totvta, "##,##0.00")))) + Format(ADO7!totvta, "##,##0.00"), Space(9)))
    
    Printer.EndDoc
    
    Set ADO7 = Nothing
    Set ADO6 = Nothing
    Set ADO5 = Nothing
    Exit Sub
ErrManejo:
    MsgBox "Se Cancela La Opción de Impresión", vbExclamation
    Exit Sub
End Sub

Public Sub printlet(zcanje As String, zLin As String, zlet As String)
   Dim aa As Integer
   Dim zcli As String, zNom As String, zFec As String, zVcm As String
   Dim zdir As String, zImp As Currency, zMon As String, ztel As String
   Dim zTdo As String, zSer As String, zfac As String
   aa = Leerado7("SELECT * FROM CANJES WHERE CANJE = '" + zcanje + "' ")
   If aa = 0 Then
      MsgBox "Numero de Canje No Existe", vbExclamation
      Exit Sub
   End If
   zTdo = Left(ADO7!numfac01, 2)
   zSer = Mid(ADO7!numfac01, 3, 3)
   zfac = Right(ADO7!numfac01, 6)
   zcli = ADO7!clifac01
   zFec = Format(ADO7!fecha, "dd/mm/yyyy")
   zMon = ADO7!moneda
   zNom = Space(1)
   zdir = Space(1)
   ztel = Space(1)
   aa = Leerado6("SELECT * FROM MAECLIENTE WHERE CODIGO = '" + zcli + "' ")
   If aa > 0 Then
      zNom = ADO6!nombre
      zdir = ADO6!direc
      ztel = ADO6!telfs
   End If
   Select Case zLin
   Case "1"
        zVcm = Format(ADO7!vcmlet01, "dd/mm/yyyy")
        zImp = ADO7!implet01
   Case "2"
        zVcm = Format(ADO7!vcmlet02, "dd/mm/yyyy")
        zImp = ADO7!implet02
   Case "3"
        zVcm = Format(ADO7!vcmlet03, "dd/mm/yyyy")
        zImp = ADO7!implet03
   Case "4"
        zVcm = Format(ADO7!vcmlet04, "dd/mm/yyyy")
        zImp = ADO7!implet04
   Case "5"
        zVcm = Format(ADO7!vcmlet05, "dd/mm/yyyy")
        zImp = ADO7!implet05
   Case "6"
        zVcm = Format(ADO7!vcmlet06, "dd/mm/yyyy")
        zImp = ADO7!implet06
   Case "7"
        zVcm = Format(ADO7!vcmlet07, "dd/mm/yyyy")
        zImp = ADO7!implet07
   Case "8"
        zVcm = Format(ADO7!vcmlet08, "dd/mm/yyyy")
        zImp = ADO7!implet08
   Case "9"
        zVcm = Format(ADO7!vcmlet09, "dd/mm/yyyy")
        zImp = ADO7!implet09
   Case "10"
        zVcm = Format(ADO7!vcmlet10, "dd/mm/yyyy")
        zImp = ADO7!implet10
   Case "11"
        zVcm = Format(ADO7!vcmlet11, "dd/mm/yyyy")
        zImp = ADO7!implet11
   Case "12"
        zVcm = Format(ADO7!vcmlet12, "dd/mm/yyyy")
        zImp = ADO7!implet12
   Case "13"
        zVcm = Format(ADO7!vcmlet13, "dd/mm/yyyy")
        zImp = ADO7!implet13
   Case "14"
        zVcm = Format(ADO7!vcmlet14, "dd/mm/yyyy")
        zImp = ADO7!implet14
   Case "15"
        zVcm = Format(ADO7!vcmlet15, "dd/mm/yyyy")
        zImp = ADO7!implet15
   Case "16"
        zVcm = Format(ADO7!vcmlet16, "dd/mm/yyyy")
        zImp = ADO7!implet16
   Case "17"
        zVcm = Format(ADO7!vcmlet17, "dd/mm/yyyy")
        zImp = ADO7!implet17
   Case "18"
        zVcm = Format(ADO7!vcmlet18, "dd/mm/yyyy")
        zImp = ADO7!implet18
   Case "19"
        zVcm = Format(ADO7!vcmlet19, "dd/mm/yyyy")
        zImp = ADO7!implet19
   Case "20"
        zVcm = Format(ADO7!vcmlet20, "dd/mm/yyyy")
        zImp = ADO7!implet20
   End Select
  
   Printer.FontName = "Draft 17cpi"
   Printer.FontBold = True
   Printer.PrintQuality = -1
   
   aa = IXY(4.8, 2.5, zlet)
   aa = IXY(10.5, 2.5, IIf(zTdo = "01", "F/", "B/") + zSer + "-" + zfac)
   aa = IXY(18.5, 2.5, "LIMA")
   aa = IXY(23.8, 2.5, Format(Day(zFec), "00") + " / " + _
                      Format(Month(zFec), "00") + " / " + _
                      Format(Year(zFec), "0000"))
   
   aa = IXY(29.9, 2.5, Format(Day(zVcm), "00") + " / " + _
                      Format(Month(zVcm), "00") + " / " + _
                      Format(Year(zVcm), "0000"))

   aa = IXY(36.2, 2.5, IIf(ADO7!moneda = "D", "US$", "S/.") + IIf(zImp > 0, Space(10 - Len(Trim(Format(zImp, "###,##0.00")))) + Format(zImp, "###,##0.00"), Space(10)))
   
   aa = IXY(4.5, 6.1, NumLetras(zImp))

   aa = IXY(5.9, 8.1, zNom)
   
   aa = IXY(5.9, 10.1, zdir)

   aa = IXY(5.9, 11.4, zcli)
   aa = IXY(16.5, 11.4, ztel)

   Printer.FontName = "Draft 17cpi"
   Printer.FontBold = True

   Printer.EndDoc
    
   Set ADO7 = Nothing
   Set ADO6 = Nothing

   Exit Sub
ErrManejo:
   MsgBox "Se Cancela La Opción de Impresión", vbExclamation
   Resume Next
End Sub

Public Function GlosaLibre(zNom As Variant)
   Dim wLen As Integer, wres As String, I As Integer
   wres = ""
   If IsNull(zNom) Then
      GlosaLibre = wres
      Exit Function
   End If
   wLen = Len(Trim(zNom))
   For I = 1 To wLen
       If Mid(zNom, I, 1) = "'" Then
          wres = wres + " "
       Else
          wres = wres + Mid(zNom, I, 1)
       End If
   Next
   GlosaLibre = wres
End Function

Public Function LlenaDat(zNom As Variant, zNum As Integer)
   If IsNull(zNom) Then
      zNom = Space(zNum)
   Else
      If Len(Trim(zNom)) < zNum Then
         zNom = Trim(zNom) + Space(zNum - Len(Trim(zNom)))
      Else
         zNom = Left(zNom, zNum)
      End If
   End If
   LlenaDat = zNom
End Function

Public Sub ActualizaSaldos(zSoc As Integer, zMes As String, zCon As String)
   On Error GoTo err

   Dim zz As Integer, zE_s As String, zMon As String, zApo As Currency, _
       zTipCob As String, zSerCob As String, zNumCob As String, zLinCob As String, _
       zSdoOld As Currency, zSdoNew As Currency, _
       zCargos As Currency, zAbonos As Currency, _
       zTotOld As Currency, zTotNew As Currency
   zTotOld = 0: zTotNew = 0
   zSdoOld = 0: zSdoNew = 0
   zCargos = 0: zAbonos = 0
   
   zz = Leerado8a("SELECT * FROM CTASXDET " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' AND " _
                & "       CONCEPTO = '" + zCon + "' " _
                & " ORDER BY FECHA, TIPMOV, TIPCOB, SERCOB, NUMCOB, LINCOB ")
   If zz > 0 Then
      Do While Not ADO8a.EOF
         zTipCob = ADO8a!tipcob
         zSerCob = ADO8a!sercob
         zNumCob = ADO8a!numcob
         zLinCob = ADO8a!lincob
         
         zSdoNew = zSdoOld + ADO8a!cargos - ADO8a!abonos
         zCargos = zCargos + ADO8a!cargos
         zAbonos = zAbonos + ADO8a!abonos
            
         Db.BeginTrans
         Db.Execute ("UPDATE CTASXDET " _
         & " SET SDOOLD = " + Str(zSdoOld) + ", " _
         & "     SDONEW = " + Str(zSdoNew) + " " _
         & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
         & "            MES = '" + zMes + "' AND " _
         & "       CONCEPTO = '" + zCon + "' AND " _
         & "         TIPCOB = '" + zTipCob + "' AND " _
         & "         SERCOB = '" + zSerCob + "' AND " _
         & "         NUMCOB = '" + zNumCob + "' AND " _
         & "         LINCOB = '" + zLinCob + "' ")
         Db.CommitTrans
   
         zSdoOld = zSdoNew
   
         ADO8a.MoveNext
      Loop
   End If
   
   zz = Leerado8a("SELECT * FROM CTASXCAB " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' AND " _
                & "       CONCEPTO = '" + zCon + "' ")
   If zz = 0 Then
      zE_s = "": zMon = "": zApo = 0
      
      zz = Leerado7a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
      If zz > 0 Then
         zE_s = IIf(IsNull(ADO7a!e_socio), "", ADO7a!e_socio)
      End If
      Set ADO7a = Nothing
      
      zz = Leerado7a("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + zE_s + "' ")
      If zz > 0 Then
         zMon = ADO7a!moneda
         zApo = ADO7a!aporte
      End If
      Set ADO7a = Nothing
      
      If zCargos <> 0 Or zAbonos <> 0 Then
         Db.BeginTrans
         Db.Execute ("INSERT INTO CTASXCAB " _
         & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, " _
         & "  CARGOS, ABONOS, SDONEW ) " _
         & " VALUES " _
         & " (" + Str(zSoc) + ", '" + zMes + "', '" + zCon + "', " _
         & "  '" + zE_s + "', '" + zMon + "', " _
         & "  " + Str(zCargos) + ", " + Str(zAbonos) + ", " + Str(zSdoNew) + " ) ")
         Db.CommitTrans
      End If
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE CTASXCAB " _
      & " SET CARGOS = " + Str(zCargos) + ", " _
      & "     ABONOS = " + Str(zAbonos) + ", " _
      & "     SDONEW = " + Str(zSdoNew) + " " _
      & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
      & "            MES = '" + zMes + "' AND " _
      & "       CONCEPTO = '" + zCon + "' ")
      Db.CommitTrans
   End If
   
   Exit Sub
err:
   MsgBox err.Description
   Resume Next
End Sub

Public Sub CreaCtasxDetOLD(zSoc As Integer, zAno As String, zMes As String, zTip As String, zDsc As Currency)
   On Error GoTo err
   
   Dim yy As Integer, zE_s As String, zFec As Date
   
   zFec = Format("20/" + zMes + "/" + zAno, "dd/mm/yyyy")
   zE_s = ""
   yy = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
   If yy > 0 Then
      zE_s = ADO6a!e_socio
   End If
   Set ADO6a = Nothing
   
   yy = Leerado6a("SELECT * FROM CTASXCAB " _
           & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
           & "            MES = '" + zAno + "/" + zMes + "' AND " _
           & "       CONCEPTO = '01' ")
   If yy = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO CTASXCAB " _
      & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, CARGOS, ABONOS, SDONEW ) " _
      & " VALUES " _
      & " (" + Str(zSoc) + ", '" + zAno + "/" + zMes + "', '01', '" + zE_s + "', 'S', 0, 0, 0 ) ")
      Db.CommitTrans
   End If
   Set ADO6a = Nothing
   
   yy = Leerado6a("SELECT * FROM CTASXDET " _
           & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
           & "            MES = '" + zAno + "/" + zMes + "' AND " _
           & "       CONCEPTO = '01' AND " _
           & "         TIPCOB = '" + zTip + "' AND " _
           & "         SERCOB = '001' AND " _
           & "         NUMCOB = '" + Right(zAno, 2) + zMes + "00001' AND " _
           & "         LINCOB = '0001' AND " _
           & "         TIPMOV = '2' ")
   If yy = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO CTASXDET " _
      & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, " _
      & "  FECHA, TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW ) " _
      & " VALUES " _
      & " (" + Str(zSoc) + ", '" + zAno + "/" + zMes + "', '01', '" + zTip + "', '001', '" + Right(zAno, 2) + zMes + "00001', " _
      & "  '0001', '2', '" + Format(zFec, "dd/mm/yyyy") + "', 0, 0, " + Str(zDsc) + ", 0, 0, 0, 0 ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE CTASXDET " _
      & " SET SOLESS = " + Str(zDsc) + ", DOLARE = 0 " _
      & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
      & "            MES = '" + zAno + "/" + zMes + "' AND " _
      & "       CONCEPTO = '01' AND " _
      & "         TIPCOB = '" + zTip + "' AND " _
      & "         SERCOB = '001' AND " _
      & "         NUMCOB = '" + Right(zAno, 2) + zMes + "00001' AND " _
      & "         LINCOB = '0001' AND " _
      & "         TIPMOV = '2' ")
      Db.CommitTrans
   End If
   Set ADO6a = Nothing
   
   Call ActualizaSaldos(zSoc, zAno + "/" + zMes, "01")
   Exit Sub
err:
   MsgBox err.Description
   Resume Next
End Sub

Public Function BuscaPariente(zSoc As Integer, zTip As String, zLin As String) As String
   On Error GoTo err
   
   Dim zz As Integer, zNom As String
   zNom = ""
   zz = Leerado6a("SELECT * FROM MAEFAMILIA " _
                & " WHERE     CODSOCIO = " + Str(zSoc) + " AND " _
                & "       TIPOPARIENTE = '" + zTip + "' AND " _
                & "                LIN = '" + zLin + "' " _
                & " ORDER BY LIN ")
   If zz > 0 Then
      zNom = ADO6a!nombre
   End If
   Set ADO6a = Nothing

   BuscaPariente = zNom
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaDatosSocio(zSoc As Integer, zSw As Integer) As Variant
   On Error GoTo err
   
   Dim yy As Integer, _
       zCod As Long, zIns As Integer, zNom As String, _
       zE_s As String, zDni As String
   
   zCod = 0: zIns = 0: zNom = ""
   yy = Leerado4a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
   If yy > 0 Then
      zCod = ADO4a!codigo
      zIns = ADO4a!ins
      zNom = ADO4a!nombre
      zDni = ADO4a!numdoc
      zE_s = ADO4a!e_socio
   End If
   Set ADO4a = Nothing

   Select Case zSw
   Case 1
        BuscaDatosSocio = zCod
   Case 2
        BuscaDatosSocio = zIns
   Case 3
        BuscaDatosSocio = zNom
   Case 4
        BuscaDatosSocio = zDni
   Case 5
        BuscaDatosSocio = zE_s
   End Select
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaEstadoAsignado(zCod As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer
   zRes = 0
   zz = Leerado5a("SELECT * FROM MAEESTADOASIGNADO WHERE CODIGO = '" + zCod + "' ")
   If zz > 0 Then
      zRes = ADO5a!num - 1
   End If
   
   BuscaEstadoAsignado = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaCodEstadoAsignado(zCod As String) As String
   On Error GoTo err
   
   Dim zz As Integer, zRes As String
   zRes = ""
   zz = Leerado5a("SELECT * FROM MAEESTADOASIGNADO WHERE NOMBRE LIKE '" + zCod + "' ")
   If zz > 0 Then
      zRes = ADO5a!codigo
   End If
   
   BuscaCodEstadoAsignado = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaGrado(zCod As Integer) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer, zNum As Integer
   zRes = 0
   zNum = 0
   zz = Leerado5a("SELECT * FROM MAEGRADO ORDER BY GRADO ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If ADO5a!grado = zCod Then
            zRes = zNum
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   BuscaGrado = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaCodGrado(zNom As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer, zLen As Integer, zNum As Integer
   zRes = 0: zLen = Len(Trim(zNom))
   zz = Leerado5a("SELECT * FROM MAEGRADO ORDER BY GRADO ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If Mid(ADO5a!nombre, 1, zLen) = Trim(zNom) Then
            zRes = ADO5a!grado
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   
   BuscaCodGrado = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaSitu(zSitu As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer, zNum As Integer
   zRes = 0
   zNum = 0
   zz = Leerado5a("SELECT * FROM MAESITU ORDER BY SITU ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If ADO5a!situ = zSitu Then
            zRes = zNum
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   BuscaSitu = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaCodSitu(zNom As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer, zLen As Integer, zNum As Integer
   zRes = 0: zLen = Len(Trim(zNom))
   zz = Leerado5a("SELECT * FROM MAESITU ORDER BY SITU ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If Mid(ADO5a!nombre, 1, zLen) = Trim(zNom) Then
            zRes = ADO5a!situ
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   
   BuscaCodSitu = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaSituEsp(zSituEsp As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer, zNum As Integer
   zRes = 0
   zNum = 0
   zz = Leerado5a("SELECT * FROM MAESITUESP ORDER BY SITUESP ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If ADO5a!situesp = zSituEsp Then
            zRes = zNum
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   BuscaSituEsp = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaCodSituEsp(zNom As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer, zLen As Integer, zNum As Integer
   zRes = 0: zLen = Len(Trim(zNom))
   zz = Leerado5a("SELECT * FROM MAESITUESP ORDER BY SITUESP ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If Mid(ADO5a!nombre, 1, zLen) = Trim(zNom) Then
            zRes = ADO5a!situesp
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   
   BuscaCodSituEsp = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaECivil(zECivil As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer, zNum As Integer
   zRes = 0
   zNum = 0
   zz = Leerado5a("SELECT * FROM MAEECIVIL ORDER BY ECIVIL ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If ADO5a!ecivil = zECivil Then
            zRes = zNum
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   BuscaECivil = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaCodECivil(zNom As String) As String
   On Error GoTo err
   
   Dim zz As Integer, zRes As String, zLen As Integer, zNum As Integer
   zRes = 0: zLen = Len(Trim(zNom))
   zz = Leerado5a("SELECT * FROM MAEECIVIL ORDER BY ECIVIL ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If Mid(ADO5a!nombre, 1, zLen) = Trim(zNom) Then
            zRes = ADO5a!ecivil
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   
   BuscaCodECivil = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaSexo(zSexo As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer, zNum As Integer
   zRes = 0
   zNum = 0
   zz = Leerado5a("SELECT * FROM MAESEXO ORDER BY SEXO ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If ADO5a!sexo = zSexo Then
            zRes = zNum
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   BuscaSexo = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaCodSexo(zNom As String) As String
   On Error GoTo err
   
   Dim zz As Integer, zRes As String, zLen As Integer, zNum As Integer
   zRes = 0: zLen = Len(Trim(zNom))
   zz = Leerado5a("SELECT * FROM MAESEXO ORDER BY SEXO ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If Mid(ADO5a!nombre, 1, zLen) = Trim(zNom) Then
            zRes = ADO5a!sexo
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   
   BuscaCodSexo = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaEsocio(zE_Socio As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer, zNum As Integer
   zRes = 0
   zNum = 0
   zz = Leerado5a("SELECT * FROM MAEE_SOCIO ORDER BY E_SOCIO ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If ADO5a!e_socio = zE_Socio Then
            zRes = zNum
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   BuscaEsocio = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaCodEsocio(zNom As String) As String
   On Error GoTo err
   
   Dim zz As Integer, zRes As String, zLen As Integer, zNum As Integer
   zRes = "": zLen = Len(Trim(zNom))
   zz = Leerado5a("SELECT * FROM MAEE_SOCIO ORDER BY E_SOCIO ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If Trim(ADO5a!nombre) = Trim(zNom) Then
            zRes = ADO5a!e_socio
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   
   BuscaCodEsocio = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaTipCob(zTipCob As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer, zNum As Integer
   zRes = 0
   zNum = 0
   zz = Leerado5a("SELECT * FROM MAETIPCOB ORDER BY TIPCOB ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If ADO5a!tipcob = zTipCob Then
            zRes = zNum
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   BuscaTipCob = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaCodTipCob(zNom As String) As String
   On Error GoTo err
   
   Dim zz As Integer, zRes As String, zLen As Integer, zNum As Integer
   zRes = 0: zLen = Len(Trim(zNom))
   zz = Leerado5a("SELECT * FROM MAETIPCOB ORDER BY TIPCOB ")
   If zz > 0 Then
      Do While Not ADO5a.EOF
         If Mid(ADO5a!nombre, 1, zLen) = Trim(zNom) Then
            zRes = ADO5a!tipcob
            Exit Do
         End If
         zNum = zNum + 1
         ADO5a.MoveNext
      Loop
   End If
   
   BuscaCodTipCob = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaConcepto(zCod As String) As Integer
   On Error GoTo err
   
   Dim zz As Integer, zRes As Integer
   zRes = 0
   zz = Leerado5a("SELECT * FROM MAECONCEPTO WHERE CONCEPTO = '" + zCod + "' ")
   If zz > 0 Then
      zRes = ADO5a!num - 1
   End If
   
   BuscaConcepto = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaCodConcepto(zCod As String) As String
   On Error GoTo err
   
   Dim zz As Integer, zRes As String
   zRes = ""
   zz = Leerado5a("SELECT * FROM MAECONCEPTO WHERE NOMBRE LIKE '" + Trim(zCod) + "%' ")
   If zz > 0 Then
      zRes = ADO5a!concepto
   End If
   
   BuscaCodConcepto = zRes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function CreaAporte(zSoc As Integer, sw As Byte) As Variant
   On Error GoTo err
   
   Dim zz As Integer, zMes As String, zE_s As String, zApo As Currency, zMon As String, zFec As Date
   zMes = "2017/09"
   zz = Leerado6a("SELECT MAX(MES) AS MES " _
                & " FROM CTASXCAB " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                & "       CONCEPTO = '01' ")
   If zz > 0 Then
      zMes = ADO6a!mes
   End If
   Set ADO6a = Nothing

   If Mid(zMes, 6, 2) < "12" Then
      zMes = Left(zMes, 5) + Format(Val(Mid(zMes, 6, 2)) + 1, "00")
   Else
      zMes = Format(Val(Mid(zMes, 1, 4)) + 1, "0000") + "/01"
   End If
   zFec = Format("01/" + Mid(zMes, 6, 2) + "/" + Mid(zMes, 1, 4), "dd/mm/yyyy")

   Db.BeginTrans
   Db.Execute ("DELETE FROM CTASXCAB " _
   & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
   & "            MES = '" + zMes + "' AND " _
   & "       CONCEPTO = '01' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM CTASXDET " _
   & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
   & "            MES = '" + zMes + "' AND " _
   & "       CONCEPTO = '01' ")
   Db.CommitTrans

   zE_s = "": zMon = "": zApo = 0
   zz = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
   If zz > 0 Then
      zE_s = ADO6a!e_socio
   End If
   Set ADO6a = Nothing

   zz = Leerado6a("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + zE_s + "' ")
   If zz > 0 Then
      zMon = ADO6a!moneda
      zApo = ADO6a!aporte
   End If
   Set ADO6a = Nothing
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO CTASXCAB " _
   & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, " _
   & "  CARGOS, ABONOS, SDONEW ) " _
   & " VALUES " _
   & " (" + Str(zSoc) + ", '" + zMes + "', '01', '" + zE_s + "', '" + zMon + "', " _
   & "  " + Str(zApo) + ", 0, " + Str(zApo) + " ) ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO CTASXDET " _
   & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
   & "  TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW, OBS ) " _
   & " VALUES " _
   & " (" + Str(zSoc) + ", '" + zMes + "', '01', '00', '', '', '', '1', " _
   & "  '" + Format(zFec, "dd/mm/yyyy") + "', " _
   & "  0, 0, 0, 0, " + Str(zApo) + ", 0, " + Str(zApo) + ", '' )  ")
   Db.CommitTrans

   If sw = 1 Then
      CreaAporte = zMes
   Else
      CreaAporte = zApo
   End If
   
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function CreaAporteMes(zSoc As Integer, zMes As String, zCon As String, sw As Byte) As Currency
   On Error GoTo err
   
   Dim zSdo As Currency, zz As Integer, zE_s As String, zApo As Currency, zMon As String, zFec As Date

   Db.BeginTrans
   Db.Execute ("DELETE FROM CTASXCAB " _
   & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
   & "            MES = '" + zMes + "' AND " _
   & "       CONCEPTO = '" + zCon + "' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("DELETE FROM CTASXDET " _
   & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
   & "            MES = '" + zMes + "' AND " _
   & "       CONCEPTO = '" + zCon + "' AND " _
   & "         TIPMOV = '1' ")
   Db.CommitTrans

   zE_s = "": zMon = "": zApo = 0
   zz = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
   If zz > 0 Then
      zE_s = ADO6a!e_socio
   End If
   Set ADO6a = Nothing

   zz = Leerado6a("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + zE_s + "' ")
   If zz > 0 Then
      zMon = ADO6a!moneda
      zApo = ADO6a!aporte
   End If
   Set ADO6a = Nothing
   
   Db.BeginTrans
   Db.Execute ("INSERT INTO CTASXCAB " _
   & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, " _
   & "  CARGOS, ABONOS, SDONEW ) " _
   & " VALUES " _
   & " (" + Str(zSoc) + ", '" + zMes + "', '" + zCon + "', '" + zE_s + "', '" + zMon + "', " _
   & "  " + Str(zApo) + ", 0, " + Str(zApo) + " ) ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("INSERT INTO CTASXDET " _
   & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
   & "  TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW, OBS ) " _
   & " VALUES " _
   & " (" + Str(zSoc) + ", '" + zMes + "', '" + zCon + "', '00', '', '', '', '1', " _
   & "  '" + Format(zFec, "dd/mm/yyyy") + "', " _
   & "  0, 0, 0, 0, " + Str(zApo) + ", 0, " + Str(zApo) + ", '' )  ")
   Db.CommitTrans

   CreaAporteMes = zApo
   
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function SaldoFoto(zSoc As Integer, zMes As String) As Currency
   On Error GoTo err
   
   Dim zSdo As Currency, zz As Integer, wMes As String, zFec As Date, _
       wCargos As Currency, wAbonos As Currency, wSdoNew As Currency, _
       zCargos As Currency, zAbonos As Currency, zSdoNew As Currency, _
       zFecIng As Date, zMesIng As String, zFecMax As Date

   zFec = fundiames(Right(zMes, 2)) + "/" + Right(zMes, 2) + "/" + Left(zMes, 4)
   zCargos = 0: zAbonos = 0: zSdoNew = 0
   
   zz = Leerado3a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
   If zz > 0 Then
      If IsNull(ADO3a!fecing) Then
         SaldoFoto = 0
         Exit Function
      End If
      zFecIng = Format(ADO3a!fecing, "dd/mm/yyyy")
   End If
   Set ADO3a = Nothing
   
   If Format(zFecIng, "yyyy/mm/dd") > Format("01/10/2017", "yyyy/mm/dd") Then
   If Day(zFecIng) >= 20 Then
      zMesIng = Format(Year(zFecIng), "0000") + "/" + Format(Month(zFecIng) + 1, "00")
      If Mid(zMesIng, 6, 2) > "12" Then
         zMesIng = Format(Val(Mid(zMesIng, 1, 4)) + 1, "0000") + "/" + Format(Val(Mid(zMesIng, 6, 2)) - 12, "00")
      End If
   Else
      zMesIng = Format(zFecIng, "yyyy/mm")
   End If
   End If
      
   If zMesIng = "" Then
      zMesIng = "2017/09"
   End If
   
   zFecMax = Format(fundiames(Format(Month(Date), "00")) + "/" + Format(Month(Date), "00") + "/" + Format(Year(Date), "0000"), "dd/mm/yyyy")
   
   zz = Leerado3a("SELECT SUM(CARGOS - ABONOS) AS CARGOS " _
              & " FROM CTASXDET " _
              & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
              & "       CONCEPTO = '01' AND " _
              & "            MES >= '" + zMesIng + "' AND " _
              & "            MES <= '" + Left(zMes, 4) + "/" + Right(zMes, 2) + "' AND " _
              & "       TIPMOV = '1' ")
   If zz > 0 Then
      ADO3a.MoveFirst
      Do While Not ADO3a.EOF
         zCargos = IIf(IsNull(ADO3a!cargos), 0, ADO3a!cargos)
         
         ADO3a.MoveNext
      Loop
   End If
   Set ADO3a = Nothing
   
   zz = Leerado3a("SELECT SUM(ABONOS - CARGOS) AS ABONOS " _
              & " FROM CTASXDET " _
              & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
              & "       CONCEPTO = '01' AND " _
              & "       FECHA <= '" + Format(zFecMax, "dd/mm/yyyy") + "' AND " _
              & "       TIPMOV = '2' ")
   If zz > 0 Then
      ADO3a.MoveFirst
      Do While Not ADO3a.EOF
         zAbonos = IIf(IsNull(ADO3a!abonos), 0, ADO3a!abonos)
         
         ADO3a.MoveNext
      Loop
   End If
   Set ADO3a = Nothing
   zSdoNew = zCargos - zAbonos
   
   
'   zz = Leerado5a("SELECT * FROM CTASXCAB " _
'                & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
'                & "            MES >= '2017/09' AND " _
'                & "            MES <= '" + Left(zMes, 4) + "/" + Right(zMes, 2) + "' AND " _
'                & "       CONCEPTO = '01' ")
'   If zz > 0 Then
'      ADO5a.MoveFirst
'      Do While Not ADO5a.EOF
'         wMes = ADO5a!mes
'         wCargos = 0: wAbonos = 0: wSdoNew = 0
   
'         zz = Leerado6a("SELECT * FROM CTASXDET " _
'                    & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
'                    & "            MES = '" + wMes + "' AND " _
'                    & "       CONCEPTO = '01' AND " _
'                    & "       FECHA <= '" + Format(zFec, "dd/mm/yyyy") + "' AND " _
'                    & "       TIPMOV <> '004' ")
'         If zz > 0 Then
'            ADO6a.MoveFirst
'            Do While Not ADO6a.EOF
'               wCargos = wCargos + ADO6a!cargos
'               wAbonos = wAbonos + ADO6a!abonos
         
'               ADO6a.MoveNext
'            Loop
'         End If
'         wSdoNew = wCargos - wAbonos
   
'         zCargos = zCargos + wCargos
'         zAbonos = zAbonos + wAbonos
'         zSdoNew = zSdoNew + wSdoNew
   
'         ADO5a.MoveNext
'      Loop
'   End If

'   zz = Leerado5a("SELECT * FROM CTASXCAB " _
'                & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
'                & "            MES >= '2017/09' AND " _
'                & "            MES > '" + Left(zMes, 4) + "/" + Right(zMes, 2) + "' AND " _
'                & "       CONCEPTO = '01' ")
'   If zz > 0 Then
'      ADO5a.MoveFirst
'      Do While Not ADO5a.EOF
'         wMes = ADO5a!mes
'         wCargos = 0: wAbonos = 0: wSdoNew = 0
   
'         zz = Leerado6a("SELECT * FROM CTASXDET " _
'                    & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
'                    & "            MES = '" + wMes + "' AND " _
'                    & "       CONCEPTO = '01' AND " _
'                    & "          FECHA <= '" + Format(zFec, "dd/mm/yyyy") + "'AND " _
'                    & "       TIPMOV <> '000' AND " _
'                    & "       TIPMOV <> '004' ")
'         If zz > 0 Then
'            ADO6a.MoveFirst
'            Do While Not ADO6a.EOF
'               wCargos = wCargos + ADO6a!cargos
'               wAbonos = wAbonos + ADO6a!abonos
         
'               ADO6a.MoveNext
'            Loop
'         End If
'         wSdoNew = wCargos - wAbonos
   
'         zCargos = zCargos + wCargos
'         zAbonos = zAbonos + wAbonos
'         zSdoNew = zSdoNew + wSdoNew
   
'         ADO5a.MoveNext
'      Loop
'   End If
   
   
   SaldoFoto = zSdoNew
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function EnvioDiecoCMP(zSoc As Integer, zMes As String, zSw As Integer) As Currency
   On Error GoTo err

   Dim zz As Integer, zEnv540 As Currency, zEnv541 As Currency
   zMes = Left(zMes, 4) + Right(zMes, 2)

   zz = Leerado6a("SELECT * FROM DIECOCAB " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnv540 = ADO6a!netsocio
      zEnv541 = ADO6a!netasig1 + ADO6a!netasig2 + ADO6a!netasig3 + ADO6a!netasig4 + ADO6a!netasig5
   End If
   Set ADO6a = Nothing

   zz = Leerado6a("SELECT * FROM CAJMPCAB " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnv540 = ADO6a!netsocio
      zEnv541 = ADO6a!netasig1 + ADO6a!netasig2 + ADO6a!netasig3 + ADO6a!netasig4 + ADO6a!netasig5
   End If
   Set ADO6a = Nothing

   If zSw = 1 Then
      EnvioDiecoCMP = zEnv540
   Else
      EnvioDiecoCMP = zEnv541
   End If
   
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaUltimoMes(zSoc As Integer, zCon As String, zLin As String) As String
   On Error GoTo err
   
   Dim zz As Integer, zCod As Long, zIns As Integer, zE_s As String, _
       zMon As String, zApo As Currency, zMes As String

   zCod = 0: zIns = 0: zMes = ""
   zz = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
   If zz = 0 Then
      MsgBox "Codigo de Socio " + Str(zSoc) + " No Existe", vbExclamation
      Exit Function
   End If
   zCod = ADO8!codigo
   zIns = ADO8!ins
   zE_s = ADO8!e_socio
   Set ADO8 = Nothing

   zz = Leerado8("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + zE_s + "' ")
   If zz > 0 Then
      zMon = ADO8!moneda
      zApo = ADO8!aporte
   End If
   Set ADO8 = Nothing
      
   Select Case zCon
   Case "01", "02"
        zMes = wanocia + "/01"
        zz = Leerado8("SELECT * FROM CTASXCAB " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                & "       CONCEPTO = '" + zCon + "' AND " _
                & "       SDONEW > 0 " _
                & " ORDER BY MES ")
        If zz > 0 Then
           ADO8.MoveFirst
           zMes = ADO8!mes
        End If
   
        Do While True
           zz = Leerado8("SELECT * FROM TMP_COBRODET " _
                & " WHERE    USU = '" + wcodusu + "' AND " _
                & "       MESCOB = '" + zMes + "' AND " _
                & "       LINCOB <> '" + zLin + "' ")
           If zz = 0 Then
              Exit Do
           End If
           If Mid(zMes, 6, 2) = "12" Then
              zMes = Format(Val(Mid(zMes, 1, 4)) + 1, "0000") + "/" + "01"
           Else
              zMes = Mid(zMes, 1, 4) + "/" + Format(Val(Mid(zMes, 6, 2)) + 1, "00")
           End If
        Loop
   Case "03"
      zz = Leerado4a("select c.NUMERO ,c.CODSOCIO, c.CODIGO, c.moneda, c.ins, d.linea, D.VCMTO, D.CARGOS, D.ABONOS, D.SDONEW " _
                    & " from FRACCAB as c INNER JOIN FRACDET AS D " _
                    & "   ON C.NUMERO = D.NUMERO " _
                    & " where CODIGO = " + Str(zCod) + " AND " _
                    & "       D.SDONEW > 0 AND " _
                    & "       D.VCMTO < '" + Format(Date, "dd/mm/yyyy") + "'")
      If zz > 0 Then
         ADO4a.MoveFirst
         zMes = Format(Year(ADO4a!vcmto), "0000") + "/" + Format(Month(ADO4a!vcmto), "00")
      End If
   End Select
   
   BuscaUltimoMes = zMes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaUltimoApo(zSoc As Integer, zMes As String, zCon As String) As Currency
   On Error GoTo err

   Dim zz As Integer, zCod As Long, zIns As Integer, zE_s As String, _
       zMon As String, zApo As Currency

   If Len(Trim(zMes)) = 0 Then
      BuscaUltimoApo = 0
      Exit Function
   End If

   zCod = 0: zIns = 0
   zz = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
   If zz = 0 Then
      MsgBox "Codigo de Socio " + Str(zSoc) + " No Existe", vbExclamation
      Exit Function
   End If
   zCod = ADO8!codigo
   zIns = ADO8!ins
   zE_s = ADO8!e_socio
   Set ADO8 = Nothing
   zMon = ""
   zApo = 0

   zz = Leerado8("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + zE_s + "' ")
   If zz > 0 Then
      Select Case zCon
      Case "01"
           zMon = ADO8!moneda
           zApo = ADO8!aporte
      Case "02"
           zMon = ADO8!monren
           zApo = ADO8!renova
      End Select
   End If
   Set ADO8 = Nothing
      
'   Select Case zCon
'   Case "01", "02"
'        zz = Leerado8("SELECT * FROM CTASXCAB " _
'                & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
'                & "            MES = '" + Trim(zMes) + "' AND " _
'                & "       CONCEPTO = '" + zCon + "' ")
'        If zz = 0 Then
'           zApo = CreaAporteMes(zSoc, zMes, zCon, 2)
'        Else
'           If ADO8!sdonew > 0 Then
'              zApo = ADO8!sdonew
'           Else
'              zApo = 0
'           End If
'        End If
'   Case "03"
'        zz = Leerado4a("select c.NUMERO ,c.CODSOCIO, c.CODIGO, c.moneda, c.ins, d.linea, D.VCMTO, D.CARGOS, D.ABONOS, D.SDONEW " _
'                    & " from FRACCAB as c INNER JOIN FRACDET AS D " _
'                    & "   ON C.NUMERO = D.NUMERO " _
'                    & " where CODIGO = " + Str(zCod) + " AND " _
'                    & "       D.SDONEW > 0 AND " _
'                    & "       D.VCMTO < '" + Format(Date, "dd/mm/yyyy") + "'")
'        If zz > 0 Then
'           ADO4a.MoveFirst
'           zApo = ADO4a!cargos
'        End If
'        Set ADO4a = Nothing
'   End Select
      
   BuscaUltimoApo = zApo
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaUltimoMes128(zSoc As Integer, zCon As String, zLin As String) As String
   On Error GoTo err
   
   Dim zz As Integer, zCod As Long, zIns As Integer, zE_s As String, _
       zMon As String, zApo As Currency, zMes As String

   zCod = 0: zIns = 0: zMes = ""
   zz = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
   If zz = 0 Then
      MsgBox "Codigo de Socio " + Str(zSoc) + " No Existe", vbExclamation
      BuscaUltimoMes128 = ""
      Exit Function
   End If
   zCod = ADO8!codigo
   zIns = ADO8!ins
   zE_s = ADO8!e_socio
   Set ADO8 = Nothing

   zz = Leerado8("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + zE_s + "' ")
   If zz > 0 Then
      zMon = ADO8!moneda
      zApo = ADO8!aporte
   End If
   Set ADO8 = Nothing
      
   zz = Leerado4a("select c.NUMERO ,c.CODSOCIO, c.CODIGO, c.moneda, c.ins, d.linea, D.VCMTO, D.CARGOS, D.ABONOS, D.SDONEW " _
                 & " from FRACCAB as c INNER JOIN FRACDET AS D " _
                 & "   ON C.NUMERO = D.NUMERO " _
                 & " where CODIGO = " + Str(zCod) + " AND " _
                 & "       D.SDONEW > 0 AND " _
                 & "       D.VCMTO < '" + Format(Date, "dd/mm/yyyy") + "'")
   If zz > 0 Then
      ADO4a.MoveFirst
      zMes = Format(Year(ADO4a!vcmto), "0000") + "/" + Format(Month(ADO4a!vcmto), "00")
   End If
   
   BuscaUltimoMes128 = zMes
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaUltimoApo128(zSoc As Integer, zMes As String, zCon As String, sw As Integer) As Variant
   On Error GoTo err

   Dim zz As Integer, zCod As Long, zIns As Integer, zE_s As String, _
       zMon As String, zApo As Currency, zNumFra As String, zLinFra As String

   If Len(Trim(zMes)) = 0 Then
      BuscaUltimoApo128 = 0
      Exit Function
   End If

   zCod = 0: zIns = 0
   zz = Leerado8("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
   If zz = 0 Then
      MsgBox "Codigo de Socio " + Str(zSoc) + " No Existe", vbExclamation
      Exit Function
   End If
   zCod = ADO8!codigo
   zIns = ADO8!ins
   zE_s = ADO8!e_socio
   Set ADO8 = Nothing
   zMon = ""
   zApo = 0: zNumFra = "": zLinFra = ""

   zz = Leerado8("SELECT * FROM MAEE_SOCIO WHERE E_SOCIO = '" + zE_s + "' ")
   If zz > 0 Then
      Select Case zCon
      Case "01"
           zMon = ADO8!moneda
           zApo = ADO8!aporte
      Case "02"
           zMon = ADO8!monren
           zApo = ADO8!renova
      End Select
   End If
   Set ADO8 = Nothing
      
   zz = Leerado4a("select c.NUMERO ,c.CODSOCIO, c.CODIGO, c.moneda, c.ins, d.linea, D.VCMTO, D.CARGOS, D.ABONOS, D.SDONEW " _
               & " from FRACCAB as c INNER JOIN FRACDET AS D " _
               & "   ON C.NUMERO = D.NUMERO " _
               & " where CODIGO = " + Str(zCod) + " AND " _
               & "       D.SDONEW > 0 AND " _
               & "       D.VCMTO < '" + Format(Date, "dd/mm/yyyy") + "'")
   If zz > 0 Then
      ADO4a.MoveFirst
      zApo = ADO4a!cargos
      zNumFra = ADO4a!numero
      zLinFra = ADO4a!linea
   End If
   Set ADO4a = Nothing
      
   Select Case sw
   Case 1
        BuscaUltimoApo128 = zApo
   Case 2
        BuscaUltimoApo128 = zNumFra
   Case 3
        BuscaUltimoApo128 = zLinFra
   End Select
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaEnvioDieco(zSoc As Integer, zMes As String) As Currency
   On Error GoTo err
   
   Dim zz As Integer, zEnvio As Currency
   zEnvio = 0
   zz = Leerado8("SELECT * FROM DIECOCAB " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnvio = ADO8!netsocio
   End If
   Set ADO8 = Nothing

   zz = Leerado8("SELECT * FROM DIECOCAB " _
                & " WHERE CODASIG1 = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnvio = ADO8!netasig1
   End If
   Set ADO8 = Nothing

   zz = Leerado8("SELECT * FROM DIECOCAB " _
                & " WHERE CODASIG2 = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnvio = ADO8!netasig2
   End If
   Set ADO8 = Nothing

   zz = Leerado8("SELECT * FROM DIECOCAB " _
                & " WHERE CODASIG3 = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnvio = ADO8!netasig3
   End If
   Set ADO8 = Nothing

   zz = Leerado8("SELECT * FROM DIECOCAB " _
                & " WHERE CODASIG4 = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnvio = ADO8!netasig4
   End If
   Set ADO8 = Nothing

   zz = Leerado8("SELECT * FROM DIECOCAB " _
                & " WHERE CODASIG5 = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnvio = ADO8!netasig5
   End If
   Set ADO8 = Nothing

   BuscaEnvioDieco = zEnvio
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function BuscaEnvioCajMP(zSoc As Integer, zMes As String) As Currency
   On Error GoTo err
   
   Dim zz As Integer, zEnvio As Currency
   zEnvio = 0
   zz = Leerado8("SELECT * FROM CAJMPCAB " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnvio = ADO8!netsocio
   End If
   Set ADO8 = Nothing

   zz = Leerado8("SELECT * FROM CAJMPCAB " _
                & " WHERE CODASIG1 = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnvio = ADO8!netasig1
   End If
   Set ADO8 = Nothing

   zz = Leerado8("SELECT * FROM CAJMPCAB " _
                & " WHERE CODASIG2 = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnvio = ADO8!netasig2
   End If
   Set ADO8 = Nothing

   zz = Leerado8("SELECT * FROM CAJMPCAB " _
                & " WHERE CODASIG3 = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnvio = ADO8!netasig3
   End If
   Set ADO8 = Nothing

   zz = Leerado8("SELECT * FROM CAJMPCAB " _
                & " WHERE CODASIG4 = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnvio = ADO8!netasig4
   End If
   Set ADO8 = Nothing

   zz = Leerado8("SELECT * FROM CAJMPCAB " _
                & " WHERE CODASIG5 = " + Str(zSoc) + " AND " _
                & "            MES = '" + zMes + "' ")
   If zz > 0 Then
      zEnvio = ADO8!netasig5
   End If
   Set ADO8 = Nothing

   BuscaEnvioCajMP = zEnvio
   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Function CalTotApo(zSoc As Integer) As Currency
   On Error GoTo err
   
   Dim zz As Integer, zCod As Long, zIns As Integer, _
       zTotDie As Currency, zTotBco As Currency, zTotCMP As Currency, _
       zTotTes As Currency, zTotDev As Currency, zTotApo As Currency
   
   zTotDie = 0: zTotBco = 0: zTotCMP = 0: zTotTes = 0: zTotDev = 0: zTotApo = 0
   
   zz = Leerado6a("SELECT * FROM MAESOCIO WHERE CODSOCIO = " + Str(zSoc) + " ")
   If zz > 0 Then
      zCod = ADO6a!codigo
      zIns = ADO6a!ins
   End If
   Set ADO6a = Nothing
   
   If zCod <> 0 Then
      zz = Leerado6a("SELECT SUM(Z.MONTO) AS MONTO " _
             & " FROM ZZZ_MRECIBOS AS Z INNER JOIN ZZZ_CONCEPTO AS M " _
             & "   ON Z.CONCEPTO = M.CCONCE " _
             & " WHERE Z.CODIGO = " + Str(zCod) + " AND " _
             & "          Z.INS = " + Str(zIns) + " AND " _
             & "      (Z.MARCA2 <> 'A' OR Z.MARCA2 IS NULL) AND " _
             & "      (M.aporte = 1) ")
      If zz > 0 Then
         zTotTes = IIf(IsNull(ADO6a!monto), 0, ADO6a!monto)
      End If
      Set ADO6a = Nothing
   
      zz = Leerado6a("SELECT SUM(APORTE) AS APORTE " _
             & " FROM ZZZ_BCORECAU " _
             & " WHERE CODIGO = " + Str(zCod) + " AND " _
             & "          INS = " + Str(zIns) + " ")
      If zz > 0 Then
         zTotBco = IIf(IsNull(ADO6a!aporte), 0, ADO6a!aporte)
      End If
      Set ADO6a = Nothing
   
      zz = Leerado6a("SELECT SUM(IMPORTE) AS IMPORTE " _
             & " FROM ZZZ_DEVOL " _
             & " WHERE CODIGO = " + Str(zCod) + " AND " _
             & "          INS = " + Str(zIns) + " ")
      If zz > 0 Then
         zTotDev = IIf(IsNull(ADO6a!importe), 0, ADO6a!importe)
      End If
      Set ADO6a = Nothing
   
      zz = Leerado6a("SELECT SUM(IMPO01 + IMPO02 + IMPO03 + IMPO04 + " _
             & "                 IMPO05 + IMPO06 + IMPO07 + IMPO08 + " _
             & "                 IMPO09 + IMPO10 + IMPO11 + IMPO12) AS TOTAL " _
             & " FROM ZZZ_APOR_PLA " _
             & " WHERE  CODIGO = " + Str(zCod) + " AND " _
             & "           INS = " + Str(zIns) + " AND " _
             & "       TIPAPOR = '1' ")
      If zz > 0 Then
         zTotDie = IIf(IsNull(ADO6a!Total), 0, ADO6a!Total)
      End If
      Set ADO6a = Nothing
          
      zz = Leerado6a("SELECT SUM(IMPO01 + IMPO02 + IMPO03 + IMPO04 + " _
             & "                 IMPO05 + IMPO06 + IMPO07 + IMPO08 + " _
             & "                 IMPO09 + IMPO10 + IMPO11 + IMPO12) AS TOTAL " _
             & " FROM ZZZ_APOR_PLA " _
             & " WHERE  CODIGO = " + Str(zCod) + " AND " _
             & "           INS = " + Str(zIns) + " AND " _
             & "       TIPAPOR <> '1' ")
      If zz > 0 Then
         zTotCMP = IIf(IsNull(ADO6a!Total), 0, ADO6a!Total)
      End If
      Set ADO6a = Nothing
   End If
   
   zTotApo = zTotTes + zTotBco - zTotDev + zTotDie + zTotCMP
   
   CalTotApo = zTotApo

   Exit Function
err:
   MsgBox err.Description
   Resume Next

End Function

Public Sub CreateAporteAnoMes(zSoc As Integer, zAno As String, zFecIng As Date)
   On Error GoTo err
   
   Dim aa As Integer, zCod As Long, zIns As Integer, zMesIni As String, _
       zDia As Integer, zMes As Integer, zmmm As String, zFec As Date, _
       zApo As Currency, zMon As String, zE_s As String
   zDia = Day(zFecIng)
   zMes = Month(zFecIng)
   zAno = Year(zFecIng)
   
   aa = Leerado8a("SELECT S.CODSOCIO, S.CODIGO, S.INS, S.E_SOCIO, E.MONEDA, E.APORTE " _
                & " FROM MAESOCIO AS S INNER JOIN MAEE_SOCIO AS E " _
                & "   ON S.E_SOCIO = E.E_SOCIO " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " ")
   If aa > 0 Then
      zE_s = ADO8a!e_socio
      zMon = ADO8a!moneda
      zApo = ADO8a!aporte
   End If
   Set ADO8a = Nothing
   
   If zDia >= 20 Then
      If zMes = 12 Then
         zMesIni = Format(zAno + 1, "00") + "/" + "/01"
         zMes = 1
         zAno = zAno + 1
      Else
         zMesIni = Format(zAno, "0000") + "/" + Format(zMes + 1, "00")
         zMes = zMes + 1
      End If
   Else
      zMesIni = Format(zAno, "0000") + "/" + Format(zMes, "00")
   End If
   
   Dim II As Integer
   
   For II = zMes To 12
       zmmm = Format(II, "00")
       zFec = Format("01/" + zmmm + "/" + Format(zAno, "0000"), "dd/mm/yyyy")
          
       aa = Leerado6a("SELECT * FROM CTASXCAB " _
                  & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                  & "            MES = '" + Format(zAno, "0000") + "/" + zmmm + "' AND " _
                  & "       CONCEPTO = '01' ")
       If aa = 0 Then
          Db.BeginTrans
          Db.Execute ("INSERT INTO CTASXCAB " _
          & " (CODSOCIO, MES, CONCEPTO, E_SOCIO, MONEDA, CARGOS, ABONOS, SDONEW ) " _
          & " VALUES " _
          & " (" + Str(zSoc) + ", '" + Format(zAno, "0000") + "/" + zmmm + "', '01', '" + zE_s + "', '" + zMon + "', " _
          & "  " + Str(zApo) + ", 0, " + Str(zApo) + " ) ")
          Db.CommitTrans
       End If
                
       aa = Leerado6a("SELECT * FROM CTASXDET " _
                  & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
                  & "            MES = '" + Format(zAno, "0000") + "/" + zmmm + "' AND " _
                  & "       CONCEPTO = '01' ")
       If aa = 0 Then
          Db.BeginTrans
          Db.Execute ("INSERT INTO CTASXDET " _
          & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, TIPMOV, FECHA, " _
          & "  DOLARE, SOLESS, SDOOLD, CARGOS, ABONOS, SDONEW) " _
          & " VALUES " _
          & " (" + Str(zSoc) + ", '" + Format(zAno, "0000") + "/" + zmmm + "', '01', '00', '', '', '', '1', " _
          & "  '" + Format(zFec, "dd/mm/yyyy") + "', 0, 0, 0, " + Str(zApo) + ", " _
          & "  0, " + Str(zApo) + " ) ")
          Db.CommitTrans
       End If
   Next
   Exit Sub
err:
   MsgBox err.Description
   Resume Next
End Sub

Public Function BuscaUltimoDiecoCajMP(zSoc As Integer, sw As Byte) As Variant
   On Error GoTo err
   
   Dim zz As Integer, _
       zMesEnvio As String, zMesRecibe As String, _
       zImpEnvio As Currency, zImpRecibe As Currency
   
   zz = Leerado8a("SELECT * FROM DIECOCAB " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " " _
                & " ORDER BY MES DESC ")
   If zz > 0 Then
      zMesEnvio = ADO8a!mes
      zMesRecibe = ADO8a!mes
      zImpEnvio = ADO8a!netsocio
      zImpRecibe = ADO8a!dscsocio
   End If
   Set ADO8a = Nothing

   zz = Leerado8a("SELECT * FROM CAJMPCAB " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " " _
                & " ORDER BY MES DESC ")
   If zz > 0 Then
      If zMesEnvio < ADO8a!mes Then
         zMesEnvio = ADO8a!mes
         zMesRecibe = ADO8a!mes
         zImpEnvio = ADO8a!netsocio
         zImpRecibe = ADO8a!dscsocio
      End If
   End If
   Set ADO8a = Nothing

   Select Case sw
   Case 1
        BuscaUltimoDiecoCajMP = zMesEnvio
   Case 2
        BuscaUltimoDiecoCajMP = zImpEnvio
   Case 3
        BuscaUltimoDiecoCajMP = zMesRecibe
   Case 4
        BuscaUltimoDiecoCajMP = zImpRecibe
   End Select

   Exit Function
err:
   MsgBox err.Description
   Resume Next
End Function

Public Sub DistribuyeDieco(zAno As String, zMes As String, zFec As Date, zSoc As Integer, zDsc As Currency)
   On Error GoTo err
   
   Dim zz As Integer, zqqq As Variant, zCod As Long, zIns As Integer, _
       zNom As String
            
   zCod = 0: zIns = 0: zNom = ""
   zz = Leerado5a("SELECT * FROM MAESOCIO " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " ")
   If zz = 0 Then
      MsgBox "Socio " + Str(zSoc) + " No Existe En Maestro"
      Exit Sub
   End If
   zCod = ADO5a!codigo
   zIns = ADO5a!ins
   zNom = Trim(ADO5a!nombre)
   Set ADO5a = Nothing
            
   zz = Leerado5a("SELECT * FROM CTASXCAB " _
               & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
               & "            MES = '" + zAno + "/" + zMes + "' AND " _
               & "       CONCEPTO = '01' ")
   If zz = 0 Then
      zqqq = CreaAporteMes(zSoc, zAno + "/" + zMes, "01", 1)
   End If
   Set ADO5a = Nothing

   zz = Leerado5a("SELECT * FROM CTASXDET " _
               & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
               & "            MES = '" + zAno + "/" + zMes + "' AND " _
               & "       CONCEPTO = '01' AND " _
               & "         TIPMOV = '2' AND " _
               & "         TIPCOB = '01' AND " _
               & "         SERCOB = '001' AND " _
               & "         NUMCOB = '" + Right(zAno, 2) + zMes + "00001' AND " _
               & "         LINCOB = '0001' ")
   If zz = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO CTASXDET " _
      & " (CODSOCIO, MES, CONCEPTO, TIPCOB, SERCOB, NUMCOB, LINCOB, " _
      & "  TIPMOV, FECHA, TIPCAM, DOLARE, SOLESS, SDOOLD, CARGOS, " _
      & "  ABONOS, SDONEW, OBS ) " _
      & " VALUES " _
      & " (" + Str(zSoc) + ", '" + zAno + "/" + zMes + "', '01', '01', '001', " _
      & "  '" + Right(zAno, 2) + zMes + "00001', '0001', " _
      & "  '2', '" + Format(zFec, "dd/mm/yyyy") + "', 0, 0, " + Str(zDsc) + ", " _
      & "  0, 0, " + Str(zDsc) + ", 0, '' ) ")
      Db.CommitTrans
   Else
      Db.BeginTrans
      Db.Execute ("UPDATE CTASXDET " _
      & " SET SOLESS = " + Str(zDsc) + ", ABONOS = " + Str(zDsc) + " " _
      & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
      & "            MES = '" + zAno + "/" + zMes + "' AND " _
      & "       CONCEPTO = '01' AND " _
      & "         TIPMOV = '2' AND " _
      & "         TIPCOB = '01' AND " _
      & "         SERCOB = '001' AND " _
      & "         NUMCOB = '" + Right(zAno, 2) + zMes + "00001' AND " _
      & "         LINCOB = '0001' ")
      Db.CommitTrans
   End If
            
   Call ActualizaSaldos(zSoc, zAno + "/" + zMes, "01")
                        
   zz = Leerado7a("SELECT * FROM ZZZ_APOR_PLA " _
           & " WHERE  CODIGO = " + Str(zCod) + " AND " _
           & "           INS = " + Str(zIns) + " AND " _
           & "        CUOANO = '" + zAno + "' AND " _
           & "       TIPAPOR = '1' ")
   If zz = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO ZZZ_APOR_PLA " _
      & " (CODIGO, INS, NOMBRE, CUOANO, TIPAPOR, " _
      & "  IMPO01, IMPO02, IMPO03, IMPO04, IMPO05, IMPO06, IMPO07, IMPO08, IMPO09, IMPO10, IMPO11, IMPO12, " _
      & "  TOTIMPO, DEUDA_PT2, PASA_PLA ) " _
      & " VALUES " _
      & " (" + Str(zCod) + ", " + Str(zIns) + ", '" + zNom + "', '" + zAno + "', '1', " _
      & "  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '' ) ")
      Db.CommitTrans
   End If
               
   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_APOR_PLA " _
   & " SET IMPO" + zMes + " = " + Str(zDsc) + " " _
   & " WHERE  CODIGO = " + Str(zCod) + " AND " _
   & "           INS = " + Str(zIns) + " AND " _
   & "        CUOANO = '" + zAno + "' AND " _
   & "       TIPAPOR = '1' ")
   Db.CommitTrans
            
   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_APOR_PLA " _
   & " SET TOTIMPO = IMPO01 + IMPO02 + IMPO03 + IMPO04 + IMPO05 + IMPO06 + " _
   & "               IMPO07 + IMPO08 + IMPO09 + IMPO10 + IMPO11 + IMPO12 " _
   & " WHERE  CODIGO = " + Str(zCod) + " AND " _
   & "           INS = " + Str(zIns) + " AND " _
   & "        CUOANO = '" + zAno + "' AND " _
   & "       TIPAPOR = '1' ")
   Db.CommitTrans
         
   Exit Sub
err:
   MsgBox err.Description
   Resume Next
End Sub

Public Sub DelProcesoDieco(zAno As String, zMes As String, zFec As Date, zSoc As Integer)
   On Error GoTo err
   
   Dim zz As Integer, zqqq As Variant, zCod As Long, zIns As Integer, _
       zNom As String
            
   zCod = 0: zIns = 0: zNom = ""
   zz = Leerado5a("SELECT * FROM MAESOCIO " _
                & " WHERE CODSOCIO = " + Str(zSoc) + " ")
   If zz = 0 Then
      MsgBox "Socio " + Str(zSoc) + " No Existe En Maestro"
      Exit Sub
   End If
   zCod = ADO5a!codigo
   zIns = ADO5a!ins
   zNom = Trim(ADO5a!nombre)
   Set ADO5a = Nothing
            
   zz = Leerado5a("SELECT * FROM CTASXCAB " _
               & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
               & "            MES = '" + zAno + "/" + zMes + "' AND " _
               & "       CONCEPTO = '01' ")
   If zz = 0 Then
      zqqq = CreaAporteMes(zSoc, zAno + "/" + zMes, "01", 1)
   End If
   Set ADO5a = Nothing

   Db.BeginTrans
   Db.Execute ("DELETE FROM CTASXDET " _
   & " WHERE CODSOCIO = " + Str(zSoc) + " AND " _
   & "            MES = '" + zAno + "/" + zMes + "' AND " _
   & "       CONCEPTO = '01' AND " _
   & "         TIPMOV = '2' AND " _
   & "         TIPCOB = '01' AND " _
   & "         SERCOB = '001' AND " _
   & "         NUMCOB = '" + Right(zAno, 2) + zMes + "00001' AND " _
   & "         LINCOB = '0001' ")
   Db.CommitTrans

   Call ActualizaSaldos(zSoc, zAno + "/" + zMes, "01")

   zz = Leerado7a("SELECT * FROM ZZZ_APOR_PLA " _
           & " WHERE  CODIGO = " + Str(zCod) + " AND " _
           & "           INS = " + Str(zIns) + " AND " _
           & "        CUOANO = '" + zAno + "' AND " _
           & "       TIPAPOR = '1' ")
   If zz = 0 Then
      Db.BeginTrans
      Db.Execute ("INSERT INTO ZZZ_APOR_PLA " _
      & " (CODIGO, INS, NOMBRE, CUOANO, TIPAPOR, " _
      & "  IMPO01, IMPO02, IMPO03, IMPO04, IMPO05, IMPO06, IMPO07, IMPO08, IMPO09, IMPO10, IMPO11, IMPO12, " _
      & "  TOTIMPO, DEUDA_PT2, PASA_PLA ) " _
      & " VALUES " _
      & " (" + Str(zCod) + ", " + Str(zIns) + ", '" + zNom + "', '" + zAno + "', '1', " _
      & "  0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, '' ) ")
      Db.CommitTrans
   End If

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_APOR_PLA " _
   & " SET IMPO" + zMes + " = 0 " _
   & " WHERE  CODIGO = " + Str(zCod) + " AND " _
   & "           INS = " + Str(zIns) + " AND " _
   & "        CUOANO = '" + zAno + "' AND " _
   & "       TIPAPOR = '1' ")
   Db.CommitTrans

   Db.BeginTrans
   Db.Execute ("UPDATE ZZZ_APOR_PLA " _
   & " SET TOTIMPO = IMPO01 + IMPO02 + IMPO03 + IMPO04 + IMPO05 + IMPO06 + " _
   & "               IMPO07 + IMPO08 + IMPO09 + IMPO10 + IMPO11 + IMPO12 " _
   & " WHERE  CODIGO = " + Str(zCod) + " AND " _
   & "           INS = " + Str(zIns) + " AND " _
   & "        CUOANO = '" + zAno + "' AND " _
   & "       TIPAPOR = '1' ")
   Db.CommitTrans

   Exit Sub
err:
   MsgBox err.Description
   Resume Next
End Sub

