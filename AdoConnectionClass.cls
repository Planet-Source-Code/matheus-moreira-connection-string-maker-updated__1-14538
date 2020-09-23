VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AdoConnectionClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

Public Enum pdsaProvider
    pdsajet = 1
    pdsasqlserver = 2
    pdsajet40 = 3
    pdsaoracle = 4
    pdsaodbc = 5
    pdsadbase = 6
    pdsaexcel = 7
End Enum

Dim moConn As ADODB.Connection

Dim mAppRole As String
Dim mAppRolePassword As String
Dim mCursorLocation As CursorLocationEnum
Dim mDataSource As String
Dim mInitialCatalog As String
Dim mPassword As String
Dim mProvider As String
Dim mProviderConst As pdsaProvider
Dim mUDLFileName As String
Dim mUseNTSecurity As Boolean
Dim mUserID As String
Event Connected()
Event Disconnected()


Public Property Get ProviderConst() As pdsaProvider
    ProviderConst = mProviderConst
End Property
Public Property Let ProviderConst(NewI As pdsaProvider)
    mProviderConst = NewI
    Select Case NewI
        Case 1 'Jet 3.51
            mProvider = "Microsoft.Jet.OLEDB.3.51"
        Case 2 'SQL Server
            mProvider = "sqloledb"
        Case 3 'Jet 4.0
            mProvider = "Microsoft.Jet.OLEDB.4.0"
        Case 4 'Oracle
            mProvider = "msdaora"
        Case 5 'ODBC
            mProvider = "msdasql.1"
        Case 6 'Dbase
            mProvider = "Microsoft.Jet.OLEDB.4.0" ';Extended Properties=dBASE III
        Case 7 'Excel 8.0
            mProvider = "Microsoft.Jet.OLEDB.4.0" ';Extended Properties=Excel 8.0
        Case Else 'Jet 3.51
            mProvider = "Microsoft.Jet.OLEDB.3.51"
            mProviderConst = 1
    End Select
End Property
Public Property Get AppRole() As String
    AppRole = mAppRole
End Property
Public Property Let AppRole(ByVal NewI As String)
    mAppRole = NewI
End Property
Public Property Get AppRolePassword() As String
    AppRolePassword = mAppRolePassword
End Property
Public Property Let AppRolePassword(ByVal NewI As String)
    mAppRolePassword = NewI
End Property
Public Property Get CursorLocation() As CursorLocationEnum
    CursorLocation = mCursorLocation
End Property
Public Property Let CursorLocation(ByVal NewI As CursorLocationEnum)
    mCursorLocation = NewI
End Property
Public Property Get DataSource() As String
    DataSource = mDataSource
End Property
Public Property Let DataSource(ByVal NewI As String)
    mDataSource = NewI
End Property
Public Property Get InitialCatalog() As String
    InitialCatalog = mInitialCatalog
End Property
Public Property Let InitialCatalog(ByVal NewI As String)
    mInitialCatalog = NewI
End Property
Public Property Get Password() As String
    Password = mPassword
End Property
Public Property Let Password(ByVal NewI As String)
    mPassword = NewI
End Property
Public Property Get Provider() As String
    Provider = mProvider
End Property
Public Property Let Provider(ByVal NewI As String)
    mProvider = NewI
End Property
Public Property Get UDLFileName() As String
    UDLFileName = mUDLFileName
End Property
Public Property Let UDLFileName(ByVal NewI As String)
    mUDLFileName = NewI
End Property
Public Property Get UseNTSecurity() As Boolean
    UseNTSecurity = mUseNTSecurity
End Property
Public Property Let UseNTSecurity(ByVal NewI As Boolean)
    mUseNTSecurity = NewI
End Property
Public Property Get UserID() As String
    UserID = mUserID
End Property
Public Property Let UserID(ByVal NewI As String)
    mUserID = NewI
End Property
Public Property Get Connection() As ADODB.Connection
    Set Connection = moConn
End Property
Public Function DataOpen() As Boolean
On Error GoTo erro
If moConn.State = adStateOpen Then
    moConn.Close
End If
moConn.CursorLocation = mCursorLocation
moConn.ConnectionString = Me.ConnectionString
moConn.Mode = adModeReadWrite
moConn.Open

On Error GoTo 0
Call AppRoleSet
RaiseEvent Connected
DataOpen = True
sair:
Exit Function
erro:
RaiseEvent Disconnected
Err.Raise vbObjectError + 1, "AdoConnectionClass", Err.Description, Err.HelpFile, Err.HelpContext
DataOpen = False
Resume sair
End Function
Public Function ConnectionString()
Dim strRet As String
If mUDLFileName = "" Then
    strRet = "Provider=" & mProvider & _
             ";Data Source=" & mDataSource
Else
    strRet = "File Name=" & mUDLFileName
End If

Select Case mProviderConst
    Case pdsajet, pdsajet40
        'Nenhum Código extra necessário
        '<No extra code necessary>
        If mPassword <> "" Then strRet = strRet & ";Jet OLEDB:Database Password=" & mPassword
    
    Case pdsasqlserver
        If mInitialCatalog <> "" Then
            strRet = strRet & ";Initial Catalog = " & mInitialCatalog
        End If
        If mUseNTSecurity Then
            strRet = strRet & ";Integrated Security=SSPI"
        Else
            If mUserID <> "" Then
                strRet = strRet & ";User ID=" & mUserID
                If mPassword <> "" Then strRet = strRet & ";Password=" & mPassword
            End If
        End If
    Case pdsaoracle
        'Nenhum Código extra necessário
        '<No extra code necessary>
    Case pdsaodbc
        'Nenhum Código extra necessário
        '<No extra code necessary>
    Case pdsadbase
        strRet = strRet & ";Extended Properties=dBASE III;"
    Case pdsaexcel
        strRet = strRet & ";Extended Properties=Excel 8.0;"
        
End Select

ConnectionString = strRet
End Function
Private Function AppRoleSet()
Dim SQL As String
On Error GoTo erro

'Somente ira executar se for SQLServer.
'Nota: Esse procedimento soh functiona para o SQL Server 7.0,
'para a versão 6.5 ou anterior deixe a propriedade AppRole em branco.

'<This routine only will be executed if this are a SQLServer connection.>
'<Note: This procedure will only work for SQL Server 7.0.>
'<to version 6.5 or below leave the property AppRole blank.>
If mProviderConst = pdsasqlserver And mAppRole <> "" Then
    SQL = "EXEC sp_setAppRole '" & mAppRole & "', '" & mAppRolePassword & "'"
    moConn.Execute SQL
End If
    
sair:
Exit Function
erro:
Err.Raise Err.Number, "AdoConnectionClass", Err.Description, Err.HelpFile, Err.HelpContext
Resume sair

End Function
Private Sub Class_Initialize()
Set moConn = New ADODB.Connection
mCursorLocation = adUseServer
moConn.CursorLocation = adUseServer
ProviderConst = pdsajet
End Sub
Private Sub Class_Terminate()
Set moConn = Nothing
End Sub
Public Function OpenAccess(ByVal Database As String, Optional Password As String) As Boolean
ProviderConst = pdsajet40
DataSource = Database
InitialCatalog = ""
UserID = ""
mPassword = Password
OpenAccess = DataOpen()

End Function
Public Function OpenSQLServer(ByVal ServerName As String, Database As String, Optional ByVal UserName As String, Optional ByVal Password As String, Optional ByVal NTSecurity As Boolean) As Boolean
ProviderConst = pdsasqlserver
DataSource = ServerName
InitialCatalog = Database
If NTSecurity Then
    UseNTSecurity = True
Else
    UseNTSecurity = False

    UserID = UserName
    Password = Password
End If
OpenSQLServer = DataOpen()

End Function
