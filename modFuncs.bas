Attribute VB_Name = "modFuncs"
'+---------------------------------------------------+
'|                                                   |
'|  Create a MySql ODBC Connection and delete any    |
'|  ODBC Data Source programmatically                |
'|                                                   |
'+---------------------------------------------------+

Option Explicit

'================================
' Registry API Definitions
'================================
Private Declare Function RegCloseKey Lib "advapi32.dll" _
                        (ByVal hKey As Long) As Long
                        
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias _
    "RegCreateKeyExA" (ByVal hKey As Long, _
                        ByVal lpSubKey As String, _
                        ByVal Reserved As Long, _
                        ByVal lpClass As Long, _
                        ByVal dwOptions As Long, _
                        ByVal samDesired As Long, _
                        ByVal lpSecurityAttributes As Long, _
                        phkResult As Long, _
                        lpdwDisposition As Long) As Long
                        
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias _
    "RegOpenKeyExA" (ByVal hKey As Long, _
                        ByVal lpSubKey As String, _
                        ByVal ulOptions As Long, _
                        ByVal samDesired As Long, _
                        phkResult As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias _
    "RegSetValueExA" (ByVal hKey As Long, _
                        ByVal lpValueName As String, _
                        ByVal Reserved As Long, _
                        ByVal dwType As Long, _
                        lpData As Any, _
                        ByVal cbData As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias _
    "RegDeleteKeyA" (ByVal hKey As Long, _
                        ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias _
    "RegDeleteValueA" (ByVal hKey As Long, _
                        ByVal lpValueName As String) As Long
                        
'================================
' Registry API Constants
'================================
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7
Const ERROR_MORE_DATA = 234
Const KEY_WRITE = &H20006
Const KEY_READ = &H20019
Const KEY_QUERY_VALUE = &H1
Const REG_OPENED_EXISTING_KEY = &H2
Const KEY_ENUMERATE_SUB_KEYS = &H8
Const ERROR_SUCCESS = 0&

'================================
' Registry API Root Key Constants
'================================
Const HKEY_CLASSES_ROOT = &H80000000
Const HKEY_CURRENT_USER = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const HKEY_USERS = &H80000003
Const HKEY_CURRENT_CONFIG = &H8000000

'================================
' Private Variables
'================================
Private lRootKey As Long
Private lHandle As Long

'================================
' API ODBC Definitions
'================================
Private Declare Function SQLAllocEnv Lib "odbc32.dll" (env As Long) As Integer

Private Declare Function SQLDataSources Lib "odbc32.dll" (ByVal henv As Long, _
                    ByVal fdir As Integer, ByVal szDSN As String, _
                    ByVal cbDSNMAx As Integer, pcbDSN As Integer, _
                    ByVal szDesc As String, ByVal cbDescMax As Integer, _
                    pcbDesc As Integer) As Integer

'==================================================================
' Creates an ODBC Entry
'==================================================================
Public Function CreateMySqlODBC(strODBCName As String, _
                            strMySqlServer As String, _
                            strMySqlUser As String, _
                            strMySqlPassword As String, _
                            strMySqlDB As String, _
                            strMySqlDescription As String, _
                            Optional strMySqlDriver As String = "C:\Windows\System32\myodbc3.dll", _
                            Optional strMySqlOption As String = "3", _
                            Optional strMySqlPort As String = "3306", _
                            Optional strMySqlStmt As String = "") As Boolean
    
    On Local Error GoTo modError
    
    Dim ODBCRootKey As String
    
    lRootKey = HKEY_LOCAL_MACHINE                           ' Set the Registry Root Key
    
    ODBCRootKey = "Software\ODBC\ODBC.INI\" & strODBCName   ' Set the ODBC.INI Key
    CreateRegistryKey ODBCRootKey                           ' Create Entry in ODBC.INI
    
    '==================================================================
    ' Set MySql Registry Values
    '==================================================================
    SetRegistryValue ODBCRootKey, "Database", strMySqlDB
    SetRegistryValue ODBCRootKey, "Description", strMySqlDescription
    SetRegistryValue ODBCRootKey, "Driver", strMySqlDriver
    SetRegistryValue ODBCRootKey, "Option", strMySqlOption
    SetRegistryValue ODBCRootKey, "Password", strMySqlPassword
    SetRegistryValue ODBCRootKey, "Port", strMySqlPort
    SetRegistryValue ODBCRootKey, "Server", strMySqlServer
    SetRegistryValue ODBCRootKey, "Stmt", strMySqlStmt
    SetRegistryValue ODBCRootKey, "User", strMySqlUser
    
    '==================================================================
    ' Create an Entry in the INI file so the ODBC Connection can be
    ' selected from the ODBC Connection Manager in Windows or a Program
    '==================================================================
    SetRegistryValue "Software\ODBC\ODBC.INI\ODBC Data Sources", _
                        strODBCName, "MySQL ODBC 3.51 Driver"
    
    CreateMySqlODBC = True
    
    Exit Function
    
modError:
    
    MsgBox Err.Number & "-" & Err.Description, vbExclamation
    CreateMySqlODBC = False

End Function
'==================================================================
' Deletes an ODBC Entry
'==================================================================
Public Sub DeleteMySqlODBC(ODBCName As String)

    On Local Error GoTo modError
    
    lRootKey = HKEY_LOCAL_MACHINE    ' Set the Registry Root Key
    
    DeleteRegistryKey "Software\ODBC\ODBC.INI\" & ODBCName
    DeleteRegistryValue "Software\ODBC\ODBC.INI\ODBC Data Sources", ODBCName

    Exit Sub
    
modError:
    MsgBox Err.Number & "-" & Err.Description, vbExclamation
    
End Sub

'==================================================================
' Creates or Sets a Registry Value
'==================================================================
Private Function SetRegistryValue(ByVal KeyName As String, _
                                 ByVal valuename As String, _
                                 ByVal Value As Variant) As Boolean
    On Local Error GoTo modError

    Dim lValue As Long
    Dim sValue As String
    Dim binValue() As Byte
    Dim lLength As Long
    Dim lRetVal As Long
        
    If RegOpenKeyEx(lRootKey, KeyName, 0, KEY_WRITE, lHandle) Then
        Exit Function
    End If

    Select Case VarType(Value)
        Case vbInteger, vbLong
            lValue = Value
            lRetVal = RegSetValueEx(lHandle, valuename, 0, REG_DWORD, lValue, 4)
        Case vbString
            sValue = Value
            lRetVal = RegSetValueEx(lHandle, valuename, 0, REG_SZ, ByVal sValue, _
                Len(sValue))
        Case vbArray + vbByte
            binValue = Value
            lLength = UBound(binValue) - LBound(binValue) + 1
            lRetVal = RegSetValueEx(lHandle, valuename, 0, REG_BINARY, _
                binValue(LBound(binValue)), lLength)
        Case Else
            RegCloseKey lHandle
            MsgBox lRootKey & KeyName & _
                        " Unsupported value type [" & Value & "] " & _
                        "for [" & valuename & "]", vbCritical
    End Select
    
    RegCloseKey lHandle
    SetRegistryValue = (lRetVal = 0)
    
    Exit Function

modError:
    
    MsgBox Err.Number & "-" & Err.Description, vbExclamation
    
End Function

'==================================================================
' Create a New Registry Key
'==================================================================
Private Function CreateRegistryKey(ByVal KeyName As String) As Boolean
    
    On Local Error GoTo modError
    
    Dim lDisposition As Long
    
    If CheckRegistryKey(KeyName) = True Then
        CreateRegistryKey = True
    Else
        If RegCreateKeyEx(lRootKey, KeyName, 0, 0, 0, 0, 0, _
                            lHandle, lDisposition) Then
            MsgBox "Unable to create the registry key " & KeyName, vbCritical
        Else
            CreateRegistryKey = (lDisposition = REG_OPENED_EXISTING_KEY)
            RegCloseKey lHandle
        End If
    End If
    
    Exit Function

modError:
    
    MsgBox Err.Number & "-" & Err.Description, vbExclamation
    
End Function

'==================================================================
' Checks to see if a Registry Key Already Exists
'==================================================================
Private Function CheckRegistryKey(ByVal KeyName As String) As Boolean
    
    On Local Error GoTo modError
        
    If RegOpenKeyEx(lRootKey, KeyName, 0, KEY_READ, lHandle) = 0 Then
        CheckRegistryKey = True
        RegCloseKey lHandle
    End If
    
    Exit Function

modError:
    
    MsgBox Err.Number & "-" & Err.Description, vbExclamation
    
End Function
'==================================================================
' Deletes a Registry Key
'==================================================================
Public Sub DeleteRegistryKey(ByVal KeyName As String)
    
    Dim lRetVal As Long
    On Local Error GoTo modError

    lRetVal = RegDeleteKey(lRootKey, KeyName)
    Exit Sub

modError:
    
    MsgBox Err.Number & "-" & Err.Description, vbExclamation
    
End Sub

'==================================================================
' Deletes a Registry Value
'==================================================================
Public Function DeleteRegistryValue(ByVal KeyName As String, _
                                    ByVal valuename As String) As Boolean
    
    On Local Error GoTo modError
    
    If RegOpenKeyEx(lRootKey, KeyName, 0, KEY_WRITE, lHandle) Then Exit Function
    
    DeleteRegistryValue = (RegDeleteValue(lHandle, valuename) = 0)
    RegCloseKey lHandle
    
    Exit Function

modError:
    
    MsgBox Err.Number & "-" & Err.Description, vbExclamation

End Function

'==================================================================
' Get Current Datasources
'==================================================================
Public Sub GetDataSources(cmbCombo As ComboBox)
    
    On Local Error GoTo modError
    
    Dim strDataSrc As String
    Dim strDesc As String
    Dim intDataSrcLen As Integer
    Dim intDescLen As Integer
    Dim intRetCode As Integer
    Dim henv As Long
    
    cmbCombo.Clear
    
    '*********************************************************
    '* Use the Win32 API to Retrieve the Data Sources that
    '* are already defined.
    '*********************************************************
    If SQLAllocEnv(henv) <> -1 Then
        strDataSrc = String$(32, 32)
        strDesc = String$(255, 32)
        
        '*****************************************
        '* Locate the First Data Source
        '*****************************************
        intRetCode = SQLDataSources(henv, 2, strDataSrc, Len(strDataSrc), _
            intDataSrcLen, strDesc, Len(strDesc), intDescLen)
            
        While intRetCode = 0 Or intRetCode = 1
            '****************************************
            '* Add the First Entry
            '****************************************
            cmbCombo.AddItem Mid(strDataSrc, 1, intDataSrcLen)
            strDataSrc = String$(32, 32)
            strDesc = String$(255, 32)
            '****************************************
            '* Add remaining
            '****************************************
            intRetCode = SQLDataSources(henv, 1, strDataSrc, Len(strDataSrc), _
                intDataSrcLen, strDesc, Len(strDesc), intDescLen)
        Wend
    End If
    
    On Local Error GoTo 0
    
    Exit Sub
    
modError:
    
    MsgBox Err.Number & "-" & Err.Description, vbExclamation
    
End Sub
