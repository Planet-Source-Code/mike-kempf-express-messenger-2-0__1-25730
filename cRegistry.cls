VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'///////////////////////////////////////////////////////////////////////////////////////////
'   APPLICATION:        EXPRESS MESSENGER
'   DEVELOPED BY:       MICHAEL J. KEMPF
'   DATE:               JULY 15, 2001
'   CLASS NAME:         cRegistry
'   CLASS DESCRIPTION:  Provide easy access to the system registry
'   COPYRIGHT:          Copyright © 1999 - 2001, Kemtech Software
'///////////////////////////////////////////////////////////////////////////////////////////

' -----------------
' ADVAPI32
' -----------------
' function prototypes, constants, and type definitions
' for Windows 32-bit Registry API

' Registry API prototypes
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal Hkey As Long) As Long

Private Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" _
    (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" _
    (ByVal Hkey As Long, ByVal lpSubKey As String) As Long

Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" _
    (ByVal Hkey As Long, ByVal lpValueName As String) As Long

Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" _
    (ByVal Hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long

Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" _
    (ByVal Hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, _
    lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
    (ByVal Hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, _
    lpData As Any, ByVal cbData As Long) As Long

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
    (ByVal Hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, _
    ByVal samDesired As Long, phkResult As Long) As Long
    Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" _
    (ByVal Hkey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, _
    lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, _
    lpData As Any, lpcbData As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (dest As _
    Any, source As Any, ByVal numBytes As Long)
    
Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" _
    (ByVal Hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, _
    ByVal cbName As Long) As Long
    
Public Enum HKeys
    HKEY_CLASSES_ROOT = &H80000000
    HKEY_CURRENT_USER = &H80000001
    HKEY_LOCAL_MACHINE = &H80000002
    HKEY_USERS = &H80000003
    HKEY_PERFORMANCE_DATA = &H80000004
    HKEY_CURRENT_CONFIG = &H80000005
    HKEY_DYN_DATA = &H80000006
End Enum
    
Const ERROR_SUCCESS = 0&
Const REG_SZ = 1
Const REG_EXPAND_SZ = 2
Const REG_BINARY = 3
Const REG_DWORD = 4
Const REG_MULTI_SZ = 7
Const ERROR_MORE_DATA = 234
Const KEY_READ = &H20019

'///////////////////////////////////////////////////////////////////////////////////////////
'  PROCEDURE:    savekey
'
'  AUTHOR:       Michael J. Kempf        7/20/2001 12:22:07 PM
'
'  PURPOSE:      Save a specific KEY in the Registry
'
'  PARAMETERS:
'                Hkey      (Long) = HKEY_LOCAL_MACHINE,HKEY_CURRENT_USER...
'                strPath (String) = Path in the registry where key is located (ex. "Software\Microsft\Windows")
'
'///////////////////////////////////////////////////////////////////////////////////////////
Public Sub savekey(Hkey As HKeys, strPath As String)
Dim keyhand&
r = RegCreateKey(Hkey, strPath, keyhand&)
r = RegCloseKey(keyhand&)
End Sub
'///////////////////////////////////////////////////////////////////////////////////////////
'  FUNCTION:     getstring
'
'  AUTHOR:       Michael J. Kempf        7/20/2001 12:22:07 PM
'
'  PURPOSE:      Get a string value in the registry
'
'  PARAMETERS:
'       [in]     Hkey           (Long) = HKEY_LOCAL_MACHINE,HKEY_CURRENT_USER...
'       [in]     strPath      (String) = Path in the registry where key is located (ex. "Software\Microsft\Windows")
'       [in]     strValueName (String) = String Value to return
'
'///////////////////////////////////////////////////////////////////////////////////////////
Public Function getstring(Hkey As HKeys, strPath As String, strValue As String)

Dim keyhand As Long
Dim datatype As Long
Dim lResult As Long
Dim strBuf As String
Dim lDataBufSize As Long
Dim intZeroPos As Integer
r = RegOpenKey(Hkey, strPath, keyhand)
lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
If lValueType = REG_SZ Then
    strBuf = String(lDataBufSize, " ")
    lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
    If lResult = ERROR_SUCCESS Then
        intZeroPos = InStr(strBuf, Chr$(0))
        If intZeroPos > 0 Then
            getstring = Left$(strBuf, intZeroPos - 1)
        Else
            getstring = strBuf
        End If
    End If
End If
End Function

'///////////////////////////////////////////////////////////////////////////////////////////
'  PROCEDURE:    savestring
'
'  AUTHOR:       Michael J. Kempf        7/20/2001 12:22:02 PM
'
'  PURPOSE:      save a specific string value to the registry
'
'  PARAMETERS:
'                Hkey       (Long) = HKEY_LOCAL_MACHINE,HKEY_CURRENT_USER...
'                strPath  (String) = Path in the registry where key is located (ex. "Software\Microsft\Windows")
'                strValue (String) = Value name to save under
'                strdata  (String) = data to save
'
'///////////////////////////////////////////////////////////////////////////////////////////
Public Sub savestring(Hkey As HKeys, strPath As String, strValue As String, strdata As String)
Dim keyhand As Long
Dim r As Long
r = RegCreateKey(Hkey, strPath, keyhand)
r = RegSetValueEx(keyhand, strValue, 0, REG_SZ, ByVal strdata, Len(strdata))
r = RegCloseKey(keyhand)
End Sub

'///////////////////////////////////////////////////////////////////////////////////////////
'  FUNCTION:     getdword
'
'  AUTHOR:       Michael J. Kempf        7/20/2001 12:21:59 PM
'
'  PURPOSE:      Get a specific DWORD from the registry
'
'  PARAMETERS:
'       [in]     Hkey           (Long) = HKEY_LOCAL_MACHINE,HKEY_CURRENT_USER...
'       [in]     strPath      (String) = Path in the registry where key is located (ex. "Software\Microsft\Windows")
'       [in]     strValueName (String) = Name of the value to get the DWORD for..
'
'  RETURN:       Long = DWORD Value of the strValueName
'
'///////////////////////////////////////////////////////////////////////////////////////////
Function getdword(ByVal Hkey As HKeys, ByVal strPath As String, ByVal strValueName As String) As Long
Dim lResult As Long
Dim lValueType As Long
Dim lBuf As Long
Dim lDataBufSize As Long
Dim r As Long
Dim keyhand As Long

r = RegOpenKey(Hkey, strPath, keyhand)

 ' Get length/data type
lDataBufSize = 4
    
lResult = RegQueryValueEx(keyhand, strValueName, 0&, lValueType, lBuf, lDataBufSize)

If lResult = ERROR_SUCCESS Then
    If lValueType = REG_DWORD Then
        getdword = lBuf
    End If
'Else
'    Call errlog("GetDWORD-" & strPath, False)
End If

r = RegCloseKey(keyhand)
    
End Function

'///////////////////////////////////////////////////////////////////////////////////////////
'  FUNCTION:     SaveDword
'
'  AUTHOR:       Michael J. Kempf        7/20/2001 12:21:53 PM
'
'  PURPOSE:      Save a specific DWORD to the registry
'
'  PARAMETERS:
'       [in]     Hkey           (Long) = HKEY_LOCAL_MACHINE,HKEY_CURRENT_USER...
'       [in]     strPath      (String) = Path in the registry where key is located (ex. "Software\Microsft\Windows")
'       [in]     strValueName (String) = Name of the value to get the DWORD for..
'       [in]     lData          (Long) = DWORD data to save
'
'
'///////////////////////////////////////////////////////////////////////////////////////////
Function SaveDword(ByVal Hkey As HKeys, ByVal strPath As String, ByVal strValueName As String, ByVal lData As Long)
    Dim lResult As Long
    Dim keyhand As Long
    Dim r As Long
    r = RegCreateKey(Hkey, strPath, keyhand)
    lResult = RegSetValueEx(keyhand, strValueName, 0&, REG_DWORD, lData, 4)
    'If lResult <> error_success Then Call errlog("SetDWORD", False)
    r = RegCloseKey(keyhand)
End Function

'///////////////////////////////////////////////////////////////////////////////////////////
'  FUNCTION:     DeleteKey
'
'  AUTHOR:       Michael J. Kempf        7/20/2001 12:21:48 PM
'
'  PURPOSE:      Delete a specific Key from the registry
'
'  PARAMETERS:
'       [in]     Hkey     (Long) = HKEY_LOCAL_MACHINE,HKEY_CURRENT_USER...
'       [in]     strKey (String) = Key name to delete
'
'
'///////////////////////////////////////////////////////////////////////////////////////////
Public Function DeleteKey(ByVal Hkey As HKeys, ByVal strKey As String)
Dim r As Long
r = RegDeleteKey(Hkey, strKey)
End Function
'///////////////////////////////////////////////////////////////////////////////////////////
'  FUNCTION:     DeleteValue
'
'  AUTHOR:       Michael J. Kempf        7/20/2001 12:21:48 PM
'
'  PURPOSE:      Delete a specific value from the registry
'
'  PARAMETERS:
'       [in]     Hkey     (Long) = HKEY_LOCAL_MACHINE,HKEY_CURRENT_USER...
'       [in]     strPath (String) = Path in the registry where key is located (ex. "Software\Microsft\Windows")
'       [in]     strValue (String) = ket to delete
'
'
'///////////////////////////////////////////////////////////////////////////////////////////
Public Function DeleteValue(ByVal Hkey As HKeys, ByVal strPath As String, ByVal strValue As String)
Dim keyhand As Long
r = RegOpenKey(Hkey, strPath, keyhand)
r = RegDeleteValue(keyhand, strValue)
r = RegCloseKey(keyhand)
End Function

'///////////////////////////////////////////////////////////////////////////////////////////
'  FUNCTION:     EnumRegistryKeys
'
'  AUTHOR:       Michael J. Kempf        7/20/2001 12:17:39 PM
'
'  PURPOSE:      Enumerate registry keys under a given key
'
'  PARAMETERS:
'       [in]     Hkey      (Long) = HKEY_LOCAL_MACHINE,HKEY_CURRENT_USER...
'       [in]     KeyName (String) = Key name to enumerate
'
'  RETURN:       Collection -  returns a collection of strings
'
'///////////////////////////////////////////////////////////////////////////////////////////
Function EnumRegistryKeys(ByVal Hkey As HKeys, ByVal KeyName As String) As _
    Collection
    Dim handle As Long
    Dim length As Long
    Dim index As Long
    Dim subkeyName As String
    
    ' initialize the result collection
    Set EnumRegistryKeys = New Collection
    
    ' Open the key, exit if not found
    If Len(KeyName) Then
        If RegOpenKeyEx(Hkey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all case the subsequent functions use hKey
        Hkey = handle
    End If
    
    Do
        ' this is the max length for a key name
        length = 260
        subkeyName = Space$(length)
        ' get the N-th key, exit the loop if not found
        If RegEnumKey(Hkey, index, subkeyName, length) Then Exit Do
        
        ' add to the result collection
        subkeyName = Left$(subkeyName, InStr(subkeyName, vbNullChar) - 1)
        EnumRegistryKeys.Add subkeyName, subkeyName
        ' prepare to query for next key
        index = index + 1
    Loop
   
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
        
End Function

'///////////////////////////////////////////////////////////////////////////////////////////
'  FUNCTION:     EnumRegistryValues
'
'  AUTHOR:       Michael J. Kempf        7/20/2001 12:17:39 PM
'
'  PURPOSE:       Enumerate values under a given registry key
'
'  PARAMETERS:
'       [in]     Hkey      (Long) = HKEY_LOCAL_MACHINE,HKEY_CURRENT_USER...
'       [in]     KeyName (String) = Key name to enumerate
'
'  RETURN:       Collection -  each element of the collection is a 2-element array of Variants:
'                element(0) is the value name, element(1) is the value's value
'
'///////////////////////////////////////////////////////////////////////////////////////////
Function EnumRegistryValues(ByVal Hkey As HKeys, ByVal KeyName As String) As _
    Collection
    Dim handle As Long
    Dim index As Long
    Dim valueType As Long
    Dim name As String
    Dim nameLen As Long
    Dim resLong As Long
    Dim resString As String
    Dim dataLen As Long
    Dim valueInfo(0 To 1) As Variant
    Dim retVal As Long
    
    ' initialize the result
    Set EnumRegistryValues = New Collection
    
    ' Open the key, exit if not found.
    If Len(KeyName) Then
        If RegOpenKeyEx(Hkey, KeyName, 0, KEY_READ, handle) Then Exit Function
        ' in all cases, subsequent functions use hKey
        Hkey = handle
    End If
    
    Do
        ' this is the max length for a key name
        nameLen = 260
        name = Space$(nameLen)
        ' prepare the receiving buffer for the value
        dataLen = 4096
        ReDim resBinary(0 To dataLen - 1) As Byte
        
        ' read the value's name and data
        ' exit the loop if not found
        retVal = RegEnumValue(Hkey, index, name, nameLen, ByVal 0&, valueType, _
            resBinary(0), dataLen)
        
        ' enlarge the buffer if you need more space
        If retVal = ERROR_MORE_DATA Then
            ReDim resBinary(0 To dataLen - 1) As Byte
            retVal = RegEnumValue(Hkey, index, name, nameLen, ByVal 0&, _
                valueType, resBinary(0), dataLen)
        End If
        ' exit the loop if any other error (typically, no more values)
        If retVal Then Exit Do
        
        ' retrieve the value's name
        valueInfo(0) = Left$(name, nameLen)
        
        ' return a value corresponding to the value type
        Select Case valueType
            Case REG_DWORD
                CopyMemory resLong, resBinary(0), 4
                valueInfo(1) = resLong
            Case REG_SZ, REG_EXPAND_SZ
                ' copy everything but the trailing null char
                resString = Space$(dataLen - 1)
                CopyMemory ByVal resString, resBinary(0), dataLen - 1
                valueInfo(1) = resString
            Case REG_BINARY
                ' shrink the buffer if necessary
                If dataLen < UBound(resBinary) + 1 Then
                    ReDim Preserve resBinary(0 To dataLen - 1) As Byte
                End If
                valueInfo(1) = resBinary()
            Case REG_MULTI_SZ
                ' copy everything but the 2 trailing null chars
                resString = Space$(dataLen - 2)
                CopyMemory ByVal resString, resBinary(0), dataLen - 2
                valueInfo(1) = resString
            Case Else
                ' Unsupported value type - do nothing
        End Select
        
        ' add the array to the result collection
        ' the element's key is the value's name
        EnumRegistryValues.Add valueInfo, valueInfo(0)
        
        index = index + 1
    Loop
   
    ' Close the key, if it was actually opened
    If handle Then RegCloseKey handle
        
End Function
