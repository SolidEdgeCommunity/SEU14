Option Strict Off
Option Explicit On
Module modRegistry
    Public Structure FILETIME
        Dim dwLowDateTime As Integer
        Dim dwHighDateTime As Integer
    End Structure
    Public Const DELETE As Integer = &H10000
    Public Const READ_CONTROL As Integer = &H20000
    Public Const WRITE_DAC As Integer = &H40000
    Public Const WRITE_OWNER As Integer = &H80000
    Public Const SYNCHRONIZE As Integer = &H100000

    Public Const STANDARD_RIGHTS_REQUIRED As Integer = &HF0000
    Public Const STANDARD_RIGHTS_READ As Integer = READ_CONTROL
    Public Const STANDARD_RIGHTS_WRITE As Integer = READ_CONTROL
    Public Const STANDARD_RIGHTS_EXECUTE As Integer = READ_CONTROL
    Public Const STANDARD_RIGHTS_ALL As Integer = &H1F0000
    Public Const SPECIFIC_RIGHTS_ALL As Integer = &HFFFFS

    Public Const KEY_QUERY_VALUE As Integer = &H1S
    Public Const KEY_SET_VALUE As Integer = &H2S
    Public Const KEY_CREATE_SUB_KEY As Integer = &H4S
    Public Const KEY_ENUMERATE_SUB_KEYS As Integer = &H8S
    Public Const KEY_NOTIFY As Integer = &H10S
    Public Const KEY_CREATE_LINK As Integer = &H20S

    Public Const KEY_READ As Integer = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
    Public Const HKEY_CURRENT_USER As Integer = &H80000001
    Public Const HKEY_LOCAL_MACHINE As Integer = &H80000002
    Public Const KEY_ALL_ACCESS As Integer = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

    Public Const ERROR_SUCCESS As Short = 0
    Public Const ERROR_NO_MORE_ITEMS As Short = 259
    Public Const ERROR_FILE_NOT_FOUND As Short = 2
    Public Const ERROR_MORE_DATA As Short = 234

    Public Const REG_DWORD As Integer = 4 ' 32-bit number
    Public Const REG_SZ As Short = 1 ' Unicode nul terminated string
    Public Const ERROR_NONE As Short = 0

    Declare Function RegCloseKey Lib "advapi32" (ByVal hKey As Integer) As Integer
    Declare Function RegOpenKeyEx Lib "advapi32" Alias "RegOpenKeyExA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByVal ulOptions As Integer, ByVal samDesired As Integer, ByRef phkResult As Integer) As Integer
    Declare Function RegQueryValueEx Lib "advapi32" Alias "RegQueryValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Integer, ByRef lpcbData As Integer) As Integer
    Declare Function RegCreateKey Lib "advapi32" Alias "RegCreateKeyA" (ByVal hKey As Integer, ByVal lpSubKey As String, ByRef phkResult As Integer) As Integer
    Declare Function RegSetValueEx Lib "advapi32" Alias "RegSetValueExA" (ByVal hKey As Integer, ByVal lpszValueName As String, ByVal dwReserved As Integer, ByVal fdwType As Integer, ByRef lpbData As Integer, ByVal cbData As Integer) As Integer
    Declare Function RegQueryValueExNULL Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As Integer, ByRef lpcbData As Integer) As Integer
    Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByVal lpData As String, ByRef lpcbData As Integer) As Integer
    Declare Function RegQueryValueExLong Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal lpReserved As Integer, ByRef lpType As Integer, ByRef lpData As Integer, ByRef lpcbData As Integer) As Integer
    Declare Function RegSetValueExString Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal Reserved As Integer, ByVal dwType As Integer, ByVal lpValue As String, ByVal cbData As Integer) As Integer
    Declare Function RegSetValueExLong Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Integer, ByVal lpValueName As String, ByVal Reserved As Integer, ByVal dwType As Integer, ByRef lpValue As Integer, ByVal cbData As Integer) As Integer
    Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Integer, ByVal lpValueName As String) As Integer

    Public Function GetAssemblyMode(ByRef Value As Short) As Boolean
        Dim lResult As Integer
        Dim hkGlobal As Integer
        Dim dwType As Integer
        Dim lpReserved As Integer
        Dim lngDataBuffer As Integer
        Dim strSEsubkey As String

        GetAssemblyMode = True
        strSEsubkey = "Software\Unigraphics Solutions\Solid Edge\Version " & SEVersion

        'Open Registry
        lResult = RegOpenKeyEx(HKEY_CURRENT_USER, strSEsubkey & "\FileOpen", 0, KEY_READ, hkGlobal)
        If lResult <> ERROR_SUCCESS Then
            GetAssemblyMode = False
            Exit Function
        End If
        'Read the mode value.
        lResult = RegQueryValueEx(hkGlobal, "Mode", lpReserved, dwType, lngDataBuffer, 4)
        If lResult <> ERROR_SUCCESS Then
            GetAssemblyMode = False
            Exit Function
        Else
            Value = lngDataBuffer
        End If

        ' Close the key.
        Call RegCloseKey(hkGlobal)
    End Function

    Public Function SetAssemblyMode(ByRef Value As Short) As Boolean
        Dim lResult As Integer
        Dim hkGlobal As Integer
        Dim cbData As Integer
        Dim lngDataBuffer As Integer
        Dim lpReserved As Integer
        Dim strSEsubkey As String

        SetAssemblyMode = True
        strSEsubkey = "Software\Unigraphics Solutions\Solid Edge\Version " & SEVersion

        ' Open Registry
        lResult = RegOpenKeyEx(HKEY_CURRENT_USER, strSEsubkey & "\FileOpen", 0, KEY_SET_VALUE, hkGlobal)
        If lResult <> ERROR_SUCCESS Then
            lResult = RegCreateKey(HKEY_CURRENT_USER, strSEsubkey & "\FileOpen", hkGlobal)

            If lResult <> ERROR_SUCCESS Then
                SetAssemblyMode = False
                Exit Function
            End If
        End If

        ' Edit or write the mode.
        lngDataBuffer = Value
        cbData = 4
        lResult = RegSetValueEx(hkGlobal, "Mode", lpReserved, REG_DWORD, lngDataBuffer, cbData)
        If lResult <> ERROR_SUCCESS Then
            SetAssemblyMode = False
            Exit Function
        End If

        ' Close the key.
        Call RegCloseKey(hkGlobal)
    End Function

    Public Sub GetSolidEdgePath()
        ' This function used to get the Solid Edge path


        Try
            Dim oSEID As SEInstallDataLib.SEInstallData
            oSEID = CreateObject("SolidEdge.InstallData")
            strSEVersion = CStr(oSEID.GetMajorVersion)
            strSEInstalledPath = oSEID.GetInstalledPath
            strSEInstalledPath = Mid(strSEInstalledPath, 1, strSEInstalledPath.LastIndexOf("\"))
            Module1.Garbage_Collect(oSEID)
            Exit Sub
        Catch ex As Exception
            MessageBox.Show("error getting the Solid Edge version" + ex.Message)
            End
        End Try
        




    End Sub

    Private Function SEVersion() As String
        Dim oSEID As SEInstallDataLib.SEInstallData

        On Error GoTo ErrorHandler
        oSEID = CreateObject("SolidEdge.InstallData")
        SEVersion = CStr(oSEID.GetMajorVersion)
        Exit Function

ErrorHandler:
        Err.Clear()

    End Function

    Public Function SetOpenSaveMacroFlag(ByRef Value As Short) As Boolean
        Dim lResult As Integer
        Dim hkGlobal As Integer
        Dim cbData As Integer
        Dim lpReserved As Integer
        Dim lngDataBuffer As Integer
        Dim strSEsubkey As String

        SetOpenSaveMacroFlag = True
        strSEsubkey = "Software\Unigraphics Solutions\Solid Edge\Version " & SEVersion

        'Open Registry
        lResult = RegOpenKeyEx(HKEY_CURRENT_USER, strSEsubkey & "\FileOpen", 0, KEY_SET_VALUE, hkGlobal)
        If lResult <> ERROR_SUCCESS Then
            lResult = RegCreateKey(HKEY_CURRENT_USER, strSEsubkey & "\FileOpen", hkGlobal)
            If lResult <> ERROR_SUCCESS Then
                SetOpenSaveMacroFlag = False
                Exit Function
            End If
        End If

        ' Edit or write the mode.
        lngDataBuffer = Value
        cbData = 4
        lResult = RegSetValueEx(hkGlobal, "OpenSaveMacro", lpReserved, REG_DWORD, lngDataBuffer, cbData)
        If lResult <> ERROR_SUCCESS Then
            SetOpenSaveMacroFlag = False
            Exit Function
        End If

        ' Close the key.
        Call RegCloseKey(hkGlobal)
    End Function

    Public Function QueryValue(ByRef lPredefinedKey As Integer, ByRef sKeyName As String, ByRef sValueName As String) As Object
        ' Description:
        '   This Function will return the data field of a value
        '
        ' Syntax:
        '   Variable = QueryValue(Location, KeyName, ValueName)
        '
        '   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
        '   , HKEY_USERS
        '
        '   KeyName is the key that the value is under (example: "Software\Microsoft\Windows\CurrentVersion\Explorer")
        '
        '   ValueName is the name of the value you want to access (example: "link")

        Dim lRetVal As Integer 'result of the API functions
        Dim hKey As Integer 'handle of opened key
        Dim vValue As Object = Nothing 'setting of queried value


        lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
        lRetVal = QueryValueEx(hKey, sValueName, vValue)
        'MsgBox vValue
        QueryValue = vValue
        RegCloseKey(hKey)
    End Function

    Function QueryValueEx(ByVal lhKey As Integer, ByVal szValueName As String, ByRef vValue As Object) As Integer
        Dim cch As Integer
        Dim lrc As Integer
        Dim lType As Integer
        Dim lValue As Integer
        Dim sValue As String

        On Error GoTo QueryValueExError

        ' Determine the size and type of data to be read

        lrc = RegQueryValueExNULL(lhKey, szValueName, 0, lType, 0, cch)
        If lrc <> ERROR_NONE Then Error (5)

        Select Case lType
            ' For strings
            Case REG_SZ
                sValue = New String(Chr(0), cch)
                lrc = RegQueryValueExString(lhKey, szValueName, 0, lType, sValue, cch)
                If lrc = ERROR_NONE Then
                    vValue = Left(sValue, cch)
                Else
                    vValue = Nothing
                End If

                ' For DWORDS
            Case REG_DWORD
                lrc = RegQueryValueExLong(lhKey, szValueName, 0, lType, lValue, cch)
                If lrc = ERROR_NONE Then vValue = lValue
            Case Else
                'all other data types not supported
                lrc = -1
        End Select

QueryValueExExit:

        QueryValueEx = lrc
        Exit Function

QueryValueExError:

        Resume QueryValueExExit

    End Function


    Public Function SetKeyValue(ByRef lPredefinedKey As Integer, ByRef sKeyName As String, ByRef sValueName As String, ByRef vValueSetting As Object, ByRef lValueType As Integer) As Object
        ' Description:
        '   This Function will set the data field of a value
        '
        ' Syntax:
        '   QueryValue Location, KeyName, ValueName, ValueSetting, ValueType
        '
        '   Location must equal HKEY_CLASSES_ROOT, HKEY_CURRENT_USER, HKEY_lOCAL_MACHINE
        '   , HKEY_USERS
        '
        '   KeyName is the key that the value is under (example: "Key1\SubKey1")
        '
        '   ValueName is the name of the value you want create, or set the value of (example: "ValueTest")
        '
        '   ValueSetting is what you want the value to equal
        '
        '   ValueType must equal either REG_SZ (a string) Or REG_DWORD (an integer)

        Dim lRetVal As Integer 'result of the SetValueEx function
        Dim hKey As Integer 'handle of open key

        'open the specified key

        lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
        lRetVal = SetValueEx(hKey, sValueName, lValueType, vValueSetting)
        RegCloseKey(hKey)

    End Function



    Public Function SetValueEx(ByVal hKey As Integer, ByRef sValueName As String, ByRef lType As Integer, ByRef vValue As Object) As Integer
        Dim lValue As Integer
        Dim sValue As String

        Select Case lType
            Case REG_SZ
                sValue = vValue
                SetValueEx = RegSetValueExString(hKey, sValueName, 0, lType, sValue, Len(sValue))
            Case REG_DWORD
                lValue = vValue
                SetValueEx = RegSetValueExLong(hKey, sValueName, 0, lType, lValue, 4)
        End Select

    End Function

    Public Function DeleteValue(ByRef lPredefinedKey As Integer, ByRef sKeyName As String, ByRef Value As String) As Integer
        Dim lRetVal As Integer 'result of the SetValueEx function
        Dim hKey As Integer 'handle of open key

        lRetVal = RegOpenKeyEx(lPredefinedKey, sKeyName, 0, KEY_ALL_ACCESS, hKey)
        lRetVal = RegDeleteValue(hKey, Value)
        RegCloseKey(hKey)

    End Function

End Module