Imports System.Runtime.InteropServices
Imports System.Text

Public Module IniFileHelper

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Function GetPrivateProfileString(
         lpAppName As String,
         lpKeyName As String,
         lpDefault As String,
         lpReturnedString As StringBuilder,
         nSize As Integer,
         lpFileName As String) As Integer
    End Function

    <DllImport("kernel32.dll", SetLastError:=True)>
    Private Function WritePrivateProfileString(
         lpAppName As String,
         lpKeyName As String,
         lpString As String,
         lpFileName As String) As Boolean
    End Function

    Public Function ReadIni(filePath As String, section As String, key As String, Optional defaultValue As String = "") As String
        Dim result As New StringBuilder(255)
        GetPrivateProfileString(section, key, defaultValue, result, result.Capacity, filePath)
        Return result.ToString()
    End Function

    Public Sub WriteIni(filePath As String, section As String, key As String, value As String)
        WritePrivateProfileString(section, key, value, filePath)
    End Sub

End Module
