Imports System.IO
Imports System.Management
Imports Microsoft.Win32

Public Class LicenseInfoCollector
    Public Shared Sub Collect(sw As StreamWriter)
        WriteRowHeader(sw)
        WriteWindowsLicenseInfo(sw)
        WriteOfficeLicenseInfo(sw)
        WriteOfficeLicenseViaOSPP(sw)
    End Sub
    Private Shared Sub WriteWindowsLicenseInfo(sw As StreamWriter)
        Try
            Dim searcher As New ManagementObjectSearcher(
            "SELECT Name, PartialProductKey, LicenseStatus FROM SoftwareLicensingProduct " &
            "WHERE PartialProductKey IS NOT NULL AND LicenseStatus IS NOT NULL")

            Dim found As Boolean = False
            For Each obj As ManagementObject In searcher.Get()
                Dim name As String = TryCast(obj("Name"), String)
                Dim key As String = TryCast(obj("PartialProductKey"), String)
                Dim statusCode As Integer = Convert.ToInt32(obj("LicenseStatus"))
                Dim statusDesc As String = GetLicenseStatusString(statusCode)

                If Not String.IsNullOrEmpty(name) Then WriteRow(sw, "Windows License", "Product Name", name, source:="WMI:SoftwareLicensingProduct")
                If Not String.IsNullOrEmpty(key) Then WriteRow(sw, "Windows License", "Partial Product Key", key, source:="WMI:SoftwareLicensingProduct")
                WriteRow(sw, "Windows License", "License Status", statusDesc, source:="WMI:SoftwareLicensingProduct")

                found = True
                Exit For ' Only the first relevant record is usually sufficient
            Next

            If Not found Then
                WriteRow(sw, "Windows License", "Status", "No license information found", source:="WMI:SoftwareLicensingProduct")
            End If
        Catch ex As Exception
            WriteRow(sw, "Windows License", "Error", ex.Message, source:="WMI:SoftwareLicensingProduct")
        End Try
    End Sub

    Private Shared Function GetLicenseStatusString(statusCode As Integer) As String
        Select Case statusCode
            Case 0 : Return "Unlicensed"
            Case 1 : Return "Licensed"
            Case 2 : Return "OOB Grace"
            Case 3 : Return "OOT Grace"
            Case 4 : Return "Non-Genuine Grace"
            Case 5 : Return "Notification"
            Case 6 : Return "Extended Grace"
            Case Else : Return "Unknown (" & statusCode & ")"
        End Select
    End Function

    Private Shared Sub WriteOfficeLicenseInfo(sw As StreamWriter)
        Try
            Dim officeVersions As New Dictionary(Of String, String()) From {
            {"Office 2016 / 2019 / 365", {"SOFTWARE\Microsoft\Office\16.0\Registration"}},
            {"Office 2013", {"SOFTWARE\Microsoft\Office\15.0\Registration"}},
            {"Office 2010", {"SOFTWARE\Microsoft\Office\14.0\Registration"}}
        }

            Dim foundAny As Boolean = False

            For Each kvp In officeVersions
                Dim officeVersionName = kvp.Key
                For Each regPath In kvp.Value
                    Try
                        Using regKey As RegistryKey = Registry.LocalMachine.OpenSubKey(regPath)
                            If regKey IsNot Nothing Then
                                For Each subKeyName In regKey.GetSubKeyNames()
                                    Using subKey As RegistryKey = regKey.OpenSubKey(subKeyName)
                                        If subKey IsNot Nothing Then
                                            Dim productName = TryCast(subKey.GetValue("ProductName"), String)
                                            Dim productId = TryCast(subKey.GetValue("ProductId"), String)
                                            Dim digitalProductId = TryCast(subKey.GetValue("DigitalProductId"), Byte())

                                            If Not String.IsNullOrEmpty(productName) Then
                                                WriteRow(sw, "Office License", "Office Version", officeVersionName, source:=$"Registry:{regPath}")
                                                WriteRow(sw, "Office License", "Product Name", productName, source:=$"Registry:{regPath}")
                                                WriteRow(sw, "Office License", "Product ID", If(String.IsNullOrEmpty(productId), "N/A", productId), source:=$"Registry:{regPath}")
                                                WriteRow(sw, "Office License", "Partial Key", "<Hidden or Not Readable>", source:=$"Registry:{regPath}")
                                                foundAny = True
                                            End If
                                        End If
                                    End Using
                                Next
                            End If
                        End Using
                    Catch ex As Exception
                        WriteRow(sw, "Office License", $"Error Reading {officeVersionName}", ex.Message, source:=$"Registry:{regPath}")
                    End Try
                Next
            Next

            If Not foundAny Then
                WriteRow(sw, "Office License", "Status", "No Office license information found", source:="Registry")
            End If
        Catch ex As Exception
            WriteRow(sw, "Office License", "Fatal Error", ex.Message, source:="System")
        End Try
    End Sub

    Private Shared Sub WriteOfficeLicenseViaOSPP(sw As StreamWriter)
        Try
            Dim officePaths As String() = {
            "C:\Program Files\Microsoft Office",
            "C:\Program Files (x86)\Microsoft Office"
        }

            Dim foundScript As Boolean = False

            For Each basePath In officePaths
                If Directory.Exists(basePath) Then
                    Dim officeDirs = Directory.GetDirectories(basePath, "Office*")

                    For Each dir As String In officeDirs
                        Dim scriptPath = Path.Combine(dir, "ospp.vbs")
                        If File.Exists(scriptPath) Then
                            foundScript = True

                            Dim psi As New ProcessStartInfo("cscript.exe") With {
                            .Arguments = $"//Nologo ""{scriptPath}"" /dstatus",
                            .CreateNoWindow = True,
                            .UseShellExecute = False,
                            .RedirectStandardOutput = True
                        }

                            Using proc As Process = Process.Start(psi)
                                Dim output As String = proc.StandardOutput.ReadToEnd()
                                proc.WaitForExit()

                                ' Parse each line
                                Dim lines = output.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                                Dim idx = 1
                                Dim category = $"Office License via OSPP ({Path.GetFileName(dir)})"

                                For Each line In lines
                                    If line.Contains(":") Then
                                        Dim parts = line.Split(New Char() {":"c}, 2)
                                        Dim key = parts(0).Trim()
                                        Dim value = parts(1).Trim()
                                        If Not String.IsNullOrWhiteSpace(key) Then
                                            WriteRow(sw, category, key, value, source:="ospp.vbs")
                                        End If
                                    Else
                                        ' For lines like "---Processing---" or status blocks
                                        WriteRow(sw, category, $"Info {idx}", line.Trim(), source:="ospp.vbs")
                                        idx += 1
                                    End If
                                Next

                                WriteRow(sw, category, "Execution Status", "Completed", source:="ospp.vbs")
                            End Using

                            Exit For ' Stop after first valid OSPP.vbs found
                        End If
                    Next
                End If

                If foundScript Then Exit For
            Next

            If Not foundScript Then
                WriteRow(sw, "Office License via OSPP", "Status", "ospp.vbs not found (Office may be Click-to-Run or not installed)", source:="ospp.vbs")
            End If
        Catch ex As Exception
            WriteRow(sw, "Office License via OSPP", "Error", ex.Message, source:="ospp.vbs")
        End Try
    End Sub
End Class
