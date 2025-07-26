Imports System.IO
Imports System.Management
Imports Microsoft.Win32

Public Class SoftwareInfoCollector
    Public Shared Sub Collect(sw As StreamWriter)
        WriteRowHeader(sw)
        WriteOSInfo(sw)
        WriteInstalledApplications(sw)
        WriteAntivirusStatus(sw)
    End Sub

    Private Shared Sub WriteOSInfo(sw As StreamWriter)
        Try
            Dim osName As String = "Unknown"
            Dim osBuild As String = "Unknown"
            Dim osDisplayVersion As String = "Unknown"
            Dim editionId As String = "Unknown"

            Using mos As New ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem")
                For Each mo As ManagementObject In mos.Get()
                    osName = TryCast(mo("Caption"), String)
                Next
            End Using

            Using regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")
                If regKey IsNot Nothing Then
                    Dim currentBuild = TryCast(regKey.GetValue("CurrentBuild"), String)
                    Dim ubrObj = regKey.GetValue("UBR")
                    Dim displayVersion = TryCast(regKey.GetValue("DisplayVersion"), String)
                    Dim edition = TryCast(regKey.GetValue("EditionID"), String)

                    If currentBuild IsNot Nothing Then
                        If ubrObj IsNot Nothing Then
                            osBuild = $"{currentBuild}.{Convert.ToInt32(ubrObj)}"
                        Else
                            osBuild = currentBuild
                        End If
                    End If

                    If Not String.IsNullOrWhiteSpace(displayVersion) Then
                        osDisplayVersion = displayVersion
                    End If

                    If Not String.IsNullOrWhiteSpace(edition) Then
                        editionId = edition
                    End If
                End If
            End Using

            WriteRow(sw, "OS", "Name", osName)
            WriteRow(sw, "OS", "Edition", editionId)
            WriteRow(sw, "OS", "Build", osBuild)
            WriteRow(sw, "OS", "Display Version", osDisplayVersion)
            WriteRow(sw, "OS", "Architecture", If(Environment.Is64BitOperatingSystem, "64-bit", "32-bit"))

        Catch ex As Exception
            WriteRow(sw, "OS", "Error", ex.Message)
        End Try
    End Sub

    Private Shared Sub WriteInstalledApplications(sw As StreamWriter)
        Try
            Dim regViews = {RegistryView.Registry64, RegistryView.Registry32}
            Dim hives = {
            New With {.Hive = RegistryHive.LocalMachine, .Source = "Registry (HKLM)"},
            New With {.Hive = RegistryHive.CurrentUser, .Source = "Registry (HKCU)"}
        }

            For Each hive In hives
                For Each view In regViews
                    Using baseKey = RegistryKey.OpenBaseKey(hive.Hive, view)
                        Using uninstallKey = baseKey.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
                            If uninstallKey IsNot Nothing Then
                                For Each subKeyName In uninstallKey.GetSubKeyNames()
                                    Using subKey = uninstallKey.OpenSubKey(subKeyName)
                                        If subKey IsNot Nothing Then
                                            Dim displayName = TryCast(subKey.GetValue("DisplayName"), String)
                                            Dim version = TryCast(subKey.GetValue("DisplayVersion"), String)
                                            Dim publisher = TryCast(subKey.GetValue("Publisher"), String)
                                            Dim installLocation = TryCast(subKey.GetValue("InstallLocation"), String)

                                            If Not String.IsNullOrWhiteSpace(displayName) Then
                                                WriteRow(sw, "Application", "Name", displayName, source:=hive.Source)
                                                WriteRow(sw, "Application", "Version", version, source:=hive.Source)
                                                If Not String.IsNullOrWhiteSpace(publisher) Then
                                                    WriteRow(sw, "Application", "Publisher", publisher, source:=hive.Source)
                                                End If
                                                If Not String.IsNullOrWhiteSpace(installLocation) Then
                                                    WriteRow(sw, "Application", "Install Location", installLocation, source:=hive.Source)
                                                End If
                                                WriteRow(sw, "Application", "----------------", "----------------", source:=hive.Source)
                                            End If
                                        End If
                                    End Using
                                Next
                            End If
                        End Using
                    End Using
                Next
            Next

        Catch ex As Exception
            WriteRow(sw, "Application", "Error", ex.Message)
        End Try
    End Sub

    Private Shared Sub WriteAntivirusStatus(sw As StreamWriter)
        Try
            Using mos As New ManagementObjectSearcher("root\SecurityCenter2", "SELECT displayName, productState FROM AntiVirusProduct")
                Dim idx As Integer = 1
                For Each mo As ManagementObject In mos.Get()
                    Dim name As String = If(TryCast(mo("displayName"), String), "Unknown")
                    Dim productState As Integer = Convert.ToInt32(mo("productState"))
                    Dim hexState As String = $"0x{productState:X6}"

                    ' Breakdown bits
                    Dim statusByte As Integer = productState And &HFF
                    Dim reportingByte As Integer = (productState >> 8) And &HFF
                    Dim definitionByte As Integer = (productState >> 16) And &HFF

                    ' Interpret state
                    Dim statusStr As String = ""
                    Select Case statusByte
                        Case &H10
                            statusStr = If((definitionByte And &H10) = &H10, "Enabled and Up-to-date", "Enabled but Outdated")
                        Case &H0
                            statusStr = "Not Reporting or Passive Mode"
                        Case Else
                            statusStr = "Unknown or Inactive"
                    End Select

                    ' Write rows using structured format
                    WriteRow(sw, $"Antivirus {idx}", "Name", name, source:="WMI:AntiVirusProduct")
                    WriteRow(sw, $"Antivirus {idx}", "Status", statusStr, source:="WMI:AntiVirusProduct")
                    WriteRow(sw, $"Antivirus {idx}", "ProductState (Hex)", hexState, source:="WMI:AntiVirusProduct")
                    WriteRow(sw, $"Antivirus {idx}", "Status Byte", statusByte.ToString(), source:="WMI:AntiVirusProduct")
                    WriteRow(sw, $"Antivirus {idx}", "Reporting Byte", reportingByte.ToString(), source:="WMI:AntiVirusProduct")
                    WriteRow(sw, $"Antivirus {idx}", "Definition Byte", definitionByte.ToString(), source:="WMI:AntiVirusProduct")

                    idx += 1
                Next
            End Using
        Catch ex As Exception
            WriteRow(sw, "Antivirus", "Error", ex.Message, source:="WMI:AntiVirusProduct")
        End Try
    End Sub
End Class
