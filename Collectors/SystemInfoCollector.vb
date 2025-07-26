Imports System.IO
Imports System.Management

Public Class SystemInfoCollector
    Public Shared Sub Collect(sw As StreamWriter)
        WriteRowHeader(sw)
        WriteHostname(sw)
        WriteLoggedInUser(sw)
        WriteDomainInfo(sw)
        WriteSystemIdentifiersInfo(sw)
    End Sub

    Private Shared Sub WriteHostname(sw As StreamWriter)
        Try
            Dim hostname As String = Environment.MachineName
            WriteRow(sw, "System Identity", "Hostname", hostname)
        Catch ex As Exception
            WriteRow(sw, "System Identity", "Hostname", $"Error: {ex.Message}", "", "System.Environment")
        End Try
    End Sub

    Private Shared Sub WriteLoggedInUser(sw As StreamWriter)
        Try
            Dim loggedUser As String = $"{Environment.UserDomainName}\{Environment.UserName}"
            WriteRow(sw, "System Identity", "Logged-In User", loggedUser, "", "System.Environment")
        Catch ex As Exception
            WriteRow(sw, "System Identity", "Logged-In User", $"Error: {ex.Message}", "", "System.Environment")
        End Try
    End Sub

    Private Shared Sub WriteDomainInfo(sw As StreamWriter)
        Try
            Using mos As New ManagementObjectSearcher("SELECT PartOfDomain, Domain FROM Win32_ComputerSystem")
                For Each mo In mos.Get()
                    Dim partOfDomain = If(mo("PartOfDomain"), False)
                    Dim domainName = If(mo("Domain"), "Unknown")
                    Dim domainStatus = If(partOfDomain, "Domain Joined", "Not Domain Joined")

                    WriteRow(sw, "System Identity", "Domain Status", domainStatus, "", "WMI:Win32_ComputerSystem")
                    WriteRow(sw, "System Identity", "Domain Name", domainName, "", "WMI:Win32_ComputerSystem")
                Next
            End Using
        Catch ex As Exception
            WriteRow(sw, "System Identity", "Domain", $"Error: {ex.Message}", "", "WMI:Win32_ComputerSystem")
        End Try
    End Sub

    Private Shared Sub WriteSystemIdentifiersInfo(sw As StreamWriter)
        Try
            ' --- Machine UUID ---
            Dim uuid As String = ""
            Using uuidSearcher As New ManagementObjectSearcher("SELECT UUID FROM Win32_ComputerSystemProduct")
                For Each obj As ManagementObject In uuidSearcher.Get()
                    uuid = obj("UUID")?.ToString()
                Next
            End Using
            WriteRow(sw, "System Identity", "Machine UUID", uuid, "", "WMI:Win32_ComputerSystemProduct")

            ' --- System Serial Number ---
            Dim serialNumber As String = ""
            Using biosSearcher As New ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BIOS")
                For Each obj As ManagementObject In biosSearcher.Get()
                    serialNumber = obj("SerialNumber")?.ToString()
                Next
            End Using
            WriteRow(sw, "System Identity", "System Serial Number", serialNumber, "", "WMI:Win32_BIOS")

            ' --- BIOS Mode (UEFI / Legacy) ---
            Dim biosMode As String = "Unknown"
            Dim firmwareType = GetFirmwareType()
            If firmwareType.HasValue Then
                biosMode = If(firmwareType.Value = 2, "UEFI", "Legacy")
            End If
            WriteRow(sw, "System Identity", "BIOS Mode", biosMode, "", "Firmware")

            ' --- Chassis Type ---
            Dim chassisType As String = ""
            Using chassisSearcher As New ManagementObjectSearcher("SELECT ChassisTypes FROM Win32_SystemEnclosure")
                For Each obj As ManagementObject In chassisSearcher.Get()
                    Dim types = TryCast(obj("ChassisTypes"), UShort())
                    If types IsNot Nothing AndAlso types.Length > 0 Then
                        chassisType = GetChassisTypeName(types(0))
                        Exit For
                    End If
                Next
            End Using
            WriteRow(sw, "System Identity", "Chassis Type", chassisType, "", "WMI:Win32_SystemEnclosure")

        Catch ex As Exception
            WriteRow(sw, "System Identity", "System Identifier Info", $"Error: {ex.Message}", "", "System")
        End Try
    End Sub

    Private Declare Function GetFirmwareType Lib "kernel32.dll" (ByRef firmwareType As UInteger) As Boolean
    Private Shared Function GetFirmwareType() As UInteger?
        Try
            Dim fwType As UInteger = 0
            If GetFirmwareType(fwType) Then
                Return fwType ' 1 = BIOS (Legacy), 2 = UEFI
            End If
        Catch
        End Try
        Return Nothing
    End Function

    Private Shared Function GetChassisTypeName(code As UShort) As String
        Select Case code
            Case 3 : Return "Desktop"
            Case 4 : Return "Low Profile Desktop"
            Case 5 : Return "Pizza Box"
            Case 6 : Return "Mini Tower"
            Case 7 : Return "Tower"
            Case 8 : Return "Portable"
            Case 9 : Return "Laptop"
            Case 10 : Return "Notebook"
            Case 11 : Return "Hand Held"
            Case 12 : Return "Docking Station"
            Case 13 : Return "All in One"
            Case 14 : Return "Sub Notebook"
            Case 15 : Return "Space-saving"
            Case 16 : Return "Lunch Box"
            Case 17 : Return "Main System Chassis"
            Case 18 : Return "Expansion Chassis"
            Case 19 : Return "SubChassis"
            Case 20 : Return "Bus Expansion Chassis"
            Case 21 : Return "Peripheral Chassis"
            Case 22 : Return "RAID Chassis"
            Case 23 : Return "Rack Mount Chassis"
            Case 24 : Return "Sealed-case PC"
            Case 30 : Return "Tablet"
            Case 31 : Return "Convertible"
            Case 32 : Return "Detachable"
            Case Else : Return "Other (" & code & ")"
        End Select
    End Function
End Class
