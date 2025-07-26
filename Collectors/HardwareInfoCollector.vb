Imports System.IO
Imports System.Management
Imports System.Text.RegularExpressions

Public Class HardwareInfoCollector
    Public Shared Sub Collect(sw As StreamWriter)
        WriteRowHeader(sw)
        WriteBIOSInfo(sw)
        WriteMotherboardInfo(sw)
        WriteCPUInfo(sw)
        WriteRAMInfo(sw)
        WriteDiskInfo(sw)
        WriteGPUInfo(sw)
        WriteBatteryInfo(sw)
    End Sub

    Private Shared Sub WriteCPUInfo(sw As StreamWriter)
        Try
            Using mos As New ManagementObjectSearcher("SELECT Name, MaxClockSpeed FROM Win32_Processor")
                Dim idx As Integer = 1
                For Each mo As ManagementObject In mos.Get()
                    Dim name = mo("Name")?.ToString()
                    Dim maxClock = mo("MaxClockSpeed")?.ToString()

                    WriteRow(sw, "CPU", $"Processor {idx} Name", name, "", "WMI:Win32_Processor")
                    WriteRow(sw, "CPU", $"Processor {idx} Max Clock Speed", maxClock, "MHz", "WMI:Win32_Processor")

                    idx += 1
                Next
            End Using
        Catch ex As Exception
            WriteRow(sw, "CPU", "Error", ex.Message, "", "System")
        End Try
    End Sub

    Private Shared Sub WriteRAMInfo(sw As StreamWriter)
        Try
            Using mos As New ManagementObjectSearcher("SELECT BankLabel, Capacity, Speed, SMBIOSMemoryType FROM Win32_PhysicalMemory")
                Dim idx = 1
                For Each mo As ManagementObject In mos.Get()
                    Dim bankLabel As String = If(mo("BankLabel"), "Unknown").ToString()
                    Dim capacityBytes As ULong = Convert.ToUInt64(If(mo("Capacity"), 0))
                    Dim capacityGB As Double = Math.Round(capacityBytes / 1024.0 / 1024 / 1024, 2)
                    Dim speedMHz As String = If(mo("Speed"), "Unknown").ToString()
                    Dim typeCode As Integer = Convert.ToInt32(If(mo("SMBIOSMemoryType"), 0))
                    Dim typeStr As String = GetMemoryTypeString(typeCode)

                    WriteRow(sw, "RAM", $"Slot {idx} Bank Label", bankLabel, "", "WMI:Win32_PhysicalMemory")
                    WriteRow(sw, "RAM", $"Slot {idx} Capacity", capacityGB.ToString("F2"), "GB", "WMI:Win32_PhysicalMemory")
                    WriteRow(sw, "RAM", $"Slot {idx} Speed", speedMHz, "MHz", "WMI:Win32_PhysicalMemory")
                    WriteRow(sw, "RAM", $"Slot {idx} Type", typeStr, "", "WMI:Win32_PhysicalMemory")

                    idx += 1
                Next
            End Using
        Catch ex As Exception
            WriteRow(sw, "RAM", "Error", ex.Message, "", "System")
        End Try
    End Sub

    Private Shared Function GetMemoryTypeString(typeCode As Integer) As String
        Select Case typeCode
            Case 20 : Return "DDR"
            Case 21 : Return "DDR2"
            Case 22 : Return "DDR2 FB-DIMM"
            Case 24 : Return "DDR3"
            Case 26 : Return "DDR4"
            Case 27 : Return "LPDDR"
            Case 28 : Return "LPDDR2"
            Case 29 : Return "LPDDR3"
            Case 30 : Return "LPDDR4"
            Case 34 : Return "DDR5"
            Case 35 : Return "LPDDR5"
            Case Else : Return "Unknown"
        End Select
    End Function

    Private Shared Sub WriteDiskInfo(sw As StreamWriter)
        Try
            Dim driveInfoList As New List(Of Tuple(Of String, String)) ' FriendlyName, MediaType

            ' Step 1: Query physical disk media types from MSFT_PhysicalDisk
            Dim scope As New ManagementScope("\\.\ROOT\Microsoft\Windows\Storage")
            scope.Connect()
            Using mos As New ManagementObjectSearcher(scope, New ObjectQuery("SELECT MediaType, FriendlyName FROM MSFT_PhysicalDisk"))
                For Each mo As ManagementObject In mos.Get()
                    Dim friendlyName As String = If(mo("FriendlyName"), "Unknown").ToString()
                    Dim mediaTypeCode As Integer = Convert.ToInt32(If(mo("MediaType"), 0))
                    Dim mediaTypeStr As String = If(mediaTypeCode = 3, "HDD", If(mediaTypeCode = 4, "SSD", If(mediaTypeCode = 5, "SCM", "Unknown")))
                    driveInfoList.Add(Tuple.Create(friendlyName, mediaTypeStr))
                Next
            End Using

            ' Step 2: Collect physical disk details from Win32_DiskDrive
            Using mosDisk As New ManagementObjectSearcher("SELECT Model, Size, SerialNumber FROM Win32_DiskDrive")
                Dim idx As Integer = 1
                For Each mo As ManagementObject In mosDisk.Get()
                    Dim model As String = If(mo("Model"), "Unknown").ToString()
                    Dim serial As String = If(mo("SerialNumber"), "Unknown").ToString()
                    Dim sizeBytes As ULong = Convert.ToUInt64(If(mo("Size"), 0))
                    Dim sizeGB As Double = sizeBytes / 1024 / 1024 / 1024
                    Dim typeStr As String = "Unknown"

                    ' Step 3: Match model with friendly name to determine disk type
                    For Each info In driveInfoList
                        If model.Contains(info.Item1) OrElse info.Item1.Contains(model) Then
                            typeStr = info.Item2
                            Exit For
                        End If
                    Next

                    ' Step 4: Write formatted rows
                    WriteRow(sw, "Disk", $"Drive {idx} Model", model, "", "WMI:Win32_DiskDrive")
                    WriteRow(sw, "Disk", $"Drive {idx} Serial Number", serial, "", "WMI:Win32_DiskDrive")
                    WriteRow(sw, "Disk", $"Drive {idx} Size", sizeGB.ToString("F2"), "GB", "WMI:Win32_DiskDrive")
                    WriteRow(sw, "Disk", $"Drive {idx} Type", typeStr, "", "WMI:MSFT_PhysicalDisk")
                    idx += 1
                Next
            End Using

            ' Additional disk information
            WriteDiskPartitionStyleInfo(sw)
            WriteEFIPartitionStatus(sw)
            WriteBootDriveInfo(sw)

        Catch ex As Exception
            WriteRow(sw, "Disk", "Error", ex.Message, "", "System")
        End Try
    End Sub

    Private Shared Sub WriteDiskPartitionStyleInfo(sw As StreamWriter)
        Try
            Using searcher As New ManagementObjectSearcher("SELECT DeviceID FROM Win32_DiskDrive")
                For Each drive As ManagementObject In searcher.Get()
                    Dim deviceId As String = drive("DeviceID").ToString()
                    Dim diskNumber As Integer = GetDiskNumberFromDeviceID(deviceId)
                    If diskNumber >= 0 Then
                        Dim scope As New ManagementScope("\\.\ROOT\Microsoft\Windows\Storage")
                        scope.Connect()
                        Dim query As New ObjectQuery($"SELECT PartitionStyle FROM MSFT_Disk WHERE Number={diskNumber}")
                        Using gptSearcher As New ManagementObjectSearcher(scope, query)
                            For Each disk As ManagementObject In gptSearcher.Get()
                                Dim style As Integer = Convert.ToInt32(disk("PartitionStyle"))
                                Dim styleName As String = If(style = 1, "MBR", If(style = 2, "GPT", "RAW/Unknown"))
                                WriteRow(sw, "Disk", $"Disk {diskNumber} Partition Style", styleName, "", "WMI:MSFT_Disk")
                            Next
                        End Using
                    End If
                Next
            End Using
        Catch ex As Exception
            WriteRow(sw, "Disk", "Partition Style Error", ex.Message, "", "System")
        End Try
    End Sub

    Private Shared Function GetDiskNumberFromDeviceID(deviceId As String) As Integer
        Try
            ' Normalize device ID (e.g., \\.\PHYSICALDRIVE0)
            Using partSearcher As New ManagementObjectSearcher("SELECT * FROM Win32_DiskDriveToDiskPartition")
                For Each obj As ManagementObject In partSearcher.Get()
                    Dim antecedent = obj("Antecedent").ToString()
                    Dim match = Regex.Match(antecedent, "PHYSICALDRIVE(\d+)")
                    If match.Success Then
                        Dim num = Integer.Parse(match.Groups(1).Value)
                        If deviceId.Contains("PHYSICALDRIVE" & num.ToString()) Then
                            Return num
                        End If
                    End If
                Next
            End Using
        Catch
        End Try
        Return -1
    End Function

    Private Shared Sub WriteEFIPartitionStatus(sw As StreamWriter)
        Try
            Dim found As Boolean = False
            Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_DiskPartition WHERE Type='EFI System Partition'")
                For Each partition As ManagementObject In searcher.Get()
                    WriteRow(sw, "Disk", "EFI System Partition", "Present", "", "WMI:Win32_DiskPartition")
                    found = True
                    Exit For
                Next
            End Using
            If Not found Then
                WriteRow(sw, "Disk", "EFI System Partition", "Not Found", "", "WMI:Win32_DiskPartition")
            End If
        Catch ex As Exception
            WriteRow(sw, "Disk", "EFI System Partition Error", ex.Message, "", "System")
        End Try
    End Sub

    Private Shared Sub WriteBootDriveInfo(sw As StreamWriter)
        Try
            Dim systemDrive As String = Environment.GetEnvironmentVariable("SystemDrive")
            WriteRow(sw, "Disk", "System Boot Drive", systemDrive, "", "System")
        Catch ex As Exception
            WriteRow(sw, "Disk", "System Boot Drive Error", ex.Message, "", "System")
        End Try
    End Sub

    Private Shared Sub WriteGPUInfo(sw As StreamWriter)
        Try
            Using mos As New ManagementObjectSearcher("SELECT Name, AdapterRAM, DriverVersion FROM Win32_VideoController")
                Dim idx As Integer = 1
                For Each mo As ManagementObject In mos.Get()
                    Dim name As String = If(mo("Name"), "Unknown").ToString()
                    Dim vramBytes As ULong = Convert.ToUInt64(If(mo("AdapterRAM"), 0UL))
                    Dim vramGB As Double = vramBytes / 1024 / 1024 / 1024
                    Dim driverVer As String = If(mo("DriverVersion"), "Unknown").ToString()

                    WriteRow(sw, "GPU", $"GPU {idx} Name", name, "", "WMI:Win32_VideoController")
                    WriteRow(sw, "GPU", $"GPU {idx} VRAM", vramGB.ToString("F2"), "GB", "WMI:Win32_VideoController")
                    WriteRow(sw, "GPU", $"GPU {idx} Driver Version", driverVer, "", "WMI:Win32_VideoController")
                    idx += 1
                Next
            End Using
        Catch ex As Exception
            WriteRow(sw, "GPU", "Error", ex.Message, "", "System")
        End Try
    End Sub

    Private Shared Sub WriteBIOSInfo(sw As StreamWriter)
        Try
            Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_BIOS")
                For Each obj As ManagementObject In searcher.Get()
                    WriteRow(sw, "BIOS", "Vendor", obj("Manufacturer")?.ToString(), "", "WMI:Win32_BIOS")
                    WriteRow(sw, "BIOS", "Version", obj("SMBIOSBIOSVersion")?.ToString(), "", "WMI:Win32_BIOS")
                    WriteRow(sw, "BIOS", "Serial Number", obj("SerialNumber")?.ToString(), "", "WMI:Win32_BIOS")

                    Dim releaseDateRaw = obj("ReleaseDate")?.ToString()
                    Dim releaseDateFormatted = If(String.IsNullOrWhiteSpace(releaseDateRaw), "Unknown", ManagementDateToDate(releaseDateRaw))
                    WriteRow(sw, "BIOS", "Release Date", releaseDateFormatted, "", "WMI:Win32_BIOS")
                Next
            End Using
        Catch ex As Exception
            WriteRow(sw, "BIOS", "Error", ex.Message, "", "System")
        End Try
    End Sub

    Private Shared Function ManagementDateToDate(wmiDate As String) As String
        If String.IsNullOrEmpty(wmiDate) OrElse wmiDate.Length < 14 Then
            Return "Unknown"
        End If

        Try
            Dim year As Integer = Integer.Parse(wmiDate.Substring(0, 4))
            Dim month As Integer = Integer.Parse(wmiDate.Substring(4, 2))
            Dim day As Integer = Integer.Parse(wmiDate.Substring(6, 2))
            Dim hour As Integer = Integer.Parse(wmiDate.Substring(8, 2))
            Dim minute As Integer = Integer.Parse(wmiDate.Substring(10, 2))
            Dim second As Integer = Integer.Parse(wmiDate.Substring(12, 2))
            Dim dt As New DateTime(year, month, day, hour, minute, second)
            Return dt.ToString("yyyy-MM-dd HH:mm:ss")
        Catch
            Return "Invalid Format"
        End Try
    End Function

    Private Shared Sub WriteMotherboardInfo(sw As StreamWriter)
        Try
            Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard")
                For Each obj As ManagementObject In searcher.Get()
                    WriteRow(sw, "Motherboard", "Manufacturer", obj("Manufacturer")?.ToString(), "", "WMI:Win32_BaseBoard")
                    WriteRow(sw, "Motherboard", "Product", obj("Product")?.ToString(), "", "WMI:Win32_BaseBoard")
                    WriteRow(sw, "Motherboard", "Serial Number", obj("SerialNumber")?.ToString(), "", "WMI:Win32_BaseBoard")
                Next
            End Using
        Catch ex As Exception
            WriteRow(sw, "Motherboard", "Error", ex.Message, "", "System")
        End Try
    End Sub

    Private Shared Sub WriteBatteryInfo(sw As StreamWriter)
        Try
            Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_Battery")
            Dim found As Boolean = False

            For Each obj As ManagementObject In searcher.Get()
                found = True
                Dim name = If(obj("Name"), "Unknown").ToString()
                Dim status = If(obj("Status"), "Unknown").ToString()
                Dim charge = If(obj("EstimatedChargeRemaining"), "Unknown").ToString()
                Dim runtime = If(obj("EstimatedRunTime"), "Unknown").ToString()
                Dim statusCode = If(obj("BatteryStatus"), "Unknown").ToString()
                Dim voltage = If(obj("DesignVoltage"), "Unknown").ToString()
                Dim chemistry = GetBatteryChemistry(If(obj("Chemistry"), "").ToString())

                WriteRow(sw, "Battery", "Name", name, "", "WMI:Win32_Battery")
                WriteRow(sw, "Battery", "Status", status, "", "WMI:Win32_Battery")
                WriteRow(sw, "Battery", "Estimated Charge", charge, "%", "WMI:Win32_Battery")
                WriteRow(sw, "Battery", "Estimated Runtime", runtime, "min", "WMI:Win32_Battery")
                WriteRow(sw, "Battery", "Status Code", statusCode, "", "WMI:Win32_Battery")
                WriteRow(sw, "Battery", "Voltage", voltage, "mV", "WMI:Win32_Battery")
                WriteRow(sw, "Battery", "Chemistry", chemistry, "", "WMI:Win32_Battery")
            Next

            If Not found Then
                WriteRow(sw, "Battery", "Status", "Not Detected", "", "WMI:Win32_Battery")
            End If
        Catch ex As Exception
            WriteRow(sw, "Battery", "Error", ex.Message, "", "System")
        End Try
    End Sub

    Private Shared Function GetBatteryChemistry(code As String) As String
        Select Case code
            Case "1" : Return "Other"
            Case "2" : Return "Unknown"
            Case "3" : Return "Lead Acid"
            Case "4" : Return "Nickel Cadmium"
            Case "5" : Return "Nickel Metal Hydride"
            Case "6" : Return "Lithium-ion"
            Case "7" : Return "Zinc air"
            Case "8" : Return "Lithium Polymer"
            Case Else : Return "Unspecified"
        End Select
    End Function
End Class
