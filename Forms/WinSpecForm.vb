Imports System.ComponentModel
Imports System.IO
Imports System.Management
Imports System.Net.NetworkInformation
Imports System.Text.RegularExpressions
Imports Microsoft.Win32

Public Class WinSpecForm
    Private exportResultPath As String = ""
    Private ReadOnly BaseDir As String = "C:\System Information"
    Private progressForm As ProgressForm

    Private Sub BtnExecute_Click(sender As Object, e As EventArgs) Handles BtnExecute.Click
        BtnExecute.Enabled = False

        ' Show Progress Form non-blocking
        progressForm = New ProgressForm()
        progressForm.Show(Me)
        progressForm.Refresh()

        exportResultPath = ""
        BgWorkerExport.RunWorkerAsync()
    End Sub

    Private Sub BgWorkerExport_DoWork(sender As Object, e As DoWorkEventArgs) Handles BgWorkerExport.DoWork
        Try
            exportResultPath = ExportSystemInformation()
        Catch ex As Exception
            e.Result = ex
        End Try
    End Sub

    Private Sub BgWorkerExport_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles BgWorkerExport.RunWorkerCompleted
        BtnExecute.Enabled = True

        BtnExecute.Enabled = True
        If progressForm IsNot Nothing Then
            progressForm.Close()
            progressForm.Dispose()
            progressForm = Nothing
        End If

        If e.Result IsNot Nothing AndAlso TypeOf e.Result Is Exception Then
            Dim ex As Exception = DirectCast(e.Result, Exception)
            MessageBox.Show("Error writing system info:" & Environment.NewLine & ex.Message, "Export Failed", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If

        If Not String.IsNullOrEmpty(exportResultPath) Then
            Clipboard.SetText(exportResultPath)

            Dim msg As String = String.Join(Environment.NewLine, {
                "System information has been saved successfully.",
                "",
                "Output Folder Structure:",
                "- Hardware",
                "- Software",
                "- License",
                "",
                "Folder path has been copied to clipboard.",
                "",
                "Do you want to open the folder now?"
            })

            Dim result = MessageBox.Show(msg, "Export Complete", MessageBoxButtons.YesNo, MessageBoxIcon.Information)
            If result = DialogResult.Yes Then
                Process.Start("explorer.exe", exportResultPath)
            End If

            Close()
        End If
    End Sub

    Private Sub WriteRow(sw As StreamWriter, category As String, key As String, value As String, Optional unit As String = "", Optional source As String = "System.Environment")
        Dim collectedAt As String = Now.ToString("yyyy-MM-dd HH:mm:ss")
        sw.WriteLine($"{category},{key},{value},{unit},{collectedAt},{source}")
    End Sub

    Private Function ExportSystemInformation() As String
        Dim hostname = Environment.MachineName
        If Not Directory.Exists(BaseDir) Then Directory.CreateDirectory(BaseDir)

        ' Determine/create folder for the current host
        Dim hostnameFolder As String = GetOrCreateHostnameFolder(hostname)

        ' --- Hardware Section ---
        Dim hardwarePath = PrepareOutputFile("Hardware", hostnameFolder)
        Using sw As New StreamWriter(hardwarePath)
            sw.WriteLine("Category,Key,Value,Unit,CollectedAt,Source")
            WriteHostname(sw)
            WriteLoggedInUser(sw)
            WriteDomainInfo(sw)

            ' System Identity Block
            WriteSystemIdentifiersInfo(sw)

            ' Platform & Core
            WriteBIOSInfo(sw) ' Trim to BIOS Version, Release Date only
            WriteMotherboardInfo(sw) ' Motherboard Model/Product
            WriteCPUInfo(sw)
            WriteRAMInfo(sw)
            WriteDiskInfo(sw)
            WriteGPUInfo(sw)

            ' Chassis & Power
            WriteBatteryInfo(sw)

            ' Security & Virtualization
            WriteTPMInfo(sw)
            WriteBitLockerInfo(sw)
            WriteVirtualizationInfo(sw)

            ' Networking
            WriteNetworkInfo(sw)
        End Using

        ' --- Software Section ---
        Dim softwarePath = PrepareOutputFile("Software", hostnameFolder)
        Using sw As New StreamWriter(softwarePath)
            sw.WriteLine("Category,Key,Value,Unit,CollectedAt,Source")
            WriteOSInfo(sw)
            WriteInstalledApplications(sw)
            WriteAntivirusStatus(sw)
        End Using

        ' --- License Section ---
        Dim licensePath = PrepareOutputFile("License", hostnameFolder)
        Using sw As New StreamWriter(licensePath)
            sw.WriteLine("Category,Key,Value,Unit,CollectedAt,Source")
            WriteWindowsLicenseInfo(sw)
            WriteOfficeLicenseInfo(sw)
            WriteOfficeLicenseViaOSPP(sw)
        End Using

        Return hostnameFolder
    End Function

    Private Function PrepareOutputFile(outputType As String, hostnameFolder As String) As String
        Dim hostname = Environment.MachineName

        Dim validTypes = {"hardware", "software", "license"}
        If Not validTypes.Contains(outputType.ToLower()) Then
            Throw New ArgumentException("OutputInformationType must be 'Hardware', 'Software', or 'License'")
        End If

        ' Ensure subfolder exists
        Dim typeFolder = Path.Combine(hostnameFolder, outputType)
        If Not Directory.Exists(typeFolder) Then Directory.CreateDirectory(typeFolder)

        ' Generate filename with sequence
        Dim files = Directory.GetFiles(typeFolder, hostname & "_*.csv")
        Dim maxSeq = -1
        Dim fileRegex = New Regex(hostname & "_(\d{3})\.csv")
        For Each f In files
            Dim m = fileRegex.Match(Path.GetFileName(f))
            If m.Success Then
                Dim seq = Integer.Parse(m.Groups(1).Value)
                If seq > maxSeq Then maxSeq = seq
            End If
        Next

        Dim newSeq = maxSeq + 1
        Dim filename = $"{hostname}_{newSeq:D3}.csv"
        Return Path.Combine(typeFolder, filename)
    End Function

    Private Function GetOrCreateHostnameFolder(hostname As String) As String
        Dim existingDirs = Directory.GetDirectories(BaseDir)
        Dim maxFolderSeq = 0
        Dim folderRegex = New Regex("^(\d{2})\. " & Regex.Escape(hostname) & "$")
        Dim hostnameFolder As String = ""

        For Each dir As String In existingDirs
            Dim dirName = Path.GetFileName(dir)
            Dim m = folderRegex.Match(dirName)
            If m.Success Then
                Dim seq = Integer.Parse(m.Groups(1).Value)
                If seq > maxFolderSeq Then
                    maxFolderSeq = seq
                    hostnameFolder = dir
                End If
            End If
        Next

        If String.IsNullOrEmpty(hostnameFolder) Then
            Dim newFolderSeq = maxFolderSeq + 1
            Dim newFolderName = $"{newFolderSeq:D2}. {hostname}"
            hostnameFolder = Path.Combine(BaseDir, newFolderName)
            Directory.CreateDirectory(hostnameFolder)
        End If

        Return hostnameFolder
    End Function

    Private Sub WriteHostname(sw As StreamWriter)
        Try
            Dim hostname As String = Environment.MachineName
            WriteRow(sw, "System Identity", "Hostname", hostname)
        Catch ex As Exception
            WriteRow(sw, "System Identity", "Hostname", $"Error: {ex.Message}", "", "System.Environment")
        End Try
    End Sub

    Private Sub WriteOSInfo(sw As StreamWriter)
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

    Private Sub WriteCPUInfo(sw As StreamWriter)
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

    Private Sub WriteRAMInfo(sw As StreamWriter)
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

    Private Sub WriteDiskInfo(sw As StreamWriter)
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

    Private Sub WriteDiskPartitionStyleInfo(sw As StreamWriter)
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

    Private Sub WriteEFIPartitionStatus(sw As StreamWriter)
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

    Private Sub WriteBootDriveInfo(sw As StreamWriter)
        Try
            Dim systemDrive As String = Environment.GetEnvironmentVariable("SystemDrive")
            WriteRow(sw, "Disk", "System Boot Drive", systemDrive, "", "System")
        Catch ex As Exception
            WriteRow(sw, "Disk", "System Boot Drive Error", ex.Message, "", "System")
        End Try
    End Sub

    Private Function GetDiskNumberFromDeviceID(deviceId As String) As Integer
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

    Private Sub WriteDiskInfoFallback(sw As StreamWriter, ex As Exception)
        Try
            Using mosDisk As New ManagementObjectSearcher("SELECT Model, Size, SerialNumber FROM Win32_DiskDrive")
                Dim idx = 1
                For Each mo In mosDisk.Get()
                    Dim model = If(mo("Model"), "Unknown")
                    Dim serial = If(mo("SerialNumber"), "Unknown")
                    Dim sizeBytes As ULong = If(mo("Size") IsNot Nothing, Convert.ToUInt64(mo("Size")), 0)
                    Dim sizeGB = sizeBytes / 1024 / 1024 / 1024
                    sw.WriteLine($"Drive {idx} Model,{model}")
                    sw.WriteLine($"Drive {idx} Serial Number,{serial}")
                    sw.WriteLine($"Drive {idx} Size (GB),{sizeGB:F2}")
                    sw.WriteLine($"Drive {idx} Type,Unknown")
                    idx += 1
                Next
            End Using
        Catch ex2 As Exception
            sw.WriteLine($"Storage,Error: {ex2.Message}")
        End Try
    End Sub

    Private Sub WriteNetworkInfo(sw As StreamWriter)
        Try
            Dim adapters = NetworkInterface.GetAllNetworkInterfaces()
            Dim idx = 1

            For Each adapter In adapters
                If adapter.NetworkInterfaceType = NetworkInterfaceType.Ethernet OrElse
               adapter.NetworkInterfaceType = NetworkInterfaceType.Wireless80211 Then

                    Dim macAddress = adapter.GetPhysicalAddress().ToString()
                    Dim properties = adapter.GetIPProperties()

                    ' Get IPv4 and Domain
                    Dim ipv4 As String = ""
                    Dim dnsSuffix As String = If(properties.DnsSuffix, "")

                    Dim ipList As New List(Of String)
                    For Each ip In properties.UnicastAddresses
                        If ip.Address.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then
                            ipList.Add(ip.Address.ToString())
                        End If
                    Next
                    ipv4 = If(ipList.Count > 0, String.Join(";", ipList), "N/A")

                    ' Collect data
                    WriteRow(sw, $"NIC {idx}", "Name", adapter.Name, "", "System.Net")
                    WriteRow(sw, $"NIC {idx}", "Description", adapter.Description, "", "System.Net")
                    WriteRow(sw, $"NIC {idx}", "Type", adapter.NetworkInterfaceType.ToString(), "", "System.Net")
                    WriteRow(sw, $"NIC {idx}", "MAC", macAddress, "", "System.Net")
                    WriteRow(sw, $"NIC {idx}", "IP (v4)", ipv4, "", "System.Net")
                    WriteRow(sw, $"NIC {idx}", "Speed (Mbps)", (adapter.Speed \ 1000000).ToString(), "", "System.Net")
                    WriteRow(sw, $"NIC {idx}", "DNS Suffix", dnsSuffix, "", "System.Net")
                    WriteRow(sw, $"NIC {idx}", "Status", adapter.OperationalStatus.ToString(), "", "System.Net")
                    idx += 1
                End If
            Next

            If idx = 1 Then
                WriteRow(sw, "Network", "Info", "No Ethernet/Wireless adapters found", "", "System.Net")
            End If

        Catch ex As Exception
            WriteRow(sw, "Network", "Error", ex.Message, "", "System.Net")
        End Try
    End Sub

    Private Sub WriteSystemInfo(sw As StreamWriter)
        Try
            Using mos As New ManagementObjectSearcher("SELECT SerialNumber, SMBIOSBIOSVersion, ReleaseDate FROM Win32_BIOS")
                For Each mo In mos.Get()
                    sw.WriteLine($"BIOS Serial Number,{mo("SerialNumber")}")
                    sw.WriteLine($"BIOS Version,{mo("SMBIOSBIOSVersion")}")
                    Dim rawDate = mo("ReleaseDate")
                    If Not String.IsNullOrEmpty(rawDate) AndAlso rawDate.Length >= 8 Then
                        Dim biosDate = rawDate.Substring(6, 2) & "-" & rawDate.Substring(4, 2) & "-" & rawDate.Substring(0, 4)
                        sw.WriteLine($"BIOS Release Date,{biosDate}")
                    End If
                Next
            End Using
            Using mos2 As New ManagementObjectSearcher("SELECT Manufacturer, Model FROM Win32_ComputerSystem")
                For Each mo In mos2.Get()
                    sw.WriteLine($"Manufacturer,{mo("Manufacturer")}")
                    sw.WriteLine($"Model,{mo("Model")}")
                Next
            End Using
        Catch ex As Exception
            sw.WriteLine($"System Info,Error: {ex.Message}")
        End Try
    End Sub

    Private Sub WriteDomainInfo(sw As StreamWriter)
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

    Private Sub WriteLoggedInUser(sw As StreamWriter)
        Try
            Dim loggedUser As String = $"{Environment.UserDomainName}\{Environment.UserName}"
            WriteRow(sw, "System Identity", "Logged-In User", loggedUser, "", "System.Environment")
        Catch ex As Exception
            WriteRow(sw, "System Identity", "Logged-In User", $"Error: {ex.Message}", "", "System.Environment")
        End Try
    End Sub

    Private Sub WriteInstalledApplications(sw As StreamWriter)
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

    Private Sub WriteTPMInfo(sw As StreamWriter)
        Try
            Using searcher As New ManagementObjectSearcher("root\CIMV2\Security\MicrosoftTpm", "SELECT * FROM Win32_Tpm")
                Dim found = False

                For Each mo As ManagementObject In searcher.Get()
                    found = True

                    ' IsEnabled_InitialValue
                    Dim isEnabledStr As String = "Unknown"
                    If mo("IsEnabled_InitialValue") IsNot Nothing Then
                        Dim isEnabled As Boolean = Convert.ToBoolean(mo("IsEnabled_InitialValue"))
                        isEnabledStr = If(isEnabled, "Yes", "No")
                    End If

                    ' SpecVersion
                    Dim version As String = If(mo("SpecVersion") IsNot Nothing, mo("SpecVersion").ToString(), "Unknown")

                    WriteRow(sw, "TPM", "Present", "Yes", "", "WMI:Win32_Tpm")
                    WriteRow(sw, "TPM", "Version", version, "", "WMI:Win32_Tpm")
                    WriteRow(sw, "TPM", "Enabled (InitialValue)", isEnabledStr, "", "WMI:Win32_Tpm")
                Next

                If Not found Then
                    WriteRow(sw, "TPM", "Present", "No", "", "WMI:Win32_Tpm")
                End If
            End Using
        Catch ex As Exception
            WriteRow(sw, "TPM", "Error", ex.Message, "", "WMI:Win32_Tpm")
        End Try
    End Sub

    Private Sub WriteGPUInfo(sw As StreamWriter)
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

    Private Sub WriteBitLockerInfo(sw As StreamWriter)
        Try
            Dim scope As New ManagementScope("\\.\root\CIMV2\Security\MicrosoftVolumeEncryption")
            scope.Connect()

            Dim query As New ObjectQuery("SELECT * FROM Win32_EncryptableVolume")
            Using searcher As New ManagementObjectSearcher(scope, query)
                Dim idx As Integer = 1

                For Each volume As ManagementObject In searcher.Get()
                    Dim driveLetter As String = TryCast(volume("DriveLetter"), String)
                    driveLetter = If(String.IsNullOrWhiteSpace(driveLetter), "N/A", driveLetter)

                    ' Call GetProtectionStatus method properly
                    Dim outParams As ManagementBaseObject = volume.InvokeMethod("GetProtectionStatus", Nothing, Nothing)
                    Dim protectionStatusCode As UInteger = CUInt(outParams("ProtectionStatus"))
                    Dim protectionStatusText As String = GetBitLockerProtectionStatusText(protectionStatusCode)

                    WriteRow(sw, $"BitLocker Volume {idx}", "Drive Letter", driveLetter, "", "BitLocker")
                    WriteRow(sw, $"BitLocker Volume {idx}", "Protection Status", protectionStatusText, "", "BitLocker")

                    idx += 1
                Next
            End Using

        Catch ex As Exception
            WriteRow(sw, "BitLocker", "Error", ex.Message, "", "BitLocker")
        End Try
    End Sub

    Private Function GetBitLockerProtectionStatusText(status As UInteger) As String
        Select Case status
            Case 0 : Return "Protection Off"
            Case 1 : Return "Protection On"
            Case 2 : Return "Protection Unknown"
            Case Else : Return "Unknown Code (" & status.ToString() & ")"
        End Select
    End Function

    Private Sub WriteAntivirusStatus(sw As StreamWriter)
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

    Private Sub WriteBIOSInfo(sw As StreamWriter)
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

    Private Sub WriteMotherboardInfo(sw As StreamWriter)
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

    Private Sub WriteChassisInfo(sw As StreamWriter)
        Try
            sw.WriteLine("Chassis Information,")
            Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_SystemEnclosure")
            For Each obj As ManagementObject In searcher.Get()
                Dim types = CType(obj("ChassisTypes"), UInt16())
                If types IsNot Nothing AndAlso types.Length > 0 Then
                    sw.WriteLine("Chassis Type," & GetChassisTypeName(types(0)))
                End If
            Next
        Catch ex As Exception
            sw.WriteLine("Chassis Info Error," & ex.Message)
        End Try
    End Sub

    Private Sub WriteWindowsLicenseInfo(sw As StreamWriter)
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

    Private Sub WriteOfficeLicenseInfo(sw As StreamWriter)
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

    Private Sub WriteOfficeLicenseViaOSPP(sw As StreamWriter)
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

    Private Sub WriteBatteryInfo(sw As StreamWriter)
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

    Private Sub WriteVirtualizationInfo(sw As StreamWriter)
        Try
            Dim manufacturer As String = ""
            Dim model As String = ""
            Dim biosVersion As String = ""
            Dim baseboardProduct As String = ""

            ' Gather Win32_ComputerSystem data
            Using csSearcher As New ManagementObjectSearcher("SELECT Manufacturer, Model FROM Win32_ComputerSystem")
                For Each obj As ManagementObject In csSearcher.Get()
                    manufacturer = If(obj("Manufacturer"), "").ToString()
                    model = If(obj("Model"), "").ToString()
                Next
            End Using

            ' Gather Win32_BIOS data
            Using biosSearcher As New ManagementObjectSearcher("SELECT Version FROM Win32_BIOS")
                For Each obj As ManagementObject In biosSearcher.Get()
                    Dim versionObj = obj("Version")
                    If versionObj IsNot Nothing Then
                        If TypeOf versionObj Is String() Then
                            biosVersion = String.Join(" ", CType(versionObj, String()))
                        Else
                            biosVersion = versionObj.ToString()
                        End If
                    End If
                Next
            End Using

            ' Gather Win32_BaseBoard data
            Using bbSearcher As New ManagementObjectSearcher("SELECT Product FROM Win32_BaseBoard")
                For Each obj As ManagementObject In bbSearcher.Get()
                    baseboardProduct = If(obj("Product"), "").ToString()
                Next
            End Using

            ' Determine platform type
            Dim vmType = DetectVMPlatform(manufacturer, model, biosVersion, baseboardProduct)
            Dim isVM = If(String.IsNullOrEmpty(vmType), "Physical Machine", vmType)

            WriteRow(sw, "Virtualization", "Platform", isVM, "", "WMI:Win32_ComputerSystem + BIOS + BaseBoard")
            WriteRow(sw, "Virtualization", "Detected From", $"{manufacturer} | {model} | {biosVersion} | {baseboardProduct}", "", "Derived")

        Catch ex As Exception
            WriteRow(sw, "Virtualization", "Error", ex.Message, "", "WMI")
        End Try
    End Sub

    Private Function DetectVMPlatform(manufacturer As String, model As String, biosVersion As String, baseboardProduct As String) As String
        Dim idString = (manufacturer & " " & model & " " & biosVersion & " " & baseboardProduct).ToLowerInvariant()

        If idString.Contains("vmware") Then Return "VMware"
        If idString.Contains("virtualbox") Then Return "VirtualBox"
        If idString.Contains("xen") Then Return "Xen"
        If idString.Contains("kvm") Then Return "KVM"
        If idString.Contains("qemu") Then Return "QEMU"
        If idString.Contains("parallels") Then Return "Parallels"
        If idString.Contains("bochs") Then Return "Bochs"
        If idString.Contains("hyper-v") OrElse
       (idString.Contains("microsoft corporation") AndAlso idString.Contains("virtual")) Then Return "Hyper-V"

        Return Nothing
    End Function


    Private Sub WriteSystemIdentifiersInfo(sw As StreamWriter)
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
    Private Function GetFirmwareType() As UInteger?
        Try
            Dim fwType As UInteger = 0
            If GetFirmwareType(fwType) Then
                Return fwType ' 1 = BIOS (Legacy), 2 = UEFI
            End If
        Catch
        End Try
        Return Nothing
    End Function

    Private Function GetChassisTypeName(code As UShort) As String
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

    Private Function GetBatteryChemistry(code As String) As String
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


    Private Function GetMemoryTypeString(typeCode As Integer) As String
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

    Private Function GetLicenseStatusString(statusCode As Integer) As String
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

    Private Function ManagementDateToDate(wmiDate As String) As String
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


    Private Sub BtnAbout_Click(sender As Object, e As EventArgs) Handles BtnAbout.Click
        Dim info As String = String.Join(Environment.NewLine, {
              "SysCollect Pro – System Inventory & Diagnostics Utility",
              "",
              "© 2025 Muhammad Rafael Ghifari. All Rights Reserved.",
              "",
              "Developer Contact",
              "Full Name     : Muhammad Rafael Ghifari",
              "Phone / WA    : +62 851-5850-8840",
              "Email         : rafaelghifari.business@gmail.com",
              "Work Email    : muhammad.rafael@yokogawa.com",
              "Organization  : Yokogawa Manufacturing Batam",
              "",
              "Application Purpose",
              "SysCollect Pro is developed exclusively for internal use",
              "within enterprise environments to support IT inventory,",
              "diagnostic collection, and asset compliance.",
              "",
              "Terms of Use",
              "This software is the intellectual property of its author.",
              "Unauthorized distribution, modification, or external use",
              "without written permission is strictly prohibited.",
              "",
              "Support",
              "For support or integration requests, please contact the",
              "developer using the details above.",
              "",
              "Thank you for using SysCollect Pro."
        })

        MessageBox.Show(info, "About SysCollect Pro", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub
End Class
