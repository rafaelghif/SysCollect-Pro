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

    Private Function ExportSystemInformation() As String
        Dim hostname = Environment.MachineName
        If Not Directory.Exists(BaseDir) Then Directory.CreateDirectory(BaseDir)

        ' Determine/create folder for the current host
        Dim hostnameFolder As String = GetOrCreateHostnameFolder(hostname)

        ' --- Hardware Section ---
        Dim hardwarePath = PrepareOutputFile("Hardware", hostnameFolder)
        Using sw As New StreamWriter(hardwarePath)
            sw.WriteLine("Property,Value")
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
            sw.WriteLine("Property,Value")
            WriteOSInfo(sw)
            WriteInstalledApplications(sw)
            WriteAntivirusStatus(sw)
        End Using

        ' --- License Section ---
        Dim licensePath = PrepareOutputFile("License", hostnameFolder)
        Using sw As New StreamWriter(licensePath)
            sw.WriteLine("Property,Value")
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
        sw.WriteLine($"Hostname,{Environment.MachineName}")
    End Sub

    Private Sub WriteOSInfo(sw As StreamWriter)
        Try
            Dim osName = "Unknown"
            Dim osBuild = ""
            Dim osDisplayVersion = ""

            Using mos As New ManagementObjectSearcher("SELECT Caption FROM Win32_OperatingSystem")
                For Each mo In mos.Get()
                    osName = If(mo("Caption"), "Unknown")
                Next
            End Using

            Using regKey = Registry.LocalMachine.OpenSubKey("SOFTWARE\Microsoft\Windows NT\CurrentVersion")
                If regKey IsNot Nothing Then
                    Dim currentBuild = If(regKey.GetValue("CurrentBuild"), "")
                    Dim ubr = regKey.GetValue("UBR")
                    Dim displayVersion = If(regKey.GetValue("DisplayVersion"), "")
                    osBuild = If(ubr IsNot Nothing, $"{currentBuild}.{ubr}", currentBuild)
                    osDisplayVersion = displayVersion
                End If
            End Using

            sw.WriteLine($"OS,{osName}")
            sw.WriteLine($"OS Build,{osBuild}")
            sw.WriteLine($"OS Display Version,{osDisplayVersion}")
        Catch ex As Exception
            sw.WriteLine($"OS,Error: {ex.Message}")
        End Try
    End Sub

    Private Sub WriteCPUInfo(sw As StreamWriter)
        Try
            Using mos As New ManagementObjectSearcher("SELECT Name, MaxClockSpeed FROM Win32_Processor")
                For Each mo In mos.Get()
                    sw.WriteLine($"CPU,{mo("Name")}")
                    sw.WriteLine($"CPU MaxClockSpeed MHz,{mo("MaxClockSpeed")}")
                Next
            End Using
        Catch ex As Exception
            sw.WriteLine($"CPU,Error: {ex.Message}")
        End Try
    End Sub

    Private Sub WriteRAMInfo(sw As StreamWriter)
        Try
            Using mos As New ManagementObjectSearcher("SELECT BankLabel, Capacity, Speed, SMBIOSMemoryType FROM Win32_PhysicalMemory")
                Dim idx = 1
                For Each mo In mos.Get()
                    Dim bankLabel = If(mo("BankLabel"), "Unknown")
                    Dim capacityGB = If(mo("Capacity"), 0) / 1024 / 1024 / 1024
                    Dim speedMHz = If(mo("Speed"), "Unknown")
                    Dim typeCode = If(mo("SMBIOSMemoryType"), 0)
                    Dim typeStr = GetMemoryTypeString(typeCode)
                    sw.WriteLine($"RAM {idx} Slot,{bankLabel}")
                    sw.WriteLine($"RAM {idx} Capacity (GB),{capacityGB:F2}")
                    sw.WriteLine($"RAM {idx} Speed (MHz),{speedMHz}")
                    sw.WriteLine($"RAM {idx} Type,{typeStr}")
                    idx += 1
                Next
            End Using
        Catch ex As Exception
            sw.WriteLine($"RAM,Error: {ex.Message}")
        End Try
    End Sub

    Private Sub WriteDiskInfo(sw As StreamWriter)
        Try
            Dim driveInfoList As New List(Of Tuple(Of String, String))

            Dim scope As New ManagementScope("\\.\ROOT\Microsoft\Windows\Storage")
            scope.Connect()
            Using mos As New ManagementObjectSearcher(scope, New ObjectQuery("SELECT MediaType, FriendlyName FROM MSFT_PhysicalDisk"))
                For Each mo In mos.Get()
                    Dim friendlyName = If(mo("FriendlyName"), "Unknown").ToString()
                    Dim mediaTypeCode = If(mo("MediaType"), 0)
                    Dim mediaTypeStr = If(mediaTypeCode = 3, "HDD", If(mediaTypeCode = 4, "SSD", If(mediaTypeCode = 5, "SCM", "Unknown")))
                    driveInfoList.Add(Tuple.Create(friendlyName, mediaTypeStr))
                Next
            End Using

            Using mosDisk As New ManagementObjectSearcher("SELECT Model, Size, SerialNumber FROM Win32_DiskDrive")
                Dim idx = 1
                For Each mo In mosDisk.Get()
                    Dim model = If(mo("Model"), "Unknown")
                    Dim serial = If(mo("SerialNumber"), "Unknown")
                    Dim sizeBytes As ULong = If(mo("Size") IsNot Nothing, Convert.ToUInt64(mo("Size")), 0)
                    Dim sizeGB = sizeBytes / 1024 / 1024 / 1024
                    Dim typeStr = "Unknown"
                    ' Step 4: Match Model with FriendlyName to get MediaType
                    For Each info In driveInfoList
                        If model.Contains(info.Item1) OrElse info.Item1.Contains(model) Then
                            typeStr = info.Item2
                            Exit For
                        End If
                    Next
                    sw.WriteLine($"Drive {idx} Model,{model}")
                    sw.WriteLine($"Drive {idx} Serial Number,{serial}")
                    sw.WriteLine($"Drive {idx} Size (GB),{sizeGB:F2}")
                    sw.WriteLine($"Drive {idx} Type,{typeStr}")
                    idx += 1
                Next
            End Using

            WriteDiskPartitionStyleInfo(sw)
            WriteEFIPartitionStatus(sw)
            WriteBootDriveInfo(sw)
        Catch ex As Exception
            WriteDiskInfoFallback(sw, ex)
        End Try
    End Sub

    Private Sub WriteDiskPartitionStyleInfo(sw As StreamWriter)
        Try
            Using searcher As New ManagementObjectSearcher("SELECT DeviceID FROM Win32_DiskDrive")
                For Each drive As ManagementObject In searcher.Get()
                    Dim deviceId = drive("DeviceID").ToString()

                    ' Use MSFT_Disk to get partition style (GPT/MBR)
                    Dim diskNumber As Integer = GetDiskNumberFromDeviceID(deviceId)
                    If diskNumber >= 0 Then
                        Dim scope = New ManagementScope("\\.\ROOT\Microsoft\Windows\Storage")
                        scope.Connect()
                        Dim query = New ObjectQuery($"SELECT PartitionStyle FROM MSFT_Disk WHERE Number={diskNumber}")
                        Using gptSearcher As New ManagementObjectSearcher(scope, query)
                            For Each disk As ManagementObject In gptSearcher.Get()
                                Dim style = CInt(disk("PartitionStyle"))
                                Dim styleName = If(style = 1, "MBR", If(style = 2, "GPT", "RAW/Unknown"))
                                sw.WriteLine($"Disk {diskNumber} Partition Style," & styleName)
                            Next
                        End Using
                    End If
                Next
            End Using
        Catch ex As Exception
            sw.WriteLine("Disk Partition Style,Error: " & ex.Message)
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

    Private Sub WriteEFIPartitionStatus(sw As StreamWriter)
        Try
            Using searcher As New ManagementObjectSearcher("SELECT * FROM Win32_DiskPartition WHERE Type='EFI System Partition'")
                Dim found As Boolean = False
                For Each partition As ManagementObject In searcher.Get()
                    sw.WriteLine("EFI System Partition,Present")
                    found = True
                    Exit For
                Next
                If Not found Then
                    sw.WriteLine("EFI System Partition,Not Found")
                End If
            End Using
        Catch ex As Exception
            sw.WriteLine("EFI System Partition,Error: " & ex.Message)
        End Try
    End Sub

    Private Sub WriteBootDriveInfo(sw As StreamWriter)
        Try
            Dim systemDrive = Environment.GetEnvironmentVariable("SystemDrive")
            sw.WriteLine("System Boot Drive," & systemDrive)
        Catch ex As Exception
            sw.WriteLine("System Boot Drive,Error: " & ex.Message)
        End Try
    End Sub

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
                If adapter.NetworkInterfaceType = NetworkInterfaceType.Ethernet OrElse adapter.NetworkInterfaceType = NetworkInterfaceType.Wireless80211 Then
                    Dim macAddress = adapter.GetPhysicalAddress().ToString()
                    Dim properties = adapter.GetIPProperties()
                    Dim ipv4 = "N/A"
                    For Each ip In properties.UnicastAddresses
                        If ip.Address.AddressFamily = Net.Sockets.AddressFamily.InterNetwork Then
                            ipv4 = ip.Address.ToString()
                            Exit For
                        End If
                    Next
                    sw.WriteLine($"NIC {idx} Name,{adapter.Name}")
                    sw.WriteLine($"NIC {idx} Description,{adapter.Description}")
                    sw.WriteLine($"NIC {idx} Type,{adapter.NetworkInterfaceType}")
                    sw.WriteLine($"NIC {idx} MAC,{macAddress}")
                    sw.WriteLine($"NIC {idx} IP,{ipv4}")
                    sw.WriteLine($"NIC {idx} Status,{adapter.OperationalStatus}")
                    idx += 1
                End If
            Next
        Catch ex As Exception
            sw.WriteLine($"Network,Error: {ex.Message}")
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
                    sw.WriteLine($"Domain Status,{domainStatus}")
                    sw.WriteLine($"Domain Name,{domainName}")
                Next
            End Using
        Catch ex As Exception
            sw.WriteLine($"Domain,Error: {ex.Message}")
        End Try
    End Sub

    Private Sub WriteLoggedInUser(sw As StreamWriter)
        Try
            sw.WriteLine($"Logged-In User,{Environment.UserDomainName}\{Environment.UserName}")
        Catch ex As Exception
            sw.WriteLine($"Logged-In User,Error: {ex.Message}")
        End Try
    End Sub

    Private Sub WriteInstalledApplications(sw As StreamWriter)
        Try
            Dim regViews = {RegistryView.Registry64, RegistryView.Registry32}
            Dim idx = 1
            For Each view In regViews
                Using baseKey = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, view)
                    Using uninstallKey = baseKey.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall")
                        If uninstallKey IsNot Nothing Then
                            For Each subKeyName In uninstallKey.GetSubKeyNames()
                                Using subKey = uninstallKey.OpenSubKey(subKeyName)
                                    If subKey IsNot Nothing Then
                                        Dim displayName = If(subKey.GetValue("DisplayName"), "")
                                        Dim version = If(subKey.GetValue("DisplayVersion"), "")
                                        If Not String.IsNullOrWhiteSpace(displayName) Then
                                            sw.WriteLine($"App {idx} Name,{displayName}")
                                            sw.WriteLine($"App {idx} Version,{version}")
                                            idx += 1
                                        End If
                                    End If
                                End Using
                            Next
                        End If
                    End Using
                End Using
            Next
        Catch ex As Exception
            sw.WriteLine($"Installed Applications,Error: {ex.Message}")
        End Try
    End Sub

    Private Sub WriteTPMInfo(sw As StreamWriter)
        Try
            Using searcher As New ManagementObjectSearcher("root\CIMV2\Security\MicrosoftTpm", "SELECT * FROM Win32_Tpm")
                Dim found = False
                For Each mo As ManagementObject In searcher.Get()
                    found = True
                    Dim isPresent = If(mo("IsEnabled_InitialValue") IsNot Nothing, mo("IsEnabled_InitialValue").ToString(), "Unknown")
                    Dim version = If(mo("SpecVersion") IsNot Nothing, mo("SpecVersion").ToString(), "Unknown")
                    sw.WriteLine($"TPM Present,Yes")
                    sw.WriteLine($"TPM Version,{version}")
                    sw.WriteLine($"TPM Enabled,{isPresent}")
                Next
                If Not found Then
                    sw.WriteLine("TPM Present,No")
                End If
            End Using
        Catch ex As Exception
            sw.WriteLine($"TPM,Error: {ex.Message}")
        End Try
    End Sub

    Private Sub WriteGPUInfo(sw As StreamWriter)
        Try
            Using mos As New ManagementObjectSearcher("SELECT Name, AdapterRAM, DriverVersion FROM Win32_VideoController")
                Dim idx = 1
                For Each mo In mos.Get()
                    Dim name = If(mo("Name"), "Unknown")
                    Dim vramBytes As ULong = If(mo("AdapterRAM"), 0UL)
                    Dim vramGB = vramBytes / 1024 / 1024 / 1024
                    Dim driverVer = If(mo("DriverVersion"), "Unknown")
                    sw.WriteLine($"GPU {idx} Name,{name}")
                    sw.WriteLine($"GPU {idx} VRAM (GB),{vramGB:F2}")
                    sw.WriteLine($"GPU {idx} Driver Version,{driverVer}")
                    idx += 1
                Next
            End Using
        Catch ex As Exception
            sw.WriteLine($"GPU,Error: {ex.Message}")
        End Try
    End Sub

    Private Sub WriteBitLockerInfo(sw As StreamWriter)
        Try
            Dim scope As New ManagementScope("\\.\ROOT\CIMV2\Security\MicrosoftVolumeEncryption")
            scope.Connect()
            Using searcher As New ManagementObjectSearcher(scope, New ObjectQuery("SELECT DriveLetter, ProtectionStatus FROM Win32_EncryptableVolume"))
                Dim idx = 1
                For Each mo As ManagementObject In searcher.Get()
                    Dim driveLetter = If(mo("DriveLetter"), "Unknown")
                    Dim protectionStatus = If(mo("ProtectionStatus"), 2)
                    Dim statusStr = If(protectionStatus = 0, "Off", If(protectionStatus = 1, "On", "Unknown"))
                    sw.WriteLine($"BitLocker {idx} Drive,{driveLetter}")
                    sw.WriteLine($"BitLocker {idx} Status,{statusStr}")
                    idx += 1
                Next
            End Using
        Catch ex As Exception
            sw.WriteLine($"BitLocker,Error: {ex.Message}")
        End Try
    End Sub

    Private Sub WriteAntivirusStatus(sw As StreamWriter)
        Try
            Using mos As New ManagementObjectSearcher("root\SecurityCenter2", "SELECT displayName, productState FROM AntiVirusProduct")
                Dim idx As Integer = 1
                For Each mo As ManagementObject In mos.Get()
                    Dim name As String = If(mo("displayName"), "Unknown")
                    Dim productState As Integer = Convert.ToInt32(mo("productState"))
                    Dim hexState As String = $"0x{productState:X}"

                    ' Breakdown productState
                    Dim statusByte As Integer = productState And &HFF
                    Dim reportingByte As Integer = (productState >> 8) And &HFF
                    Dim definitionByte As Integer = (productState >> 16) And &HFF

                    ' Interpret status
                    Dim statusStr As String
                    If statusByte = 0 Then
                        statusStr = "Not Reporting or Passive Mode"
                    ElseIf (statusByte And &H10) = &H10 Then
                        statusStr = If((definitionByte And &H10) = &H10, "Enabled and Up-to-date", "Enabled but Outdated")
                    Else
                        statusStr = "Unknown Status"
                    End If

                    ' Write to CSV
                    sw.WriteLine($"Antivirus {idx} Name,{name}")
                    sw.WriteLine($"Antivirus {idx} Status,{statusStr}")
                    sw.WriteLine($"Antivirus {idx} ProductState (Hex),{hexState}")
                    sw.WriteLine($"Antivirus {idx} Status Byte,{statusByte}")
                    sw.WriteLine($"Antivirus {idx} Reporting Byte,{reportingByte}")
                    sw.WriteLine($"Antivirus {idx} Definitions Byte,{definitionByte}")
                    idx += 1
                Next
            End Using
        Catch ex As Exception
            sw.WriteLine($"Antivirus,Error: {ex.Message}")
        End Try
    End Sub

    Private Sub WriteBIOSInfo(sw As StreamWriter)
        Try
            sw.WriteLine("BIOS Information,")
            Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_BIOS")
            For Each obj As ManagementObject In searcher.Get()
                sw.WriteLine("BIOS Vendor," & obj("Manufacturer"))
                sw.WriteLine("BIOS Version," & obj("SMBIOSBIOSVersion"))
                sw.WriteLine("BIOS Serial Number," & obj("SerialNumber"))
                sw.WriteLine("BIOS Release Date," & ManagementDateToDate(obj("ReleaseDate")))
            Next
        Catch ex As Exception
            sw.WriteLine("BIOS Info Error," & ex.Message)
        End Try
    End Sub

    Private Sub WriteMotherboardInfo(sw As StreamWriter)
        Try
            sw.WriteLine("Motherboard Information,")
            Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_BaseBoard")
            For Each obj As ManagementObject In searcher.Get()
                sw.WriteLine("Motherboard Manufacturer," & obj("Manufacturer"))
                sw.WriteLine("Motherboard Product," & obj("Product"))
                sw.WriteLine("Motherboard Serial Number," & obj("SerialNumber"))
            Next
        Catch ex As Exception
            sw.WriteLine("Motherboard Info Error," & ex.Message)
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
            sw.WriteLine("Windows License Info,")

            Dim searcher As New ManagementObjectSearcher(
            "SELECT * FROM SoftwareLicensingProduct WHERE PartialProductKey IS NOT NULL AND LicenseStatus IS NOT NULL")

            For Each obj As ManagementObject In searcher.Get()
                Dim name = TryCast(obj("Name"), String)
                Dim key = TryCast(obj("PartialProductKey"), String)
                Dim statusCode = Convert.ToInt32(obj("LicenseStatus"))
                Dim statusDesc = GetLicenseStatusString(statusCode)

                If Not String.IsNullOrEmpty(name) Then sw.WriteLine("Product Name," & name)
                If Not String.IsNullOrEmpty(key) Then sw.WriteLine("Partial Product Key," & key)
                sw.WriteLine("License Status," & statusDesc)
                Exit For ' Typically only one relevant item
            Next
        Catch ex As Exception
            sw.WriteLine("Windows License Info,Error retrieving license info: " & ex.Message)
        End Try
    End Sub

    Private Sub WriteOfficeLicenseInfo(sw As StreamWriter)
        sw.WriteLine("Microsoft Office License Info,")

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
                                            sw.WriteLine("Version," & officeVersionName)
                                            sw.WriteLine("Product Name," & productName)
                                            sw.WriteLine("Product ID," & If(String.IsNullOrEmpty(productId), "N/A", productId))
                                            sw.WriteLine("Partial Key (Base64?),<Hidden or Not Readable>")
                                            sw.WriteLine("") ' Spacer
                                            foundAny = True
                                        End If
                                    End If
                                End Using
                            Next
                        End If
                    End Using
                Catch ex As Exception
                    sw.WriteLine("Error reading registry (" & officeVersionName & "): " & ex.Message)
                End Try
            Next
        Next

        If Not foundAny Then
            sw.WriteLine("No Office license information found.")
        End If
    End Sub

    Private Sub WriteOfficeLicenseViaOSPP(sw As StreamWriter)
        sw.WriteLine("Microsoft Office Activation Status (via ospp.vbs),")

        Dim officePaths As String() = {
        "C:\Program Files\Microsoft Office",
        "C:\Program Files (x86)\Microsoft Office"
    }

        Dim found As Boolean = False

        For Each basePath In officePaths
            If Directory.Exists(basePath) Then
                Dim officeDirs = Directory.GetDirectories(basePath, "Office*")

                For Each dir As String In officeDirs
                    Dim scriptPath = Path.Combine(dir, "ospp.vbs")
                    If File.Exists(scriptPath) Then
                        found = True

                        Dim psi As New ProcessStartInfo("cscript.exe") With {
                            .Arguments = "//Nologo """ & scriptPath & """ /dstatus",
                            .CreateNoWindow = True,
                            .UseShellExecute = False,
                            .RedirectStandardOutput = True
                        }

                        Using proc As Process = Process.Start(psi)
                            Dim output As String = proc.StandardOutput.ReadToEnd()
                            proc.WaitForExit()

                            sw.WriteLine(output.Replace(vbCrLf, Environment.NewLine))
                            sw.WriteLine("")
                        End Using

                        Exit For
                    End If
                Next
            End If

            If found Then Exit For
        Next

        If Not found Then
            sw.WriteLine("ospp.vbs not found on this machine (Office may be Click-to-Run or not installed).")
        End If
    End Sub

    Private Sub WriteBatteryInfo(sw As StreamWriter)
        Try
            Dim searcher As New ManagementObjectSearcher("SELECT * FROM Win32_Battery")
            Dim found As Boolean = False

            For Each obj As ManagementObject In searcher.Get()
                found = True
                sw.WriteLine("Battery Name," & obj("Name"))
                sw.WriteLine("Battery Status," & obj("Status"))
                sw.WriteLine("Battery Estimated Charge (%)," & obj("EstimatedChargeRemaining") & "%")
                sw.WriteLine("Battery Estimated Runtime (min)," & obj("EstimatedRunTime"))
                sw.WriteLine("Battery Status Code," & obj("BatteryStatus"))
                sw.WriteLine("Battery Voltage (mV)," & obj("DesignVoltage"))
                sw.WriteLine("Battery Chemistry," & GetBatteryChemistry(CStr(obj("Chemistry"))))
            Next

            If Not found Then
                sw.WriteLine("Battery,Not Detected")
            End If
        Catch ex As Exception
            sw.WriteLine("Battery Info,Error: " & ex.Message)
        End Try
    End Sub

    Private Sub WriteVirtualizationInfo(sw As StreamWriter)
        Try
            Dim manufacturer = ""
            Dim model = ""
            Dim biosVersion = ""
            Dim baseboardProduct = ""

            ' Win32_ComputerSystem
            Using csSearcher As New ManagementObjectSearcher("SELECT Manufacturer, Model FROM Win32_ComputerSystem")
                For Each obj As ManagementObject In csSearcher.Get()
                    manufacturer = obj("Manufacturer")?.ToString()
                    model = obj("Model")?.ToString()
                Next
            End Using

            ' Win32_BIOS
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

            ' Win32_BaseBoard
            Using bbSearcher As New ManagementObjectSearcher("SELECT Product FROM Win32_BaseBoard")
                For Each obj As ManagementObject In bbSearcher.Get()
                    baseboardProduct = obj("Product")?.ToString()
                Next
            End Using

            Dim vmType = DetectVMPlatform(manufacturer, model, biosVersion, baseboardProduct)
            sw.WriteLine("Virtualization Platform," & If(String.IsNullOrEmpty(vmType), "Physical Machine", vmType))

        Catch ex As Exception
            sw.WriteLine("Virtualization Platform,Error: " & ex.Message)
        End Try
    End Sub

    Private Sub WriteSystemIdentifiersInfo(sw As StreamWriter)
        Try
            ' --- Machine UUID ---
            Dim uuid As String = ""
            Using uuidSearcher As New ManagementObjectSearcher("SELECT UUID FROM Win32_ComputerSystemProduct")
                For Each obj As ManagementObject In uuidSearcher.Get()
                    uuid = obj("UUID")?.ToString()
                Next
            End Using
            sw.WriteLine("Machine UUID," & uuid)

            ' --- System Serial Number ---
            Dim serialNumber As String = ""
            Using biosSearcher As New ManagementObjectSearcher("SELECT SerialNumber FROM Win32_BIOS")
                For Each obj As ManagementObject In biosSearcher.Get()
                    serialNumber = obj("SerialNumber")?.ToString()
                Next
            End Using
            sw.WriteLine("System Serial Number," & serialNumber)

            ' --- BIOS Mode (UEFI / Legacy) ---
            Dim biosMode As String = "Unknown"
            Dim firmwareType = GetFirmwareType()
            If firmwareType.HasValue Then
                biosMode = If(firmwareType.Value = 2, "UEFI", "Legacy")
            End If
            sw.WriteLine("BIOS Mode," & biosMode)

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
            sw.WriteLine("Chassis Type," & chassisType)

        Catch ex As Exception
            sw.WriteLine("System Identifier Info,Error: " & ex.Message)
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


    Private Function DetectVMPlatform(manufacturer As String, model As String, biosVersion As String, baseboardProduct As String) As String
        Dim idString = (manufacturer & " " & model & " " & biosVersion & " " & baseboardProduct).ToLower()

        If idString.Contains("vmware") Then Return "VMware"
        If idString.Contains("virtualbox") Then Return "VirtualBox"
        If idString.Contains("xen") Then Return "Xen"
        If idString.Contains("kvm") Then Return "KVM"
        If idString.Contains("hyper-v") Or idString.Contains("microsoft corporation") AndAlso idString.Contains("virtual") Then Return "Hyper-V"
        If idString.Contains("qemu") Then Return "QEMU"
        If idString.Contains("parallels") Then Return "Parallels"
        If idString.Contains("bochs") Then Return "Bochs"

        Return Nothing ' Assume physical machine
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
            Case 2 : Return "Out-of-Box Grace Period"
            Case 3 : Return "Out-of-Tolerance Grace Period"
            Case 4 : Return "Non-Genuine Grace Period"
            Case 5 : Return "Notification Mode"
            Case 6 : Return "Extended Grace Period"
            Case Else : Return "Unknown Status"
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
