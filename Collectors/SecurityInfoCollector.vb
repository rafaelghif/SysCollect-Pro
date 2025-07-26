Imports System.IO
Imports System.Management

Public Class SecurityInfoCollector
    Public Shared Sub Collect(sw As StreamWriter)
        WriteRowHeader(sw)
        WriteTPMInfo(sw)
        WriteBitLockerInfo(sw)
        WriteVirtualizationInfo(sw)
    End Sub

    Private Shared Sub WriteTPMInfo(sw As StreamWriter)
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

    Private Shared Sub WriteBitLockerInfo(sw As StreamWriter)
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
                    Dim protectionStatusCode As UInteger = outParams("ProtectionStatus")
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

    Private Shared Function GetBitLockerProtectionStatusText(status As UInteger) As String
        Select Case status
            Case 0 : Return "Protection Off"
            Case 1 : Return "Protection On"
            Case 2 : Return "Protection Unknown"
            Case Else : Return "Unknown Code (" & status.ToString() & ")"
        End Select
    End Function

    Private Shared Sub WriteVirtualizationInfo(sw As StreamWriter)
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

    Private Shared Function DetectVMPlatform(manufacturer As String, model As String, biosVersion As String, baseboardProduct As String) As String
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
End Class
