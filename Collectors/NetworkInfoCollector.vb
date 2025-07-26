Imports System.IO
Imports System.Net.NetworkInformation

Public Class NetworkInfoCollector
    Public Shared Sub Collect(sw As StreamWriter)
        WriteRowHeader(sw)
        WriteNetworkInfo(sw)
    End Sub
    Private Shared Sub WriteNetworkInfo(sw As StreamWriter)
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
End Class
