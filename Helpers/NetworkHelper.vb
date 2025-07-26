Imports System.Net.NetworkInformation
Imports System.Net.Sockets

Public Module NetworkHelper
    Public Function GetServerIp() As String
        Try
            For Each ni As NetworkInterface In NetworkInterface.GetAllNetworkInterfaces()
                If ni.OperationalStatus <> OperationalStatus.Up Then Continue For
                If ni.NetworkInterfaceType = NetworkInterfaceType.Loopback Then Continue For

                Dim ipProps = ni.GetIPProperties()
                For Each addrInfo In ipProps.UnicastAddresses
                    If addrInfo.Address.AddressFamily <> AddressFamily.InterNetwork Then Continue For

                    Dim ip = addrInfo.Address.ToString()

                    If ip.StartsWith("10.137.") Then
                        Return "10.137.1.35"
                    ElseIf ip.StartsWith("192.168.") Then
                        Return "192.168.100.200"
                    End If
                Next
            Next
        Catch ex As Exception
            ' Log or ignore
        End Try

        Return Nothing ' Not in known network
    End Function

End Module
