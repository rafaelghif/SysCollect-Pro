Imports System.IO

Public Class MainCollector
    Public Shared Sub CollectData(hostnameFolder)
        Using sw As New StreamWriter(PrepareOutputFile("Hardware", hostnameFolder))
            SystemInfoCollector.Collect(sw)
            HardwareInfoCollector.Collect(sw)
            SecurityInfoCollector.Collect(sw)
            NetworkInfoCollector.Collect(sw)
        End Using

        ' --- Software ---
        Using sw As New StreamWriter(PrepareOutputFile("Software", hostnameFolder))
            SoftwareInfoCollector.Collect(sw)
        End Using

        ' --- License ---
        Using sw As New StreamWriter(PrepareOutputFile("License", hostnameFolder))
            LicenseInfoCollector.Collect(sw)
        End Using
    End Sub
End Class
