Imports System.IO
Imports System.Text.RegularExpressions

Module CSVWriterHelper
    Public Sub WriteRowHeader(sw As StreamWriter)
        sw.WriteLine("Category,Key,Value,Unit,CollectedAt,Source")
    End Sub

    Public Sub WriteRow(sw As StreamWriter, category As String, key As String, value As String, Optional unit As String = "", Optional source As String = "System.Environment")
        Dim collectedAt As String = Now.ToString("yyyy-MM-dd HH:mm:ss")
        sw.WriteLine($"{category},{key},{value},{unit},{collectedAt},{source}")
    End Sub

    Public Function PrepareOutputFile(outputType As String, hostnameFolder As String) As String
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
End Module
