Imports System.ComponentModel
Imports System.IO
Imports System.Text.RegularExpressions

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

        MainCollector.CollectData(hostnameFolder)

        Return hostnameFolder
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
