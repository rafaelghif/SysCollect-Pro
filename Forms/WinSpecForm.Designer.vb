<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class WinSpecForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(WinSpecForm))
        Me.BtnExecute = New System.Windows.Forms.Button()
        Me.BtnAbout = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.LblCopyright = New System.Windows.Forms.ToolStripStatusLabel()
        Me.BgWorkerExport = New System.ComponentModel.BackgroundWorker()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'BtnExecute
        '
        Me.BtnExecute.Location = New System.Drawing.Point(12, 10)
        Me.BtnExecute.Name = "BtnExecute"
        Me.BtnExecute.Size = New System.Drawing.Size(288, 27)
        Me.BtnExecute.TabIndex = 0
        Me.BtnExecute.Text = "Execute"
        Me.BtnExecute.UseVisualStyleBackColor = True
        '
        'BtnAbout
        '
        Me.BtnAbout.Location = New System.Drawing.Point(12, 43)
        Me.BtnAbout.Name = "BtnAbout"
        Me.BtnAbout.Size = New System.Drawing.Size(288, 27)
        Me.BtnAbout.TabIndex = 1
        Me.BtnAbout.Text = "About"
        Me.BtnAbout.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.LblCopyright})
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 87)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(313, 22)
        Me.StatusStrip1.TabIndex = 2
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'LblCopyright
        '
        Me.LblCopyright.Name = "LblCopyright"
        Me.LblCopyright.Size = New System.Drawing.Size(90, 17)
        Me.LblCopyright.Text = "© 2025 Flontive"
        '
        'BgWorkerExport
        '
        '
        'WinSpecForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(313, 109)
        Me.Controls.Add(Me.BtnExecute)
        Me.Controls.Add(Me.BtnAbout)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "WinSpecForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "SysCollect Pro"
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BtnExecute As Button
    Friend WithEvents BtnAbout As Button
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents LblCopyright As ToolStripStatusLabel
    Friend WithEvents BgWorkerExport As System.ComponentModel.BackgroundWorker
End Class
