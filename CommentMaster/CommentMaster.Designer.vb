<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class CommentMaster
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.btnOpenFile = New System.Windows.Forms.Button()
        Me.TabControl1 = New System.Windows.Forms.TabControl()
        Me.tabMainForm = New System.Windows.Forms.TabPage()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtProgramPurpose = New System.Windows.Forms.TextBox()
        Me.txtFIlePurpose = New System.Windows.Forms.TextBox()
        Me.txtAuthorDate = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtAuthorName = New System.Windows.Forms.TextBox()
        Me.txtProjectName = New System.Windows.Forms.TextBox()
        Me.lblProjectName = New System.Windows.Forms.Label()
        Me.txtFormName = New System.Windows.Forms.TextBox()
        Me.lblFormName = New System.Windows.Forms.Label()
        Me.tabConstants = New System.Windows.Forms.TabPage()
        Me.tabControlConstants = New System.Windows.Forms.TabControl()
        Me.tabGlobalStructs = New System.Windows.Forms.TabPage()
        Me.tabControlStructs = New System.Windows.Forms.TabControl()
        Me.tabGlobalVariables = New System.Windows.Forms.TabPage()
        Me.tabControlGlobalVars = New System.Windows.Forms.TabControl()
        Me.tabMethods = New System.Windows.Forms.TabPage()
        Me.tabControlMethods = New System.Windows.Forms.TabControl()

        'Main form
        Me.OpenFileDialog1 = New System.Windows.Forms.OpenFileDialog()
        Me.btnParseFile = New System.Windows.Forms.Button()
        Me.btnSaveFile = New System.Windows.Forms.Button()
        Me.lblCurrentSelectedFile = New System.Windows.Forms.Label()
        Me.lstConsole = New System.Windows.Forms.ListBox()
        Me.TabControl1.SuspendLayout()
        Me.tabMainForm.SuspendLayout()
        Me.tabGlobalStructs.SuspendLayout()
        Me.tabMethods.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnOpenFile
        '
        Me.btnOpenFile.Location = New System.Drawing.Point(36, 30)
        Me.btnOpenFile.Name = "btnOpenFile"
        Me.btnOpenFile.Size = New System.Drawing.Size(258, 46)
        Me.btnOpenFile.TabIndex = 0
        Me.btnOpenFile.Text = "Open File"
        Me.btnOpenFile.UseVisualStyleBackColor = True
        '
        'TabControl1
        '
        Me.TabControl1.Controls.Add(Me.tabMainForm)
        Me.TabControl1.Controls.Add(Me.tabConstants)
        Me.TabControl1.Controls.Add(Me.tabGlobalStructs)
        Me.TabControl1.Controls.Add(Me.tabGlobalVariables)
        Me.TabControl1.Controls.Add(Me.tabMethods)
        Me.TabControl1.Location = New System.Drawing.Point(12, 111)
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.SelectedIndex = 0
        Me.TabControl1.Size = New System.Drawing.Size(2450, 975)
        Me.TabControl1.TabIndex = 1
        '
        'tabMainForm
        '
        Me.tabMainForm.Controls.Add(Me.Label4)
        Me.tabMainForm.Controls.Add(Me.Label3)
        Me.tabMainForm.Controls.Add(Me.txtProgramPurpose)
        Me.tabMainForm.Controls.Add(Me.txtFIlePurpose)
        Me.tabMainForm.Controls.Add(Me.txtAuthorDate)
        Me.tabMainForm.Controls.Add(Me.Label2)
        Me.tabMainForm.Controls.Add(Me.Label1)
        Me.tabMainForm.Controls.Add(Me.txtAuthorName)
        Me.tabMainForm.Controls.Add(Me.txtProjectName)
        Me.tabMainForm.Controls.Add(Me.lblProjectName)
        Me.tabMainForm.Controls.Add(Me.txtFormName)
        Me.tabMainForm.Controls.Add(Me.lblFormName)
        Me.tabMainForm.Location = New System.Drawing.Point(8, 46)
        Me.tabMainForm.Name = "tabMainForm"
        Me.tabMainForm.Padding = New System.Windows.Forms.Padding(3)
        Me.tabMainForm.Size = New System.Drawing.Size(2250, 636)
        Me.tabMainForm.TabIndex = 1
        Me.tabMainForm.Text = "Main Form"
        Me.tabMainForm.UseVisualStyleBackColor = True
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(875, 325)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(200, 32)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Program Purpose"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(925, 25)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(150, 32)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "File Purpose"
        '
        'txtProgramPurpose
        '

        Me.txtProgramPurpose.Location = New System.Drawing.Point(1100, 325)
        Me.txtProgramPurpose.Multiline = True
        Me.txtProgramPurpose.Name = "txtProgramPurpose"
        Me.txtProgramPurpose.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtProgramPurpose.Size = New System.Drawing.Size(1007, 275)
        Me.txtProgramPurpose.TabIndex = 9
        '
        'txtFIlePurpose
        '
        'Me.txtFIlePurpose.Font = New System.Drawing.Font("Cascadia Mono", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.txtFIlePurpose.Location = New System.Drawing.Point(1100, 20)
        Me.txtFIlePurpose.Multiline = True
        Me.txtFIlePurpose.Name = "txtFIlePurpose"
        Me.txtFIlePurpose.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtFIlePurpose.Size = New System.Drawing.Size(1300, 275)
        Me.txtFIlePurpose.TabIndex = 8
        Me.txtFIlePurpose.Text = ""
        '
        'txtAuthorDate
        '
        Me.txtAuthorDate.Location = New System.Drawing.Point(325, 226)
        Me.txtAuthorDate.Name = "txtAuthorDate"
        Me.txtAuthorDate.Size = New System.Drawing.Size(486, 39)
        Me.txtAuthorDate.TabIndex = 7
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(29, 233)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(144, 32)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Author Date"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 177)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(158, 32)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "Author Name"
        '
        'txtAuthorName
        '
        Me.txtAuthorName.Location = New System.Drawing.Point(325, 170)
        Me.txtAuthorName.Name = "txtAuthorName"
        Me.txtAuthorName.Size = New System.Drawing.Size(486, 39)
        Me.txtAuthorName.TabIndex = 4
        '
        'txtProjectName
        '
        Me.txtProjectName.Location = New System.Drawing.Point(325, 76)
        Me.txtProjectName.Name = "txtProjectName"
        Me.txtProjectName.Size = New System.Drawing.Size(486, 39)
        Me.txtProjectName.TabIndex = 3
        '
        'lblProjectName
        '
        Me.lblProjectName.AutoSize = True
        Me.lblProjectName.Location = New System.Drawing.Point(29, 76)
        Me.lblProjectName.Name = "lblProjectName"
        Me.lblProjectName.Size = New System.Drawing.Size(175, 32)
        Me.lblProjectName.TabIndex = 2
        Me.lblProjectName.Text = "Project Name"
        '
        'txtFormName
        '
        Me.txtFormName.Font = New System.Drawing.Font("Cascadia Mono", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        Me.txtFormName.Location = New System.Drawing.Point(325, 23)
        Me.txtFormName.Name = "txtFormName"
        Me.txtFormName.Size = New System.Drawing.Size(486, 39)
        Me.txtFormName.TabIndex = 1
        Me.txtFormName.Enabled = False
        '
        'lblFormName
        '
        Me.lblFormName.AutoSize = True
        Me.lblFormName.Location = New System.Drawing.Point(33, 25)
        Me.lblFormName.Name = "lblFormName"
        Me.lblFormName.Size = New System.Drawing.Size(140, 32)
        Me.lblFormName.TabIndex = 0
        Me.lblFormName.Text = "File Name"


        '
        'tabGlobalStructs
        '
        Me.tabConstants.Controls.Add(Me.tabControlConstants)
        Me.tabConstants.Location = New System.Drawing.Point(8, 46)
        Me.tabConstants.Name = "tabConstants"
        Me.tabConstants.Padding = New System.Windows.Forms.Padding(3)
        Me.tabConstants.Size = New System.Drawing.Size(2123, 636)
        Me.tabConstants.TabIndex = 0
        Me.tabConstants.Text = "Global Constants"
        Me.tabConstants.UseVisualStyleBackColor = True
        '
        'tabControlStructs
        '
        Me.tabControlConstants.Location = New System.Drawing.Point(27, 25)
        Me.tabControlConstants.Name = "tabControlConstants"
        Me.tabControlConstants.SelectedIndex = 0
        Me.tabControlConstants.Size = New System.Drawing.Size(2350, 875)
        Me.tabControlConstants.TabIndex = 0

        '
        'tabGlobalStructs
        '
        Me.tabGlobalStructs.Controls.Add(Me.tabControlStructs)
        Me.tabGlobalStructs.Location = New System.Drawing.Point(8, 46)
        Me.tabGlobalStructs.Name = "tabGlobalStructs"
        Me.tabGlobalStructs.Padding = New System.Windows.Forms.Padding(3)
        Me.tabGlobalStructs.Size = New System.Drawing.Size(2123, 636)
        Me.tabGlobalStructs.TabIndex = 0
        Me.tabGlobalStructs.Text = "Global Structures"
        Me.tabGlobalStructs.UseVisualStyleBackColor = True
        '
        'tabControlStructs
        '
        Me.tabControlStructs.Location = New System.Drawing.Point(27, 25)
        Me.tabControlStructs.Name = "TabControl2"
        Me.tabControlStructs.SelectedIndex = 0
        Me.tabControlStructs.Size = New System.Drawing.Size(2350, 875)
        Me.tabControlStructs.TabIndex = 0
        '
        'tabGlobalVariables
        '
        Me.tabGlobalVariables.Controls.Add(Me.tabControlGlobalVars)
        Me.tabGlobalVariables.Location = New System.Drawing.Point(8, 46)
        Me.tabGlobalVariables.Name = "tabGlobalVariables"
        Me.tabGlobalVariables.Padding = New System.Windows.Forms.Padding(3)
        Me.tabGlobalVariables.Size = New System.Drawing.Size(2123, 636)
        Me.tabGlobalVariables.TabIndex = 2
        Me.tabGlobalVariables.Text = "Global Variables"
        Me.tabGlobalVariables.UseVisualStyleBackColor = True
        '
        'tabControlVars
        '

        Me.tabControlGlobalVars.Location = New System.Drawing.Point(27, 25)
        Me.tabControlGlobalVars.Name = "TabControl2"
        Me.tabControlGlobalVars.SelectedIndex = 0
        Me.tabControlGlobalVars.Size = New System.Drawing.Size(2350, 875)
        Me.tabControlGlobalVars.TabIndex = 0

        '
        'tabSubprograms
        '
        Me.tabMethods.Controls.Add(Me.tabControlMethods)
        Me.tabMethods.Location = New System.Drawing.Point(8, 46)
        Me.tabMethods.Name = "tabSubprograms"
        Me.tabMethods.Size = New System.Drawing.Size(2123, 636)
        Me.tabMethods.TabIndex = 3
        Me.tabMethods.Text = "Methods"
        Me.tabMethods.UseVisualStyleBackColor = True
        '
        'TabControl3
        '
        Me.tabControlMethods.Location = New System.Drawing.Point(27, 25)
        Me.tabControlMethods.Name = "TabControl3"
        Me.tabControlMethods.SelectedIndex = 0
        Me.tabControlMethods.Size = New System.Drawing.Size(2350, 875)
        Me.tabControlMethods.TabIndex = 0
        '
        'OpenFileDialog1
        '
        Me.OpenFileDialog1.FileName = "OpenFileDialogMain"
        '
        'btnParseFile
        '
        Me.btnParseFile.Location = New System.Drawing.Point(2300, 75)
        Me.btnParseFile.Name = "btnParseFile"
        Me.btnParseFile.Size = New System.Drawing.Size(150, 46)
        Me.btnParseFile.TabIndex = 2
        Me.btnParseFile.Text = "Reload"
        Me.btnParseFile.UseVisualStyleBackColor = True
        Me.btnParseFile.Enabled = False
        '
        'btnSaveFile
        '
        Me.btnSaveFile.Location = New System.Drawing.Point(2300, 15)
        Me.btnSaveFile.Name = "btnGenerateComments"
        Me.btnSaveFile.Size = New System.Drawing.Size(150, 46)
        Me.btnSaveFile.TabIndex = 3
        Me.btnSaveFile.Text = "Generate Comments"
        Me.btnSaveFile.Enabled = False
        Me.btnSaveFile.UseVisualStyleBackColor = True
        '
        'lblCurrentSelectedFile
        '
        Me.lblCurrentSelectedFile.AutoSize = True
        Me.lblCurrentSelectedFile.Location = New System.Drawing.Point(328, 35)
        Me.lblCurrentSelectedFile.Name = "lblCurrentSelectedFile"
        Me.lblCurrentSelectedFile.Size = New System.Drawing.Size(196, 32)
        Me.lblCurrentSelectedFile.TabIndex = 4
        Me.lblCurrentSelectedFile.Text = "No file selected..."
        '
        'lstConsole
        '
        Me.lstConsole.FormattingEnabled = True
        Me.lstConsole.ItemHeight = 32
        Me.lstConsole.Location = New System.Drawing.Point(10, 1100)
        Me.lstConsole.Name = "lstConsole"
        Me.lstConsole.Size = New System.Drawing.Size(2450, 292)
        Me.lstConsole.TabIndex = 6
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(13.0!, 32.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(2500, 1500)
        Me.Controls.Add(Me.lstConsole)
        Me.Controls.Add(Me.lblCurrentSelectedFile)
        Me.Controls.Add(Me.btnSaveFile)
        Me.Controls.Add(Me.btnParseFile)
        Me.Controls.Add(Me.TabControl1)
        Me.Controls.Add(Me.btnOpenFile)
        Me.Name = "Form1"
        Me.Text = "CommentMaster311"
        Me.TabControl1.ResumeLayout(False)
        Me.tabMainForm.ResumeLayout(False)
        Me.tabMainForm.PerformLayout()
        Me.tabGlobalStructs.ResumeLayout(False)
        Me.tabMethods.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()
        Me.MaximizeBox = False
        Me.MaximumSize = New System.Drawing.Size(2500, 1500)
        Me.MinimizeBox = True
        Me.MinimumSize = New System.Drawing.Size(2500, 1500)

    End Sub

    Friend WithEvents btnOpenFile As Button
    Friend WithEvents TabControl1 As TabControl
    Friend WithEvents tabMainForm As TabPage
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents txtProjectName As TextBox
    Friend WithEvents lblProjectName As Label
    Friend WithEvents txtFormName As TextBox
    Friend WithEvents lblFormName As Label
    Friend WithEvents tabConstants As TabPage
    Friend WithEvents tabControlConstants As TabControl
    Friend WithEvents tabGlobalStructs As TabPage
    Friend WithEvents tabControlStructs As TabControl
    Friend WithEvents tabGlobalVariables As TabPage
    Friend WithEvents tabControlGlobalVars As TabControl
    Friend WithEvents tabMethods As TabPage
    Friend WithEvents tabControlMethods As TabControl
    Friend WithEvents OpenFileDialog1 As OpenFileDialog
    Friend WithEvents btnParseFile As Button
    Friend WithEvents btnSaveFile As Button
    Friend WithEvents lblCurrentSelectedFile As Label
    Friend WithEvents lstConsole As ListBox
    Friend WithEvents txtAuthorDate As TextBox
    Friend WithEvents txtAuthorName As TextBox
    Friend WithEvents txtFIlePurpose As TextBox
    Friend WithEvents Label4 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents txtProgramPurpose As TextBox
End Class
