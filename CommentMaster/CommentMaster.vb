Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Xml

'Comment Master by Brendan Fuller - importprogram.me - @ImportProgram
'Usage on readme

Class GlobalConstant
    Public line As Integer
    Public name As String

    Public blockDesc As String
    Public inlineDesc As String

    Public tab As TabPage
    Public txt As TextBox
    Public txtInline As TextBox
    Public Sub New(ByVal n As String, ByVal d As String, ByVal l As Integer)
        name = n
        blockDesc = d
        line = l
        tab = New TabPage()
        tab.Location = New System.Drawing.Point(8, 46)
        tab.Name = "tabConstants_" + n
        tab.Size = New System.Drawing.Size(2220, 636)
        tab.TabIndex = 0
        tab.Text = n
        tab.UseVisualStyleBackColor = True

        Dim lbl As New Label()

        lbl.AutoSize = True
        lbl.Location = New System.Drawing.Point(30, 50)
        lbl.Name = "lblConstants_" + n
        lbl.Size = New System.Drawing.Size(175, 32)
        lbl.TabIndex = 0
        lbl.Text = "Description For " + n + " (block)"

        txt = New TextBox()
        txt.Location = New System.Drawing.Point(50, 100)
        txt.Multiline = False
        txt.Name = "txtConstant_" + name
        txt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        txt.Size = New System.Drawing.Size(1007, 272)
        txt.TabIndex = 0
        txt.Enabled = False

        tab.Controls.Add(txt)
        tab.Controls.Add(lbl)

        lbl = New Label()
        lbl.AutoSize = True
        lbl.Location = New System.Drawing.Point(30, 150)
        lbl.Name = "lblConstants_" + n
        lbl.Size = New System.Drawing.Size(175, 32)
        lbl.TabIndex = 0
        lbl.Text = "Description for " + n + " (inline)"

        txtInline = New TextBox()
        txtInline.Location = New System.Drawing.Point(50, 200)
        txtInline.Multiline = False
        txtInline.Name = "txtInlineConstant_" + name
        txtInline.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        txtInline.Size = New System.Drawing.Size(1007, 272)
        txtInline.TabIndex = 0

        tab.Controls.Add(txtInline)
        tab.Controls.Add(lbl)
    End Sub

    Sub updateInterface()
        txtInline.Text = inlineDesc
    End Sub

End Class

Class GlobalDimension
    Public line As Integer
    Public name As String

    Public inlineDesc As String
    Public blockDesc As String

    Public tab As TabPage
    Public blockTxt As TextBox
    Public txt As TextBox

    Public Sub New(ByVal n As String, ByVal l As Integer)
        name = n
        line = l

        tab = New TabPage()
        tab.Location = New System.Drawing.Point(8, 46)
        tab.Name = "tabGlobalVars" + n
        tab.Size = New System.Drawing.Size(2220, 636)
        tab.TabIndex = 0
        tab.Text = n
        tab.UseVisualStyleBackColor = True

        Dim mlbl As New Label()

        mlbl.AutoSize = True
        mlbl.Location = New System.Drawing.Point(30, 50)
        mlbl.Name = "lblMinifiedGlobalVars_" + n
        mlbl.Size = New System.Drawing.Size(175, 32)
        mlbl.TabIndex = 0
        mlbl.Text = "Description for " + n + " (block)"

        Dim lbl As New Label()

        lbl.AutoSize = True
        lbl.Location = New System.Drawing.Point(30, 150)
        lbl.Name = "lblGlobalVars_" + n
        lbl.Size = New System.Drawing.Size(175, 32)
        lbl.TabIndex = 0
        lbl.Text = "Description for " + n + " (inline)"

        blockTxt = New TextBox()
        'minifiedTxt.Font = New System.Drawing.Font("Cascadia Mono", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        blockTxt.Location = New System.Drawing.Point(50, 100)
        blockTxt.Multiline = False
        blockTxt.Name = "txtMnfiedGlobal_" + name
        blockTxt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        blockTxt.Size = New System.Drawing.Size(1007, 272)
        blockTxt.TabIndex = 0

        txt = New TextBox()
        'txt.Font = New System.Drawing.Font("Cascadia Mono", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        txt.Location = New System.Drawing.Point(50, 200)
        txt.Multiline = False
        txt.Name = "txtGlobal_" + name
        txt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        txt.Size = New System.Drawing.Size(1007, 272)
        txt.TabIndex = 1

        tab.Controls.Add(txt)
        tab.Controls.Add(blockTxt)
        tab.Controls.Add(mlbl)
        tab.Controls.Add(lbl)
    End Sub

    Sub updateInterface()
        blockTxt.Text = blockDesc
        txt.Text = inlineDesc
    End Sub
End Class

Class Dimension
    Public line As Integer
    Public name As String
    Public desc As String

    Public txt As New TextBox
    Public gTxt As New TextBox
    Public lbl As New Label
    Public Sub New(ByVal n As String, ByVal l As Integer)
        name = n

        If name = "sender" Then
            desc = "Identifies which particular control that raised the click event"
        ElseIf name = "e" Then
            desc = "Holds the EventArgs object sent to the routine"
        End If

        line = l
    End Sub
End Class

Class GlobalStructure
    Public line As Integer
    Public name As String
    Public desc As String
    Public dimensions As New Dictionary(Of String, Dimension)

    Public tab As TabPage
    Public txt As TextBox

    Public Sub New(ByVal n As String, ByVal l As Integer)
        name = n
        line = l

        tab = New TabPage()
        tab.Location = New System.Drawing.Point(8, 46)
        tab.Name = "tabStruct_" + n
        tab.Size = New System.Drawing.Size(2220, 636)
        tab.TabIndex = 0
        tab.Text = n
        tab.UseVisualStyleBackColor = True
        tab.AutoScroll = True

    End Sub

    Public Sub updateInterface()
        Dim lbl As New Label()

        lbl.AutoSize = True
        lbl.Location = New System.Drawing.Point(50, 25)
        lbl.Name = "lblStructure_" + name
        lbl.Size = New System.Drawing.Size(175, 32)
        lbl.TabIndex = 0
        lbl.Text = "Description For " + name

        txt = New TextBox()
        'txt.Font = New System.Drawing.Font("Cascadia Mono", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)

        txt.Location = New System.Drawing.Point(500, 25)
        txt.Multiline = False
        txt.Name = "txtStructure_" + name
        txt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        txt.Size = New System.Drawing.Size(1000, 272)
        txt.TabIndex = 0
        txt.Text = desc

        tab.Controls.Add(txt)
        tab.Controls.Add(lbl)
    End Sub

    Public Sub addDimension(n As String, l As Integer)
        If dimensions.ContainsKey(n) = False Then
            dimensions.Add(n, New Dimension(n, l))
        End If
    End Sub

    Public Sub removeDimension(ByVal dimension As String)
        dimension.Remove(dimension)
    End Sub

End Class

Class GlobalMethods

    Public name As String

    Public writtenOn As String = ""
    Public writtenBy As String = ""

    Public purpose As String

    Public returnDescription As String
    Public returnType As String

    Public parameters As New Dictionary(Of String, Dimension)

    Public line As Integer
    Public dimensions As New Dictionary(Of String, Dimension)
    Public hasReturn As Boolean = False
    Public returnLine As Integer

    Public tab As TabPage

    Public featureTabs As TabControl
    Public tabFeatureMain As TabPage
    Public tabFeatureVars As TabPage
    Public tabFeatureParameters As TabPage

    Public txtWrittenBy As TextBox
    Public txtWrittenOn As TextBox
    Public txtPurpose As TextBox
    Public txtReturnType As TextBox
    Public txtReturnDescription As TextBox

    Public Sub New(ByVal n As String, ByVal l As Integer)
        name = n
        line = l

        tab = New TabPage()
        tab.Location = New System.Drawing.Point(8, 46)
        tab.Name = "tabMethod_" + n
        tab.Size = New System.Drawing.Size(2220, 636)
        tab.TabIndex = 0
        Dim method_name As String = name
        If method_name.Count > 20 Then
            method_name = method_name.Substring(0, 20) + "..."
        End If
        tab.Text = method_name
        tab.UseVisualStyleBackColor = True

        Dim lbl As New Label()

        lbl.AutoSize = True
        lbl.Location = New System.Drawing.Point(10, 10)
        lbl.Font = New Font("Comic Sans MC", 14, FontStyle.Regular)
        lbl.Name = "lblMethod_" + n
        lbl.Size = New System.Drawing.Size(175, 32)
        lbl.TabIndex = 0
        lbl.Text = n

        tab.Controls.Add(lbl)

        tabFeatureMain = New TabPage()
        tabFeatureMain.Location = New System.Drawing.Point(8, 46)
        tabFeatureMain.Name = "tabMethodDescriptions" + n
        tabFeatureMain.Size = New System.Drawing.Size(2220, 636)
        tabFeatureMain.TabIndex = 0
        tabFeatureMain.Text = "Method Descriptions"
        tabFeatureMain.UseVisualStyleBackColor = True

        tabFeatureVars = New TabPage()
        tabFeatureVars.Location = New System.Drawing.Point(8, 46)
        tabFeatureVars.Name = "tabMethod_" + n
        tabFeatureVars.Size = New System.Drawing.Size(2220, 636)
        tabFeatureVars.TabIndex = 0
        tabFeatureVars.Text = "Constant/Variables"
        tabFeatureVars.UseVisualStyleBackColor = True
        tabFeatureVars.AutoScroll = True

        tabFeatureParameters = New TabPage()
        tabFeatureParameters.Location = New System.Drawing.Point(8, 46)
        tabFeatureParameters.Name = "tabMethod_" + n
        tabFeatureParameters.Size = New System.Drawing.Size(2220, 636)
        tabFeatureParameters.TabIndex = 0
        tabFeatureParameters.Text = "Parameters/Arguments"
        tabFeatureParameters.UseVisualStyleBackColor = True
        tabFeatureParameters.AutoScroll = True

        featureTabs = New TabControl()
        featureTabs.Location = New System.Drawing.Point(12, 70)
        featureTabs.Name = "tcMethod_" + n
        featureTabs.SelectedIndex = 0
        featureTabs.Size = New System.Drawing.Size(2300, 750)
        featureTabs.TabIndex = 1

        featureTabs.Controls.Add(tabFeatureMain)
        featureTabs.Controls.Add(tabFeatureParameters)
        featureTabs.Controls.Add(tabFeatureVars)

        tab.Controls.Add(featureTabs)

    End Sub

    Public Sub updateInterface()

        Dim lbl As New Label

        lbl.AutoSize = True
        lbl.Location = New System.Drawing.Point(30, 10)
        lbl.Name = "lblMethod_" + name
        lbl.Size = New System.Drawing.Size(175, 32)
        lbl.TabIndex = 0
        lbl.Text = "Name"

        tabFeatureMain.Controls.Add(lbl)

        Dim txtName As New TextBox
        'txtWrittenBy.Font = New System.Drawing.Font("Cascadia Mono", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        txtName.Location = New System.Drawing.Point(300, 10)
        txtName.Enabled = False
        txtName.Name = "txtMethodName_" + name
        txtName.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        txtName.Size = New System.Drawing.Size(1000, 272)
        txtName.TabIndex = 1
        txtName.Text = name

        tabFeatureMain.Controls.Add(txtName)

        lbl = New Label
        lbl.AutoSize = True
        lbl.Location = New System.Drawing.Point(30, 100)
        lbl.Name = "lblMethod_" + name
        lbl.Size = New System.Drawing.Size(175, 32)
        lbl.TabIndex = 0
        lbl.Text = "Written By"

        tabFeatureMain.Controls.Add(lbl)

        txtWrittenBy = New TextBox()
        'txtWrittenBy.Font = New System.Drawing.Font("Cascadia Mono", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        txtWrittenBy.Location = New System.Drawing.Point(300, 100)
        txtWrittenBy.Multiline = False
        txtWrittenBy.Name = "txtMethodWrittenBy_" + name
        txtWrittenBy.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        txtWrittenBy.Size = New System.Drawing.Size(1000, 272)
        txtWrittenBy.TabIndex = 1
        txtWrittenBy.Text = writtenBy

        tabFeatureMain.Controls.Add(txtWrittenBy)

        lbl = New Label
        lbl.AutoSize = True
        lbl.Location = New System.Drawing.Point(30, 150)
        lbl.Name = "lblMethod_" + name
        lbl.Size = New System.Drawing.Size(175, 32)
        lbl.TabIndex = 0
        lbl.Text = "Written On"

        tabFeatureMain.Controls.Add(lbl)

        txtWrittenOn = New TextBox()
        'txtWrittenOn.Font = New System.Drawing.Font("Cascadia Mono", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
        txtWrittenOn.Location = New System.Drawing.Point(300, 150)
        txtWrittenOn.Multiline = False
        txtWrittenOn.Name = "txtMethodWrittenOn_" + name
        txtWrittenOn.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        txtWrittenOn.Size = New System.Drawing.Size(1000, 272)
        txtWrittenOn.TabIndex = 1
        txtWrittenOn.Text = writtenOn

        tabFeatureMain.Controls.Add(txtWrittenOn)

        If hasReturn = True Then
            lbl = New Label
            lbl.AutoSize = True
            lbl.Location = New System.Drawing.Point(30, 200)
            lbl.Name = "lblMethod_" + name
            lbl.Size = New System.Drawing.Size(175, 32)
            lbl.TabIndex = 0
            lbl.Text = "Return Type"

            tabFeatureMain.Controls.Add(lbl)

            txtReturnType = New TextBox()
            txtReturnType.Font = New System.Drawing.Font("Cascadia Mono", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            txtReturnType.Location = New System.Drawing.Point(300, 200)
            txtReturnType.Multiline = False
            txtReturnType.Name = "txtMethodReturnType_" + name
            txtReturnType.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
            txtReturnType.Size = New System.Drawing.Size(1000, 272)
            txtReturnType.TabIndex = 1
            txtReturnType.Enabled = False
            txtReturnType.Text = returnType

            tabFeatureMain.Controls.Add(txtReturnType)

            lbl = New Label
            lbl.AutoSize = True
            lbl.Location = New System.Drawing.Point(30, 250)
            lbl.Name = "lblMethod_" + name
            lbl.Size = New System.Drawing.Size(175, 32)
            lbl.TabIndex = 0
            lbl.Text = "Return Description"

            tabFeatureMain.Controls.Add(lbl)

            txtReturnDescription = New TextBox()
            'txtReturnDescription.Font = New System.Drawing.Font("Cascadia Mono", 10.125!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point)
            txtReturnDescription.Location = New System.Drawing.Point(300, 250)
            txtReturnDescription.Multiline = False
            txtReturnDescription.Name = "txtMethodReturnDescription_" + name
            txtReturnDescription.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
            txtReturnDescription.Size = New System.Drawing.Size(1000, 272)
            txtReturnDescription.TabIndex = 1
            txtReturnDescription.Text = returnDescription

            tabFeatureMain.Controls.Add(txtReturnDescription)

        End If

        lbl = New Label
        lbl.AutoSize = True
        lbl.Location = New System.Drawing.Point(1310, 10)
        lbl.Name = "lblMethodPurpose_" + name
        lbl.Size = New System.Drawing.Size(175, 32)
        lbl.TabIndex = 0
        lbl.Text = "Subprogram Purpose"

        tabFeatureMain.Controls.Add(lbl)

        txtPurpose = New TextBox
        txtPurpose.Location = New System.Drawing.Point(1310, 50)
        txtPurpose.Multiline = True
        txtPurpose.Name = "txtMethodPurpose_" + name
        txtPurpose.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        txtPurpose.Size = New System.Drawing.Size(950, 320)
        txtPurpose.TabIndex = 9

        tabFeatureMain.Controls.Add(txtPurpose)

    End Sub

    Public Sub addDimension(n As String, l As Integer)
        If dimensions.ContainsKey(n) = False Then
            dimensions.Add(n, New Dimension(n, l))
        End If
    End Sub

    Public Sub removeDimension(ByVal dimension As String)
        dimension.Remove(dimension)
    End Sub

    Public Sub setParameter(name As String)
        If parameters.ContainsKey(name) = False Then
            parameters.Add(name, New Dimension(name, -1))
        Else
            parameters(name) = New Dimension(name, -1)
        End If
    End Sub

    Public Function getParameter(name As String) As Dimension
        If parameters.ContainsKey(name) = True Then
            Return parameters(name)
        End If
    End Function

End Class

Class IndexedForm
    Public strFileName As String = ""
    Public projectName As String = ""

    Public strWrittenBy As String = ""
    Public writtenOn As String = ""

    Public filePurpose As String = ""
    Public programPurpose As String = ""

    Public commentGlobals As Dictionary(Of String, String)

    Public line As Integer

    'Make a new list of global constants
    Public constants As New Dictionary(Of String, GlobalConstant)

    'Make a new list of global structs
    Public structures As New Dictionary(Of String, GlobalStructure)

    'Make a new list of global vars
    Public globals As New Dictionary(Of String, GlobalDimension)

    Public methods As New Dictionary(Of String, GlobalMethods)

    Sub addGlobals(name As String, line As Integer)
        If globals.ContainsKey(name) = False Then
            globals.Add(name, New GlobalDimension(name, line))
        End If
    End Sub

    Public Function getGlobal(name As String) As GlobalDimension
        Return globals.GetValueOrDefault(name, Nothing)
    End Function

    Sub addConstant(name As String, line As Integer)
        If name.Count() > 3 Then
            constants.Add(name, New GlobalConstant(name, "", line))
        End If
    End Sub

    Public Function getConstant(name As String) As GlobalConstant
        Return constants.GetValueOrDefault(name, Nothing)
    End Function

    Public Sub addStructure(name As String, line As Integer)
        structures.Add(name, New GlobalStructure(name, line))
    End Sub

    Public Function getStructure(name As String) As GlobalStructure
        Return structures.GetValueOrDefault(name, Nothing)
    End Function

    Public Sub addMethod(name As String, line As Integer)
        If methods.ContainsKey(name) = False Then
            methods.Add(name, New GlobalMethods(name, line))
        End If
    End Sub

    Public Function getMethod(name As String) As GlobalMethods
        Return methods.GetValueOrDefault(name, Nothing)
    End Function
End Class

Public Class CommentMaster
    '------------------------------------------------------------
    '-                File Name: CommentMaster.vb               -
    '-              Part of Project: CommentMaster              -
    '------------------------------------------------------------
    '-                Written By: Brendan Fuller                -
    '-               Written On: January, 17 2022               -
    '------------------------------------------------------------
    '- File Purpose:                                            -
    '-                                                          -
    '- This file contains the main form for the application. All-
    '- operations are performed by this application ocurr from  -
    '- this file.                                               -
    '------------------------------------------------------------
    '- Program Purpose:                                         -
    '-                                                          -
    '- The purpose of this program is to auto generate comments -
    '- for an Advanced Visual Basic course. The program takes in-
    '- a vb file and parses the code. Said code is placed into a-
    '- class base structure to perfrom load all of the needed   -
    '- user interface components. The program also saves the    -
    '- comments in a XML file based on the selected file. All   -
    '- comments are then finally generate back into the file    -
    '- orignally made. Backups are recommend during operation.  -
    '------------------------------------------------------------
    '- Global Variables Dictionary (alphabetically)             -
    '- addedLines -                                             -
    '- copyFileText -                                           -
    '- lstStrFileText -                                         -
    '- path - Directory path of the file that was selected, used-
    '-        for XML path generation                           -
    '- strProgramFilePath -                                     -
    '- udtForm -                                                -
    '- xmlPath - The path for the XML file                      -
    '------------------------------------------------------------

    '---------------------------------------------------------------------------------------
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---
    '---------------------------------------------------------------------------------------

    Const BLOCK_ROW = "'------------------------------------------------------------" ' The
    '-                                                                                  characters
    '-                                                                                  that make
    '-                                                                                  up a
    '-                                                                                  comment
    '-                                                                                  block row.
    '-                                                                                  It's also
    '-                                                                                  used for
    '-                                                                                  the length
    '-                                                                                  of
    '-                                                                                  comments.
    Const SOMETHING = 1 ' Just a test constant. Use for debugging.
    Const MAX_COLUMN_WIDTH = 115
    '-------------------------------------------------------------------------------------------
    '--- GLOBAL STRUCTURES --- GLOBAL STRUCTURES --- GLOBAL STRUCTURES --- GLOBAL STRUCTURES ---
    '--- GLOBAL STRUCTURES --- GLOBAL STRUCTURES --- GLOBAL STRUCTURES --- GLOBAL STRUCTURES ---
    '--- GLOBAL STRUCTURES --- GLOBAL STRUCTURES --- GLOBAL STRUCTURES --- GLOBAL STRUCTURES ---
    '-------------------------------------------------------------------------------------------

    '-  The example structure provided in class. It's used for a billing data
    Structure udtBillInfo
        Dim strCompanyName As String ' What
        Dim sngGoodsValue As Single ' does
        Dim sngServicesValue As Single ' the
        Dim sngSubTotal As Single ' fox
        Dim sngSalesTax As Single ' say
        Dim sngGrandTotal As Single ' ?
    End Structure

    '-------------------------------------------------------------------------------------------
    '--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---
    '--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---
    '--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---
    '-------------------------------------------------------------------------------------------

    Dim lstStrFileText As New List(Of String) ' 
    Dim strProgramFilePath As String ' 
    Dim path As String ' Directory path of the file that was selected.
    Dim xmlPath As String ' The path for the XML file, which saves the configuration
    Dim udtForm As IndexedForm ' 
    Dim addedLines As Integer ' 
    Dim copyFileText As New ArrayList ' 

    '-------------------------------------------------------------------
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---
    '-------------------------------------------------------------------

    Private Sub btnOpenFile_Click(sender As Object, e As EventArgs) Handles btnOpenFile.Click

        '------------------------------------------------------------
        '-               Subprogram Name: btnOpenFile               -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender - Identifies which particular control that raised -
        '-          the click Event                                 -
        '- e - Holds the EventArgs Object sent To the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- result -                                                 -
        '------------------------------------------------------------

        ' Call ShowDialog.
        OpenFileDialog1.FileName = ""
        Dim result As DialogResult = OpenFileDialog1.ShowDialog()

        ' Test result.
        If result = Windows.Forms.DialogResult.OK Then

            ' Get the file name.
            strProgramFilePath = OpenFileDialog1.FileName
            reload()
        End If
    End Sub

    Private Sub btnGenerate_Click(sender As Object, e As EventArgs) Handles btnSaveFile.Click

        '------------------------------------------------------------
        '-               Subprogram Name: btnGenerate               -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender - Identifies which particular control that raised -
        '-          the click Event                                 -
        '- e - Holds the EventArgs Object sent To the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        saveXML()
        log("[CM] Saved commented to XML at: " + xmlPath)
        generateComments()
        log("[CM] Generated Comments...")
    End Sub

    Private Sub btnReload_CLick(sender As Object, e As EventArgs) Handles btnParseFile.Click

        '------------------------------------------------------------
        '-             Subprogram Name: btnReload_CLick             -
        '------------------------------------------------------------
        '-                Written By: Brendan Fuller                -
        '-                 Written On: Jan 17, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- sender - Identifies which particular control that raised -
        '-          the click Event                                 -
        '- e - Holds the EventArgs Object sent To the routine       -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        reload()
        log("[CM] Reloading")
    End Sub

    Private Sub reload()

        '------------------------------------------------------------
        '-                  Subprogram Name: reload                 -
        '------------------------------------------------------------
        '-                Written By: Brendan Fuller                -
        '-                 Written On: Jan 17, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- fileName -                                               -
        '------------------------------------------------------------

        path = System.IO.Path.GetDirectoryName(strProgramFilePath)
        Dim fileName As String = System.IO.Path.GetFileName(strProgramFilePath).Replace(".", "_")
        xmlPath = System.IO.Path.Combine(path, fileName + ".comments.xml")
        setSelectedFile()
        loadFile()
        parseCode()
        loadXML()
        updateInterface()
    End Sub

    Private Sub setSelectedFile()

        '------------------------------------------------------------
        '-             Subprogram Name: setSelectedFile             -
        '------------------------------------------------------------
        '-                Written By: Brendan Fuller                -
        '-                 Written On: Jan 17, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        lblCurrentSelectedFile.Text = strProgramFilePath
    End Sub
    Private Sub loadFile()

        '------------------------------------------------------------
        '-                 Subprogram Name: loadFile                -
        '------------------------------------------------------------
        '-                Written By: Brendan Fuller                -
        '-                 Written On: Jan 17, 2022                 -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- strLine -                                                -
        '------------------------------------------------------------

        Try
            ' Read in text.
            If System.IO.File.Exists(strProgramFilePath) Then
                lstStrFileText.Clear()
                Using streamReader As New StreamReader(strProgramFilePath)
                    While Not streamReader.EndOfStream
                        Dim strLine As String = streamReader.ReadLine()
                        If strLine.Trim().StartsWith("'-") = False Then
                            'Now add the line to the fileText
                            lstStrFileText.Add(strLine + vbCrLf)
                        End If
                    End While
                End Using
            End If
        Catch ex As Exception

            ' Report an error.
            MessageBox.Show(ex.Message)

        End Try
    End Sub
    Private Sub parseCode()

        '------------------------------------------------------------
        '-                Subprogram Name: parseCode                -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- blnInMethod -                                            -
        '- blnInStructure -                                         -
        '- gMInstance -                                             -
        '- hasParameters = True -                                   -
        '- hasReturn -                                              -
        '- inFormClass -                                            -
        '- intCurrentLine -                                         -
        '- name -                                                   -
        '- returnType -                                             -
        '- str -                                                    -
        '- strConstant -                                            -
        '- strCurrentMethod -                                       -
        '- strCurrentStructure -                                    -
        '- temp -                                                   -
        '- type -                                                   -
        '- var_name -                                               -
        '- variables -                                              -
        '- vars -                                                   -
        '------------------------------------------------------------

        Dim intCurrentLine As Integer = 1

        Dim inFormClass As Boolean = False
        Dim blnInMethod As Boolean = False
        Dim blnInStructure As Boolean = False

        Dim strCurrentStructure As String = ""
        Dim strCurrentMethod As String = ""

        udtForm = New IndexedForm

        'Set some defaults
        udtForm.strFileName = System.IO.Path.GetFileName(strProgramFilePath)
        udtForm.strWrittenBy = txtAuthorName.Text

        For Each strLine As String In lstStrFileText
            'MessageBox.Show(s)
            If strLine.StartsWith("'-") = False Then
                'Check if we are in the form block of code
                If inFormClass = True Then
                    'Are we in a structure block of code?
                    If blnInStructure = True Then
                        'Check if there is a variable within the structure.
                        If strLine.Contains("Dim" + Space(1)) = True Then
                            Dim strDimension As String = Split(Split(strLine, "Dim")(1), "As")(0).Trim()
                            udtForm.getStructure(strCurrentStructure).addDimension(strDimension, intCurrentLine)
                            log("   - " + strDimension)
                        End If
                        'Check if the line contains the ending to a structure, if so no longer within a structure
                        If strLine.Contains("End" + Space(1) + "Structure") Then
                            strCurrentStructure = ""
                            blnInStructure = False
                        End If
                    End If

                    'Are we in a method?
                    If blnInMethod = True Then
                        'Check for local variables in a method, and makes sure to not account for ReDim's
                        If strLine.Contains("Dim" + Space(1)) = True And strLine.Contains("ReDim") = False Then
                            Dim strDimension As String = Split(Split(strLine, "Dim")(1), "As")(0).Trim()
                            udtForm.getMethod(strCurrentMethod).addDimension(strDimension, intCurrentLine)
                            log("    ~ Dim: " + strDimension)
                        End If
                        'Checks for consts, and adds them to the list of local variables
                        If strLine.Contains("Const" + Space(1)) = True Then
                            Dim strConstant As String = Split(Split(strLine, "Const")(1), "=")(0).Trim()
                            udtForm.getMethod(strCurrentMethod).addDimension(strConstant, intCurrentLine)
                            log("    ~ Const: " + strConstant)
                        End If
                        'Check if are the end of the method, which is either e function or a subroutine.
                        If strLine.Contains("End" + Space(1) + "Function") Or strLine.Contains("End" + Space(1) + "Sub") Then
                            strCurrentMethod = ""
                            blnInMethod = False
                        End If

                    End If

                    'Are we NOT in a struct and NOT in a method (function/sub)?
                    If blnInStructure = False And blnInMethod = False Then

                        'If the current line is a var, add it global vars
                        If strLine.Contains("Dim" + Space(1)) = True Then
                            Dim strDimension As String = Split(Split(strLine, "Dim")(1), "As")(0).Trim()
                            udtForm.addGlobals(strDimension, intCurrentLine)
                            lstConsole.Items.Add($"[CM] Adding Global: {strDimension}")
                        End If

                        'If the current line is a constant, add it global constant
                        If strLine.Contains("Const" + Space(1)) = True Then
                            Dim strConstant As String = Split(Split(strLine, "Const")(1), "=")(0).Trim()
                            udtForm.addConstant(strConstant, intCurrentLine)
                            log($"[CM] Adding Constant: {strConstant}")
                        End If

                    End If
                    'Check for the beggining of a structure, then change the boolean values and make a new instance of the struct within IndexedForm
                    If strLine.Contains(Space(1) + "Structure" + Space(1)) = True And strLine.Contains(Space(1) + "End") = False Then
                        Dim name As String = Split(strLine, "Structure" + Space(1))(1).Trim()
                        blnInStructure = True
                        strCurrentStructure = name
                        udtForm.addStructure(strCurrentStructure, intCurrentLine)
                        log($"[CM] Adding Struct: {name} with:")

                    End If

                    'Check if we have a method (function or subroutine) and grab the parameters and return type
                    If blnInStructure = False AndAlso (strLine.Contains("Sub" + Space(1)) = True Or strLine.Contains("Function" + Space(1))) Then
                        Dim temp As String()
                        Dim type As String = "?"
                        Dim hasReturn As Boolean = False
                        Dim returnType As String
                        'Check if we have a sub or a function
                        If strLine.Contains("Sub" + Space(1)) = True Then
                            temp = Split(Split(strLine, "Sub")(1), "(")
                            type = "Sub"
                        Else
                            'If we have a function, its likely to have a return type
                            temp = Split(Split(strLine, "Function")(1), "(")
                            type = "Function"
                            hasReturn = True
                            returnType = Split(temp(1), ")", 2)(1).Trim()

                        End If

                        'Get the name from the current looped line
                        Dim name As String = temp(0).Trim().Replace("_Click", "")
                        'Now create the method, with reference to the line in the List
                        udtForm.addMethod(name, intCurrentLine)

                        'Now get the variables (both constants and dims)
                        Dim vars As String = Split(temp(1), ")")(0)

                        'Temp variable for setting dimensions
                        Dim variables As String()

                        'By default we assume we have parameters
                        Dim hasParameters = True
                        'Now check if it contains a commans, that means we have two or more paramerters
                        If vars.Contains(",") = True Then
                            'Split the values and set the list of variables
                            variables = Split(vars, ",")
                            'Now if we only have one parameter, we know a As will exist
                            'Assuming that the user is using types....
                        ElseIf vars.Contains(" As ") And vars IsNot Nothing Then
                            'Just set the length of array to 0 (which is 1?)
                            ReDim variables(0)
                            'Set the value to the string of vars, it's just easy to parse this later than now
                            variables(0) = vars
                        Else
                            'If we have nothing, well change the default state of having parameters
                            hasParameters = False
                        End If

                        log($"[CM] Adding Method ({type}): {name} with parameters:")
                        'Now get a local reference to the current method we are adding
                        Dim gMInstance As GlobalMethods = udtForm.getMethod(name)

                        'Set the value of who writes the methods, if the users sets the Written On for the project 
                        'Before file loads, ALL methods will have the same written by name :D

                        'gMInstance.writtenBy = form.writtenBy

                        'Add the parameters, if we have them
                        If hasParameters = True Then
                            For Each g As String In variables
                                'Fix the variable name if it's not perfect (this is what we didn't do above)
                                Dim var_name As String = g.Split("As")(0).Trim().Replace("ByVal", "").Replace("ByRef", "")
                                gMInstance.setParameter(var_name)
                                log("    - " + var_name)
                            Next
                        End If
                        'Tell this loop we are in a method
                        blnInMethod = True
                        'Set the current name of the method
                        strCurrentMethod = name
                        'If we have a return type, just set the vaue to the instance of the method we have above
                        If hasReturn = True Then
                            gMInstance.returnType = returnType.Replace("As", "").Trim()
                            log($"    * {returnType}")
                            gMInstance.hasReturn = True
                        End If
                    End If
                End If

                If strLine.Contains("Public Class" + Space(1)) Then
                    inFormClass = True
                    udtForm.line = intCurrentLine
                End If
            End If
            intCurrentLine = intCurrentLine + 1
        Next
        log($"[CM] End Line: {intCurrentLine.ToString}")
    End Sub
    Private Sub loadXML()

        '------------------------------------------------------------
        '-                 Subprogram Name: loadXML                 -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- attribute -                                              -
        '- blnInCommentMaster -                                     -
        '- blnInConstant -                                          -
        '- blnInGlobal -                                            -
        '- blnInMethod -                                            -
        '- blnInStructure -                                         -
        '- currentConstant -                                        -
        '- currentGlobal -                                          -
        '- currentMethod -                                          -
        '- currentStructure -                                       -
        '- m -                                                      -
        '- objGlobal -                                              -
        '- objGlobalConstant -                                      -
        '- objGlobalMethods -                                       -
        '- objGlobalStructure -                                     -
        '- s -                                                      -
        '- strAttribute -                                           -
        '------------------------------------------------------------

        If System.IO.File.Exists(xmlPath) Then
            log("[CM] Loading save from XML: " + xmlPath)

            Dim blnInCommentMaster As Boolean = False
            Dim blnInConstant As Boolean = False
            Dim blnInStructure As Boolean = False
            Dim blnInMethod As Boolean = False
            Dim blnInGlobal As Boolean = False

            Dim currentConstant As String
            Dim currentStructure As String
            Dim currentMethod As String
            Dim currentGlobal As String

            Using objXMLReader As XmlReader = XmlReader.Create(xmlPath)
                While objXMLReader.Read()
                    ' Check for start elements.
                    If objXMLReader.IsStartElement() Then

                        ' See if perls element or article element.
                        If objXMLReader.Name = "CommentMaster" Then
                            blnInCommentMaster = True
                        End If

                        If blnInCommentMaster = True And Not blnInStructure And Not blnInMethod And Not blnInGlobal And Not blnInConstant Then
                            If objXMLReader.Name = "ProjectName" Then
                                If objXMLReader.Read() Then
                                    udtForm.projectName = objXMLReader.Value.Trim()
                                End If
                            ElseIf objXMLReader.Name = "WrittenBy" Then
                                If objXMLReader.Read() Then
                                    udtForm.strWrittenBy = objXMLReader.Value.Trim()
                                End If
                            ElseIf objXMLReader.Name = "WrittenOn" Then
                                If objXMLReader.Read() Then
                                    udtForm.writtenOn = objXMLReader.Value.Trim()
                                End If
                            ElseIf objXMLReader.Name = "FilePurpose" Then
                                If objXMLReader.Read() Then
                                    udtForm.filePurpose = objXMLReader.Value.Trim()
                                End If
                            ElseIf objXMLReader.Name = "ProgramPurpose" Then
                                If objXMLReader.Read() Then
                                    udtForm.programPurpose = objXMLReader.Value.Trim()
                                End If
                            End If
                        End If

                        If blnInConstant = True Then
                            Dim objGlobalConstant As GlobalConstant = udtForm.getConstant(currentConstant)
                            If objGlobalConstant IsNot Nothing Then
                                'MessageBox.Show(reader.Name)
                                If objXMLReader.Name = "Inline" Then
                                    If objXMLReader.Read() Then
                                        objGlobalConstant.inlineDesc = objXMLReader.Value.Trim()
                                    End If
                                ElseIf objXMLReader.Name = "Block" Then
                                    If objXMLReader.Read() Then
                                        objGlobalConstant.blockDesc = objXMLReader.Value.Trim()
                                    End If
                                End If
                            End If
                        End If

                        If blnInGlobal = True Then
                            Dim objGlobalDimension As GlobalDimension = udtForm.getGlobal(currentGlobal)
                            If objGlobalDimension IsNot Nothing Then
                                If objXMLReader.Name = "Inline" Then
                                    If objXMLReader.Read() Then
                                        objGlobalDimension.inlineDesc = objXMLReader.Value.Trim()
                                    End If
                                ElseIf objXMLReader.Name = "Block" Then
                                    If objXMLReader.Read() Then
                                        objGlobalDimension.blockDesc = objXMLReader.Value.Trim()
                                    End If
                                End If
                            End If
                        End If

                        If blnInStructure = True Then
                            Dim objGlobalStructure As GlobalStructure = udtForm.getStructure(currentStructure)
                            If objGlobalStructure IsNot Nothing Then
                                If objXMLReader.Name = "Description" Then
                                    If objXMLReader.Read() Then
                                        objGlobalStructure.desc = objXMLReader.Value.Trim()
                                    End If
                                End If
                            End If
                        End If

                        If blnInMethod = True Then
                            Dim objGlobalMethods As GlobalMethods = udtForm.getMethod(currentMethod)
                            If objGlobalMethods IsNot Nothing Then
                                If objXMLReader.Name = "Purpose" Then
                                    If objXMLReader.Read() Then
                                        objGlobalMethods.purpose = objXMLReader.Value.Trim()
                                    End If
                                ElseIf objXMLReader.Name = "WrittenOn" Then
                                    If objXMLReader.Read() Then
                                        objGlobalMethods.writtenOn = objXMLReader.Value.Trim()
                                    End If
                                ElseIf objXMLReader.Name = "WrittenBy" Then
                                    If objXMLReader.Read() Then
                                        objGlobalMethods.writtenBy = objXMLReader.Value.Trim()
                                    End If
                                End If
                                If objGlobalMethods.hasReturn Then
                                    If objXMLReader.Name = "ReturnDescription" Then
                                        If objXMLReader.Read() Then
                                            objGlobalMethods.returnDescription = objXMLReader.Value.Trim()
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        If objXMLReader.Name = "Dimension" Then
                            If blnInStructure = True Then
                                Dim strAttribute As String = objXMLReader("name")
                                Dim s As GlobalStructure = udtForm.getStructure(currentStructure)
                                If strAttribute IsNot Nothing And s IsNot Nothing Then
                                    If s.dimensions.ContainsKey(strAttribute) = True Then
                                        If objXMLReader.Read() Then

                                            s.dimensions(strAttribute).desc = objXMLReader.Value.Trim()
                                        End If
                                    End If
                                End If
                            ElseIf blnInMethod = True Then
                                Dim strAttribute As String = objXMLReader("name")
                                Dim m As GlobalMethods = udtForm.getMethod(currentMethod)
                                If strAttribute IsNot Nothing And m IsNot Nothing Then
                                    If m.dimensions.ContainsKey(strAttribute) = True Then
                                        If objXMLReader.Read() Then
                                            m.dimensions(strAttribute).desc = objXMLReader.Value.Trim()
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        If objXMLReader.Name = "Parameter" Then
                            If blnInMethod = True Then
                                Dim attribute As String = objXMLReader("name")
                                Dim m As GlobalMethods = udtForm.getMethod(currentMethod)
                                If attribute IsNot Nothing And m IsNot Nothing Then
                                    If m.parameters.ContainsKey(attribute) = True Then
                                        If objXMLReader.Read() Then
                                            m.parameters(attribute).desc = objXMLReader.Value.Trim()
                                        End If
                                    End If
                                End If
                            End If
                        End If

                        If objXMLReader.Name = "Structure" Then
                            blnInMethod = False
                            blnInStructure = False
                            blnInGlobal = False
                            blnInConstant = False
                            Dim attribute As String = objXMLReader("name")
                            If attribute IsNot Nothing Then
                                blnInStructure = True
                                blnInMethod = False
                                blnInConstant = False
                                blnInGlobal = False
                                currentStructure = attribute
                            End If
                        End If

                        If objXMLReader.Name = "Method" Then
                            blnInMethod = False
                            blnInStructure = False
                            blnInGlobal = False
                            blnInConstant = False
                            Dim attribute As String = objXMLReader("name")
                            If attribute IsNot Nothing Then
                                blnInStructure = False
                                blnInMethod = True
                                blnInConstant = False
                                blnInGlobal = False
                                currentMethod = attribute
                            End If
                        End If

                        If objXMLReader.Name = "Global" Then
                            blnInMethod = False
                            blnInStructure = False
                            blnInGlobal = False
                            blnInConstant = False
                            Dim attribute As String = objXMLReader("name")
                            If attribute IsNot Nothing Then
                                blnInStructure = False
                                blnInMethod = False
                                blnInConstant = False
                                blnInGlobal = True
                                currentGlobal = attribute
                            End If
                        End If

                        If objXMLReader.Name = "Constant" Then
                            blnInMethod = False
                            blnInStructure = False
                            blnInGlobal = False
                            blnInConstant = False
                            Dim attribute As String = objXMLReader("name")
                            If attribute IsNot Nothing Then
                                blnInStructure = False
                                blnInMethod = False
                                blnInConstant = True
                                blnInGlobal = False
                                currentConstant = attribute
                            End If

                        End If
                    End If
                End While
            End Using
        Else
            log("[CM] No comments saved for this project..." + xmlPath)
        End If
    End Sub
    Private Sub updateInterface()

        '------------------------------------------------------------
        '-             Subprogram Name: updateInterface             -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- desc -                                                   -
        '- gTxt -                                                   -
        '- i -                                                      -
        '- lbl -                                                    -
        '- name -                                                   -
        '- txt -                                                    -
        '- y -                                                      -
        '------------------------------------------------------------

        txtProjectName.Text = udtForm.projectName
        txtFormName.Text = udtForm.strFileName
        txtAuthorName.Text = udtForm.strWrittenBy
        txtAuthorDate.Text = udtForm.writtenOn

        txtProgramPurpose.Text = udtForm.programPurpose
        txtFIlePurpose.Text = udtForm.filePurpose

        tabControlConstants.Controls.Clear()
        tabControlStructs.Controls.Clear()
        tabControlGlobalVars.Controls.Clear()
        tabControlMethods.Controls.Clear()

        btnSaveFile.Enabled = True
        btnParseFile.Enabled = True
        'Constants
        For Each key In udtForm.constants.Keys()
            tabControlConstants.Controls.Add(udtForm.constants(key).tab)
            udtForm.constants(key).updateInterface()
        Next

        For Each key In udtForm.structures.Keys()
            udtForm.structures(key).tab.Controls.Clear()
            tabControlStructs.Controls.Add(udtForm.structures(key).tab)
            Dim y As Integer = 110
            Dim i As Integer
            For Each var In udtForm.structures(key).dimensions.Keys()
                Dim txt As TextBox = udtForm.structures(key).dimensions(var).txt
                Dim lbl As Label = udtForm.structures(key).dimensions(var).lbl
                Dim name As String = udtForm.structures(key).dimensions(var).name
                Dim desc As String = udtForm.structures(key).dimensions(var).desc

                lbl.AutoSize = True
                lbl.Location = New System.Drawing.Point(50, y)
                lbl.Name = "lblStructDim_" + key + "_" + var
                lbl.Size = New System.Drawing.Size(175, 32)
                lbl.TabIndex = 0
                lbl.Text = name

                txt.Location = New System.Drawing.Point(500, y)
                txt.Name = "txtStructDim_" + key + "_" + var
                txt.Size = New System.Drawing.Size(1000, 272)
                txt.Text = desc

                udtForm.structures(key).tab.Controls.Add(lbl)
                udtForm.structures(key).tab.Controls.Add(txt)
                i = i + 1
                y = y + 60
            Next
            udtForm.structures(key).updateInterface()
        Next

        'Globals
        For Each key In udtForm.globals.Keys()
            tabControlGlobalVars.Controls.Add(udtForm.globals(key).tab)
            udtForm.globals(key).updateInterface()
        Next

        'Methods
        For Each key In udtForm.methods.Keys()
            tabControlMethods.Controls.Add(udtForm.methods(key).tab)
            udtForm.methods(key).updateInterface()
            Dim y As Integer = 10
            Dim i As Integer
            For Each var In udtForm.methods(key).dimensions.Keys()
                Dim txt As TextBox = udtForm.methods(key).dimensions(var).txt
                Dim lbl As Label = udtForm.methods(key).dimensions(var).lbl
                Dim gTxt As TextBox = udtForm.methods(key).dimensions(var).gTxt

                Dim desc As String = udtForm.methods(key).dimensions(var).desc

                lbl.AutoSize = True
                lbl.Location = New System.Drawing.Point(30, y)
                lbl.Name = "lblMethodFeature_" + key + "_" + var
                lbl.Size = New System.Drawing.Size(175, 32)
                lbl.TabIndex = 0
                lbl.Text = udtForm.methods(key).dimensions(var).name

                txt.Location = New System.Drawing.Point(500, y)
                txt.Name = "txtMethodFeature_" + key + "_" + var
                txt.Size = New System.Drawing.Size(1700, 272)
                txt.Text = desc

                udtForm.methods(key).tabFeatureVars.Controls.Add(lbl)
                udtForm.methods(key).tabFeatureVars.Controls.Add(txt)
                i = i + 1
                y = y + 60
            Next
            y = 10
            i = 0
            For Each var In udtForm.methods(key).parameters.Keys()
                Dim txt As TextBox = udtForm.methods(key).parameters(var).txt
                Dim lbl As Label = udtForm.methods(key).parameters(var).lbl
                Dim desc As String = udtForm.methods(key).parameters(var).desc

                lbl.AutoSize = True
                lbl.Location = New System.Drawing.Point(30, y)
                lbl.Name = "lblMethodFeature_" + key + "_" + var
                lbl.Size = New System.Drawing.Size(175, 32)
                lbl.TabIndex = 0
                lbl.Text = udtForm.methods(key).parameters(var).name

                txt.Location = New System.Drawing.Point(500, y)
                txt.Name = "txtMethodFeature_" + key + "_" + var
                txt.Size = New System.Drawing.Size(1000, 272)
                txt.Text = desc

                udtForm.methods(key).tabFeatureParameters.Controls.Add(lbl)
                udtForm.methods(key).tabFeatureParameters.Controls.Add(txt)
                i = i + 1
                y = y + 60
            Next
        Next
    End Sub
    Private Sub interfaceToIndexedForm()

        '------------------------------------------------------------
        '-          Subprogram Name: interfaceToIndexedForm         -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- method -                                                 -
        '------------------------------------------------------------

        udtForm.projectName = txtProjectName.Text

        udtForm.strWrittenBy = txtAuthorName.Text
        udtForm.writtenOn = txtAuthorDate.Text

        udtForm.filePurpose = txtFIlePurpose.Text
        udtForm.programPurpose = txtProgramPurpose.Text

        'Constants
        For Each key In udtForm.constants.Keys()
            udtForm.constants(key).blockDesc = udtForm.constants(key).txt.Text
            udtForm.constants(key).inlineDesc = udtForm.constants(key).txtInline.Text
        Next

        'Structs
        For Each key In udtForm.structures.Keys()
            udtForm.structures(key).desc = udtForm.structures(key).txt.Text
            For Each dim_key In udtForm.structures(key).dimensions.Keys()
                udtForm.structures(key).dimensions(dim_key).desc = udtForm.structures(key).dimensions(dim_key).txt.Text
            Next
        Next

        'Globals
        For Each key In udtForm.globals.Keys()
            udtForm.globals(key).inlineDesc = udtForm.globals(key).txt.Text
            udtForm.globals(key).blockDesc = udtForm.globals(key).blockTxt.Text
        Next

        'Methods
        For Each key In udtForm.methods.Keys()
            Dim method As GlobalMethods = udtForm.methods(key)

            method.writtenBy = method.txtWrittenBy.Text
            method.writtenOn = method.txtWrittenOn.Text

            If method.hasReturn = True Then
                method.returnDescription = method.txtReturnDescription.Text
            End If

            method.purpose = method.txtPurpose.Text
            For Each var In udtForm.methods(key).dimensions.Keys()
                udtForm.methods(key).dimensions(var).desc = udtForm.methods(key).dimensions(var).txt.Text
            Next

            For Each var In udtForm.methods(key).parameters.Keys()
                udtForm.methods(key).parameters(var).desc = udtForm.methods(key).parameters(var).txt.Text
            Next
        Next
    End Sub
    Private Sub saveXML()

        '------------------------------------------------------------
        '-                 Subprogram Name: saveXML                 -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- method -                                                 -
        '- writer -                                                 -
        '------------------------------------------------------------

        interfaceToIndexedForm()
        Dim writer As New XmlTextWriter(xmlPath, System.Text.Encoding.UTF8)

        writer.WriteStartDocument(True)
        writer.Formatting = Formatting.Indented
        writer.Indentation = 2
        writer.WriteStartElement("CommentMaster")

        writer.WriteAttributeString("author", "Brendan Fuller (https://importprogram.me)")
        writer.WriteAttributeString("version", "311")

        'Project Name
        writer.WriteStartElement("ProjectName")
        writer.WriteString(udtForm.projectName)
        writer.WriteEndElement()

        'Written By
        writer.WriteStartElement("WrittenBy")
        writer.WriteString(udtForm.strWrittenBy)
        writer.WriteEndElement()

        'Written On
        writer.WriteStartElement("WrittenOn")
        writer.WriteString(udtForm.writtenOn)
        writer.WriteEndElement()

        'File purpose
        writer.WriteStartElement("FilePurpose")
        writer.WriteString(udtForm.filePurpose)
        writer.WriteEndElement()

        'File purpose
        writer.WriteStartElement("ProgramPurpose")
        writer.WriteString(udtForm.programPurpose)
        writer.WriteEndElement()

        'Parse more here
        writer.WriteStartElement("Constants")
        For Each key In udtForm.constants.Keys()
            writer.WriteStartElement("Constant")
            writer.WriteAttributeString("name", key)

            writer.WriteStartElement("Inline")
            writer.WriteString(udtForm.constants(key).inlineDesc)
            writer.WriteEndElement()

            writer.WriteStartElement("Block")
            writer.WriteString(udtForm.constants(key).blockDesc)
            writer.WriteEndElement()

            writer.WriteEndElement()
        Next
        writer.WriteEndElement()

        writer.WriteStartElement("Structures")
        For Each key In udtForm.structures.Keys()
            writer.WriteStartElement("Structure")
            writer.WriteAttributeString("name", key)

            writer.WriteStartElement("Description")
            writer.WriteString(udtForm.structures(key).desc)
            writer.WriteEndElement()

            writer.WriteStartElement("Dimensions")
            For Each dim_key In udtForm.structures(key).dimensions.Keys()
                writer.WriteStartElement("Dimension")
                writer.WriteAttributeString("name", dim_key)
                writer.WriteString(udtForm.structures(key).dimensions(dim_key).desc)
                writer.WriteEndElement()
            Next
            writer.WriteEndElement()
            writer.WriteEndElement()
        Next
        writer.WriteEndElement()

        writer.WriteStartElement("Globals")
        For Each key In udtForm.globals.Keys()
            writer.WriteStartElement("Global")
            writer.WriteAttributeString("name", key)

            writer.WriteStartElement("Inline")
            writer.WriteString(udtForm.globals(key).inlineDesc)
            writer.WriteEndElement()

            writer.WriteStartElement("Block")
            writer.WriteString(udtForm.globals(key).blockDesc)
            writer.WriteEndElement()

            writer.WriteEndElement()
        Next
        writer.WriteEndElement()

        'Methods
        writer.WriteStartElement("Methods")
        For Each key In udtForm.methods.Keys()
            Dim method As GlobalMethods = udtForm.methods(key)

            writer.WriteStartElement("Method")
            writer.WriteAttributeString("name", key)

            writer.WriteStartElement("WrittenBy")
            writer.WriteString(method.writtenBy)
            writer.WriteEndElement()

            writer.WriteStartElement("WrittenOn")
            writer.WriteString(method.writtenOn)
            writer.WriteEndElement()

            If method.hasReturn = True Then
                writer.WriteStartElement("ReturnDescription")
                writer.WriteString(method.returnDescription)
                writer.WriteEndElement()
            End If

            writer.WriteStartElement("Purpose")
            writer.WriteString(method.purpose)
            writer.WriteEndElement()

            writer.WriteStartElement("Dimensions")
            For Each var In udtForm.methods(key).dimensions.Keys()
                writer.WriteStartElement("Dimension")
                writer.WriteAttributeString("name", var)
                writer.WriteString(udtForm.methods(key).dimensions(var).desc)
                writer.WriteEndElement()
            Next
            writer.WriteEndElement()

            writer.WriteStartElement("Parameters")
            For Each var In udtForm.methods(key).parameters.Keys()
                writer.WriteStartElement("Parameter")
                writer.WriteAttributeString("name", var)
                writer.WriteString(udtForm.methods(key).parameters(var).desc)
                writer.WriteEndElement()
                'form.methods(key).parameters(var).desc = form.methods(key).parameters(var).txt.Text
            Next
            writer.WriteEndElement()

            writer.WriteEndElement()
        Next
        writer.WriteEndElement()

        writer.WriteEndElement()
        writer.WriteEndDocument()
        writer.Close()
    End Sub
    Private Function generateVariableWithDescWrapping(variable As String, desc As String) As List(Of String)

        '------------------------------------------------------------
        '-     Subprogram Name: generateVariableWithDescWrapping    -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- variable -                                               -
        '- desc -                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- currentLine -                                            -
        '- firstRow -                                               -
        '- length -                                                 -
        '- lines -                                                  -
        '- spaces -                                                 -
        '- spacesNeeded -                                           -
        '- thisLine -                                               -
        '- total -                                                  -
        '- words -                                                  -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- List -                                                   -
        '------------------------------------------------------------

        'Split a list of words
        Dim words As String() = Split(Regex.Replace(desc.Replace(vbCrLf, ""), " {2,}", " "), " ")
        'Get a total length of the comment minus 4 for the comment structure
        Dim total As Integer = BLOCK_ROW.Count - 4
        'The lines to be returned
        Dim lines As New List(Of String)
        'The current line to be generated
        Dim currentLine As String = ""
        'The line that is being added to the current lines
        Dim thisLine As String
        'Loop all words

        Dim firstRow As Boolean = True
        For Each s As String In words
            'Get the length of the word
            Dim length As Integer = s.Count
            'If the length of the word doesn't exceed the total width of the comment block length, 
            'Then go to the else and added it to the currentLine (with a space). Otherwise
            'add the all of the words to this moment to the lines array and save the current word
            'for the next line array
            If currentLine.Count + length > total Then
                'Make the new line comment
                thisLine = "'-" + currentLine
                'Calculate the spaces for the right padding
                Dim spacesNeeded As Integer = BLOCK_ROW.Count - 1 - thisLine.Count
                'Add the spaces to the end, and close the block
                thisLine = thisLine + Space(spacesNeeded) + "-"
                'Add it to the list of lines
                lines.Add(thisLine)
                'Reset for each line
                currentLine = Space(variable.Trim().Count + 4) + s
                thisLine = ""
            Else
                If firstRow = True Then
                    firstRow = False
                    currentLine = " " + variable.Trim() + " - " + s.Trim()
                Else
                    currentLine = currentLine + " " + s.Trim()
                End If
            End If
        Next
        'Now if the for loop is done but has values left, we need to add the last line in
        If currentLine.Count > 0 Then
            thisLine = "'-" + currentLine
            Dim spaces As Integer = BLOCK_ROW.Count - 1 - thisLine.Count
            If spaces < 0 Then
                spaces = 0
                thisLine = thisLine.Substring(0, BLOCK_ROW.Count - 1)
            End If
            thisLine = thisLine + Space(spaces) + "-"
            lines.Add(thisLine)
        End If
        Return lines
    End Function
    Private Function generateInlineWordWrapping(code As String, desc As String) As List(Of String)

        '------------------------------------------------------------
        '-        Subprogram Name: generateInlineWordWrapping       -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- code -                                                   -
        '- desc -                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- currentLine -                                            -
        '- firstRow -                                               -
        '- length -                                                 -
        '- lines -                                                  -
        '- thisLine -                                               -
        '- total -                                                  -
        '- words -                                                  -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- List -                                                   -
        '------------------------------------------------------------

        'Split a list of words
        Dim words As String() = Split(Regex.Replace(desc.Replace(vbCrLf, ""), " {2,}", " "), " ")
        'Get a total length of the comment minus 4 for the comment structure
        Dim total As Integer = MAX_COLUMN_WIDTH
        'The lines to be returned
        Dim lines As New List(Of String)
        'The current line to be generated
        Dim currentLine As String = ""
        'The line that is being added to the current lines
        Dim thisLine As String
        'Loop all words

        Dim firstRow As Boolean = True
        For Each s As String In words
            'Get the length of the word
            Dim length As Integer = s.Count
            'If the length of the word doesn't exceed the total width of the comment block length, 
            'Then go to the else and added it to the currentLine (with a space). Otherwise
            'add the all of the words to this moment to the lines array and save the current word
            'for the next line array
            If currentLine.Count + length > total Then
                'Add it to the list of lines
                lines.Add(currentLine)
                'Reset for each line
                currentLine = "'- " + Space(code.Trim().Count) + s

                thisLine = ""
            Else
                If firstRow = True Then
                    firstRow = False
                    currentLine = code.Trim() + Space(1) + "'" + Space(1) + s.Trim()
                Else
                    currentLine = currentLine + " " + s.Trim()
                End If
            End If
        Next
        'Now if the for loop is done but has values left, we need to add the last line in
        If currentLine.Count > 0 Then
            lines.Add(currentLine)
        End If
        Return lines
    End Function
    Private Function generateWordWrappingBlock(text As String) As List(Of String)

        '------------------------------------------------------------
        '-        Subprogram Name: generateWordWrappingBlock        -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- text -                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- currentLine -                                            -
        '- length -                                                 -
        '- lines -                                                  -
        '- spaces -                                                 -
        '- spacesNeeded -                                           -
        '- thisLine -                                               -
        '- total -                                                  -
        '- words -                                                  -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- List -                                                   -
        '------------------------------------------------------------

        'Split a list of words
        Dim words As String() = Split(Regex.Replace(text.Replace(vbCrLf, ""), " {2,}", " "), " ")
        'Get a total length of the comment minus 4 for the comment structure
        Dim total As Integer = BLOCK_ROW.Count - 4
        'The lines to be returned
        Dim lines As New List(Of String)
        'The current line to be generated
        Dim currentLine As String = ""
        'The line that is being added to the current lines
        Dim thisLine As String
        'Loop all words
        For Each s As String In words
            'Get the length of the word
            Dim length As Integer = s.Count
            'If the length of the word doesn't exceed the total width of the comment block length, 
            'Then go to the else and added it to the currentLine (with a space). Otherwise
            'add the all of the words to this moment to the lines array and save the current word
            'for the next line array
            If currentLine.Count + length > total Then
                'Make the new line comment
                thisLine = "'-" + currentLine
                'Calculate the spaces for the right padding
                Dim spacesNeeded As Integer = BLOCK_ROW.Count - 1 - thisLine.Count
                'Also if the word is to long to fit on a single line, it just get's truncated
                'This was done because I really didn't feel like hyphoning the words
                If spacesNeeded < 0 Then
                    spacesNeeded = 0
                    thisLine = thisLine.Substring(0, total + 3)
                End If
                'Add the spaces to the end, and close the block
                thisLine = thisLine + Space(spacesNeeded) + "-"
                'Add it to the list of lines
                lines.Add(thisLine)
                'Reset for each line
                currentLine = " " + s
                thisLine = ""
            Else
                currentLine = currentLine + " " + s.Trim()
            End If
        Next
        'Now if the for loop is done but has values left, we need to add the last line in
        If currentLine.Count > 0 Then
            thisLine = "'-" + currentLine
            Dim spaces As Integer = BLOCK_ROW.Count - 1 - thisLine.Count
            If spaces < 0 Then
                spaces = 0
                thisLine = thisLine.Substring(0, total + 3)
            End If
            thisLine = thisLine + Space(spaces) + "-"
            lines.Add(thisLine)
        End If
        Return lines
    End Function
    Private Function generateWordWrapping(text As String) As List(Of String)

        '------------------------------------------------------------
        '-           Subprogram Name: generateWordWrapping          -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- text -                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- currentLine -                                            -
        '- length -                                                 -
        '- lines -                                                  -
        '- thisLine -                                               -
        '- total -                                                  -
        '- words -                                                  -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- List -                                                   -
        '------------------------------------------------------------

        'Split a list of words
        Dim words As String() = Split(Regex.Replace(text.Replace(vbCrLf, ""), " {2,}", " "), " ")
        'Get a total length of the comment minus 4 for the comment structure
        Dim total As Integer = MAX_COLUMN_WIDTH
        'The lines to be returned
        Dim lines As New List(Of String)
        'The current line to be generated
        Dim currentLine As String = ""
        'The line that is being added to the current lines
        Dim thisLine As String
        'Loop all words
        For Each s As String In words
            'Get the length of the word
            Dim length As Integer = s.Count
            'If the length of the word doesn't exceed the total width of the comment block length, 
            'Then go to the else and added it to the currentLine (with a space). Otherwise
            'add the all of the words to this moment to the lines array and save the current word
            'for the next line array
            If currentLine.Count + length > total Then
                'Make the new line comment
                thisLine = "'- " + currentLine
                lines.Add(thisLine)
                'Reset for each line
                currentLine = s
                thisLine = ""
            Else
                currentLine = currentLine + " " + s.Trim()
            End If
        Next
        'Now if the for loop is done but has values left, we need to add the last line in
        If currentLine.Count > 0 Then
            thisLine = "'- " + currentLine
            lines.Add(thisLine)
        End If
        Return lines
    End Function

    Private Function getLineSpacing(line As Integer) As String

        '------------------------------------------------------------
        '-              Subprogram Name: getLineSpacing             -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- line -                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- String -                                                 -
        '------------------------------------------------------------

        Return copyFileText(line).ToString().Count - LTrim(copyFileText(line).ToString()).Count
    End Function

    Private Function PadCenterComment(line As String) As String

        '------------------------------------------------------------
        '-             Subprogram Name: PadCenterComment            -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- line -                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- left -                                                   -
        '- lengthSpaces -                                           -
        '- right -                                                  -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- String -                                                 -
        '------------------------------------------------------------

        Dim lengthSpaces As Double = (BLOCK_ROW.Count - line.Count - 4) / 2
        'MessageBox.Show(lengthSpaces)
        Dim left As String
        Dim right As String
        If Math.Truncate(lengthSpaces) = lengthSpaces Then
            right = Space(lengthSpaces)
            left = Space(lengthSpaces)
        Else
            right = Space(Math.Ceiling(lengthSpaces))
            left = Space(Math.Floor(lengthSpaces))
        End If

        Return "'- " + left + line + right + "-"
    End Function

    Private Function PadRightComment(line As String) As String

        '------------------------------------------------------------
        '-             Subprogram Name: PadRightComment             -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- line -                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- left -                                                   -
        '- lengthSpaces -                                           -
        '- right -                                                  -
        '------------------------------------------------------------
        '- Returns:                                                 -
        '- String -                                                 -
        '------------------------------------------------------------

        Dim lengthSpaces As Double = (BLOCK_ROW.Count - line.Count - 4)
        'MessageBox.Show(lengthSpaces)
        Dim left As String
        Dim right As String
        If Math.Truncate(lengthSpaces) = lengthSpaces Then
            right = Space(lengthSpaces)
            left = Space(lengthSpaces)
        Else
            right = Space(Math.Ceiling(lengthSpaces))
            left = Space(Math.Floor(lengthSpaces))
        End If

        Return "'- " + line + right + "-"
    End Function

    Private Sub insertLines(startLine As String, lines As List(Of String))

        '------------------------------------------------------------
        '-               Subprogram Name: insertLines               -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- startLine -                                              -
        '- lines -                                                  -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- number -                                                 -
        '------------------------------------------------------------

        Dim number As Integer = 0
        For Each s As String In lines
            insertLine(startLine + number, "    " + s)
            number = number + 1
        Next
    End Sub

    Private Sub insertLine(line As Integer, value As String)

        '------------------------------------------------------------
        '-                Subprogram Name: insertLine               -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- line -                                                   -
        '- value -                                                  -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        addedLines = addedLines + 1
        copyFileText.Insert(line, value + vbCrLf)
    End Sub

    Private Sub insertLinesSpaces(startLine As String, lines As List(Of String), spaces As String)

        '------------------------------------------------------------
        '-            Subprogram Name: insertLinesSpaces            -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- startLine -                                              -
        '- lines -                                                  -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- number -                                                 -
        '------------------------------------------------------------

        Dim number As Integer = 0
        For Each s As String In lines
            insertLine(startLine + number, spaces + s)
            number = number + 1
        Next
    End Sub

    Private Sub generateComments()

        '------------------------------------------------------------
        '-             Subprogram Name: generateComments            -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- (None)                                                   -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- desc -                                                   -
        '- descInline -                                             -
        '- dim_desc -                                               -
        '- dim_line -                                               -
        '- firstConstant -                                          -
        '- firstGlobal -                                            -
        '- firstMethod -                                            -
        '- firstStructure -                                         -
        '- global_keys -                                            -
        '- hasGlobal -                                              -
        '- keys -                                                   -
        '- lastLineWasEmpty -                                       -
        '- line -                                                   -
        '- method -                                                 -
        '- name -                                                   -
        '- spaces -                                                 -
        '- tempFile -                                               -
        '- tempFilePath -                                           -
        '- text -                                                   -
        '------------------------------------------------------------

        addedLines = 0
        copyFileText.Clear()

        copyFileText.AddRange(lstStrFileText)

        Dim spaces As String = "    "

        'File Name & Project
        insertLine(udtForm.line, spaces + BLOCK_ROW)
        insertLine(udtForm.line + addedLines, spaces + PadCenterComment("File Name: " + udtForm.strFileName))
        insertLine(udtForm.line + addedLines, spaces + PadCenterComment("Part of Project: " + udtForm.projectName))
        insertLine(udtForm.line + addedLines, spaces + BLOCK_ROW)

        'Written By and Written On
        insertLine(udtForm.line + addedLines, spaces + PadCenterComment("Written By: " + udtForm.strWrittenBy))
        insertLine(udtForm.line + addedLines, spaces + PadCenterComment("Written On: " + udtForm.writtenOn))
        insertLine(udtForm.line + addedLines, spaces + BLOCK_ROW)

        'File purpose
        insertLine(udtForm.line + addedLines, spaces + PadRightComment("File Purpose:"))
        insertLine(udtForm.line + addedLines, spaces + PadRightComment(""))
        insertLines(udtForm.line + addedLines, generateWordWrappingBlock(udtForm.filePurpose))
        insertLine(udtForm.line + addedLines, spaces + BLOCK_ROW)

        'Program Purpose
        insertLine(udtForm.line + addedLines, spaces + PadRightComment("Program Purpose:"))
        insertLine(udtForm.line + addedLines, spaces + PadRightComment(""))
        insertLines(udtForm.line + addedLines, generateWordWrappingBlock(udtForm.programPurpose))
        insertLine(udtForm.line + addedLines, spaces + BLOCK_ROW)

        'Global Vars
        insertLine(udtForm.line + addedLines, spaces + PadRightComment("Global Variables Dictionary (alphabetically)"))

        'Sort the global keys alphabetically
        Dim global_keys As List(Of String) = udtForm.globals.Keys.ToList
        global_keys.Sort()

        Dim hasGlobal As Boolean = False
        For Each key In global_keys
            Dim desc As String = udtForm.globals(key).blockDesc
            Dim name As String = udtForm.globals(key).name

            insertLines(udtForm.line + addedLines, generateVariableWithDescWrapping(name, desc))
            hasGlobal = True
        Next
        'Check if we have any globals to comment anyways
        If hasGlobal = False Then
            insertLine(udtForm.line + addedLines, spaces + PadRightComment("(None)"))
        End If
        insertLine(udtForm.line + addedLines, spaces + BLOCK_ROW)

        'CONSTANTS
        'While we loop, we need to append the bannar for GLOBAL CONSTANTS, so first iteration will be used
        Dim firstConstant As Boolean = True
        'Iterate over all of the constants

        For Each key In udtForm.constants.Keys()

            'Grab the values from each of the classes of GlobalConstants and make local referencnes
            Dim descInline As String = udtForm.constants(key).inlineDesc
            Dim name As String = udtForm.constants(key).name
            Dim line As Integer = udtForm.constants(key).line
            'Check if we have the first constant being looped
            If firstConstant = True Then
                firstConstant = False

                insertLine(line - 1 + addedLines, "")
                insertLine(line - 1 + addedLines, spaces + "'---------------------------------------------------------------------------------------")
                insertLine(line - 1 + addedLines, spaces + "'--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---")
                insertLine(line - 1 + addedLines, spaces + "'--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---")
                insertLine(line - 1 + addedLines, spaces + "'--- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS --- GLOBAL CONSTANTS ---")
                insertLine(line - 1 + addedLines, spaces + "'---------------------------------------------------------------------------------------")
                insertLine(line - 1 + addedLines, "") 'Add an empty line

            End If

            Dim text As String = copyFileText(line + addedLines - 1)
            If text.Contains(Space(1) + "'" + Space(1)) Then
                text = text.Split(Space(1) + "'" + Space(1), 2)(0)
            End If
            copyFileText.RemoveAt(line + addedLines - 1)
            'Also remove a line from the list of total added
            addedLines = addedLines - 1
            insertLinesSpaces(line + addedLines, generateInlineWordWrapping(text, descInline), spaces)

        Next

        'STRUCTURES
        Dim firstStructure As Boolean = True
        'MessageBox.Show(form.constants.Keys().Count.ToString)
        For Each key In udtForm.structures.Keys()
            Dim desc As String = udtForm.structures(key).desc
            Dim line As Integer = udtForm.structures(key).line

            If firstStructure = True Then
                firstStructure = False

                insertLine(line - 1 + addedLines, "") 'Add an empty line
                insertLine(line - 1 + addedLines, spaces + "'-------------------------------------------------------------------------------------------")
                insertLine(line - 1 + addedLines, spaces + "'--- GLOBAL STRUCTURES --- GLOBAL STRUCTURES --- GLOBAL STRUCTURES --- GLOBAL STRUCTURES ---")
                insertLine(line - 1 + addedLines, spaces + "'--- GLOBAL STRUCTURES --- GLOBAL STRUCTURES --- GLOBAL STRUCTURES --- GLOBAL STRUCTURES ---")
                insertLine(line - 1 + addedLines, spaces + "'--- GLOBAL STRUCTURES --- GLOBAL STRUCTURES --- GLOBAL STRUCTURES --- GLOBAL STRUCTURES ---")
                insertLine(line - 1 + addedLines, spaces + "'-------------------------------------------------------------------------------------------")
                insertLine(line - 1 + addedLines, "") 'Add an empty line

            End If

            insertLines(line - 1 + addedLines, generateWordWrapping(desc))

            For Each dim_key In udtForm.structures(key).dimensions.Keys()
                Dim dim_desc As String = udtForm.structures(key).dimensions(dim_key).desc
                Dim dim_line As String = udtForm.structures(key).dimensions(dim_key).line

                Dim text As String = copyFileText(dim_line + addedLines - 1)
                If text.Contains(Space(1) + "'" + Space(1)) Then
                    text = text.Split(Space(1) + "'" + Space(1), 2)(0)
                End If
                copyFileText.RemoveAt(dim_line + addedLines - 1)
                'Also remove a line from the list of total added
                addedLines = addedLines - 1
                insertLinesSpaces(dim_line + addedLines, generateInlineWordWrapping(text, dim_desc), Space(8))

            Next
        Next

        'GLOBALS (inline)
        Dim firstGlobal As Boolean = True
        For Each key In udtForm.globals.Keys()
            Dim desc As String = udtForm.globals(key).inlineDesc
            Dim name As String = udtForm.globals(key).name
            Dim line As Integer = udtForm.globals(key).line

            'If its the first instance of a global variable, then let's add the head above it
            If firstGlobal = True Then
                firstGlobal = False

                insertLine(line - 1 + addedLines, "") 'Add an empty line
                insertLine(line - 1 + addedLines, spaces + "'-------------------------------------------------------------------------------------------")
                insertLine(line - 1 + addedLines, spaces + "'--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---")
                insertLine(line - 1 + addedLines, spaces + "'--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---")
                insertLine(line - 1 + addedLines, spaces + "'--- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES --- GLOBAL VARIABLES ---")
                insertLine(line - 1 + addedLines, spaces + "'-------------------------------------------------------------------------------------------")
                insertLine(line - 1 + addedLines, "") 'Add an empty line

            End If

            'Get the current value of the text (because we are going to be replacing the line)
            Dim text As String = copyFileText(line + addedLines - 1)
            'Now check if on the line if the contains a apostrophe with spaces on both sides. If so, get the first index of the split.
            If text.Contains(Space(1) + "'" + Space(1)) Then
                text = text.Split(Space(1) + "'" + Space(1), 2)(0)
            End If
            'MessageBox.Show("Line: " + (line + addedLines).ToString + vbCrLf + "Text: " + text + vbCrLf + "Desc: " + desc + vbCrLf + "Name: " + name)
            'Now remove the line that was saved (because it will be readded in a moment)
            copyFileText.RemoveAt(line + addedLines - 1)
            'Also remove a line from the list of total added
            addedLines = addedLines - 1

            'Now insert the description (with the code of the variable, which is text)
            insertLines(line + addedLines, generateInlineWordWrapping(text, desc))
        Next

        'Methods
        Dim firstMethod As Boolean = True
        For Each key In udtForm.methods.Keys()
            Dim line As Integer = udtForm.methods(key).line
            Dim method As GlobalMethods = udtForm.methods(key)

            'Are we in the first method? If so let's add the header of SUBPROGRAMS above it.
            If firstMethod = True Then
                firstMethod = False

                insertLine(line - 1 + addedLines, "") 'Add an empty line
                insertLine(line - 1 + addedLines, spaces + "'-------------------------------------------------------------------")
                insertLine(line - 1 + addedLines, spaces + "'--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---")
                insertLine(line - 1 + addedLines, spaces + "'--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---")
                insertLine(line - 1 + addedLines, spaces + "'--- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS --- SUBPROGRAMS ---")
                insertLine(line - 1 + addedLines, spaces + "'-------------------------------------------------------------------")
                insertLine(line - 1 + addedLines, "") 'Add an empty line
                spaces = "        "
            End If

            'Subprogram's Name
            insertLine(line + addedLines, "")
            insertLine(line + addedLines, spaces + BLOCK_ROW)
            insertLine(line + addedLines, spaces + PadCenterComment("Subprogram Name: " + method.name))
            insertLine(line + addedLines, spaces + BLOCK_ROW)

            'Written By and Written On
            insertLine(line + addedLines, spaces + PadCenterComment("Written By: " + method.writtenBy))
            insertLine(line + addedLines, spaces + PadCenterComment("Written On: " + method.writtenOn))
            insertLine(line + addedLines, spaces + BLOCK_ROW)

            'Subprogram Purpose
            insertLine(line + addedLines, spaces + PadRightComment("Subprogram Purpose:"))
            insertLine(line + addedLines, spaces + PadRightComment(""))
            insertLinesSpaces(line + addedLines, generateWordWrappingBlock(method.purpose), spaces)
            insertLine(line + addedLines, spaces + BLOCK_ROW)

            'Parameters
            insertLine(line + addedLines, spaces + PadRightComment("Parameter Dictionary (in parameter order):"))
            For Each param_key In udtForm.methods(key).parameters.Keys()
                Dim name As String = udtForm.methods(key).parameters(param_key).name
                Dim desc As String = udtForm.methods(key).parameters(param_key).desc

                insertLinesSpaces(line + addedLines, generateVariableWithDescWrapping(name, desc), spaces)
            Next

            'Check if we have no parameters, if so just list none
            If udtForm.methods(key).parameters.Count = 0 Then
                insertLine(line + addedLines, spaces + PadRightComment("(None)"))
            End If
            insertLine(line + addedLines, spaces + BLOCK_ROW)

            'Local Variables
            insertLine(line + addedLines, spaces + PadRightComment("Local Variable Dictionary (alphabetically):"))

            'Sort the keys alphabetically
            Dim keys As List(Of String) = method.dimensions.Keys.ToList
            keys.Sort()

            'Iterate over all of the dimension keys, and insert the descriptions with word wrapping
            For Each dim_key As String In keys
                Dim name As String = udtForm.methods(key).dimensions(dim_key).name
                Dim desc As String = udtForm.methods(key).dimensions(dim_key).desc
                insertLinesSpaces(line + addedLines, generateVariableWithDescWrapping(name, desc), spaces)
            Next
            'Check if we have no dimensions, if so just list none
            If keys.Count = 0 Then
                insertLine(line + addedLines, spaces + PadRightComment("(None)"))
            End If
            insertLine(line + addedLines, spaces + BLOCK_ROW)

            'Check if the method has a return, if so add the return description
            If method.hasReturn = True Then
                insertLine(line + addedLines, spaces + PadRightComment("Returns:"))
                insertLinesSpaces(line + addedLines, generateVariableWithDescWrapping(method.returnType, method.returnDescription), spaces)
                insertLine(line + addedLines, spaces + BLOCK_ROW)
            End If
            insertLine(line + addedLines, "")
        Next

        Dim tempFilePath As String = System.IO.Path.Combine(path, strProgramFilePath)
        'Dim tempFile As FileStream = System.IO.File.Open(tempFilePath, FileMode.OpenOrCreate, FileAccess.ReadWrite)
        Using sr As StreamWriter = New StreamWriter(tempFilePath)
            Dim lastLineWasEmpty As String = False
            For Each s As String In copyFileText
                'MessageBox.Show(s.Count())
                If s.Count() < 3 = True Then
                    If lastLineWasEmpty = False Then
                        sr.Write(s)
                    End If
                Else
                    sr.Write(s)
                End If
                If s.Count() = 2 Then
                    lastLineWasEmpty = True
                Else
                    lastLineWasEmpty = False
                End If

            Next
        End Using

    End Sub

    Private Sub log(message As String)

        '------------------------------------------------------------
        '-                   Subprogram Name: log                   -
        '------------------------------------------------------------
        '-                       Written By:                        -
        '-                       Written On:                        -
        '------------------------------------------------------------
        '- Subprogram Purpose:                                      -
        '-                                                          -
        '-                                                          -
        '------------------------------------------------------------
        '- Parameter Dictionary (in parameter order):               -
        '- message -                                                -
        '------------------------------------------------------------
        '- Local Variable Dictionary (alphabetically):              -
        '- (None)                                                   -
        '------------------------------------------------------------

        lstConsole.Items.Add(message)
        lstConsole.TopIndex = lstConsole.Items.Count - 1
    End Sub
End Class
