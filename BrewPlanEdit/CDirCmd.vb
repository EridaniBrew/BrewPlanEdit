Imports System
Imports System.IO

Public Class CDirCmd : Inherits CommandBaseClass
    ' this class will be further subclassed for specific commands:
    ' allows single implementation of the commands using various File and Folder dialog boxes
    Protected Const FileDialog = 0       ' Using a File dialog box

    Public Const FILE_DIALOG = 0       ' index of first Int field
    Public Const FOLDER_DIALOG = 1       ' index of second Int field

    Protected _dialogType As Integer            ' which type of dialog?
    Protected _useFullPath As Boolean              ' True = use the full path in the directive
    Protected _blankAllowed As Boolean          ' Is a blank field allowed?

    Protected _path As String                   ' file or dir path
    Protected _fullPath As String
    Protected _baseName As String = ""
    Protected _origPath As String



    Public Sub New()
        ' must be format #cmd int date time   ' optional comment here
        MyBase.New()

        _path = ""
        _fullPath = ""
        _baseName = ""
        _useFullPath = False
        _dialogType = FILE_DIALOG
        _blankAllowed = True

    End Sub

    Public Function UpdateFilePath(curText As String) As Boolean
        ' if the string has changed, update the value
        ' returns True if update occurred
        ' curText is a path
        Dim updated As Boolean = False

        Dim errMsg As String = ""
        If ((curText = "") And (Not _blankAllowed)) Then
            Throw New Exception("Blank path not allowed ")        ' field is required
        End If
        _fullPath = curText
        _baseName = Path.GetFileName(curText)
        If (_useFullPath) Then
            _path = _fullPath
        Else
            _path = _baseName
        End If
        _origPath = _path
        myTxtDirPath.Text = _path
        updated = True
        MainForm.GetPlan().Update()       ' update the plan
        Return updated
    End Function


    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.pnlDirCommand.Visible = True
        MainForm.txtDirPath.Text = _path
        MainForm.cbDirFullPath.Checked = _useFullPath

    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase._command & " " & _path
        Return lineText
    End Function



#Region "Field Control Events"
    Private WithEvents myBtnDirDialog As Button = MainForm.btnDirDialog
    Private WithEvents myCbDirFullPath As CheckBox = MainForm.cbDirFullPath
    Private WithEvents myTxtDirPath As TextBox = MainForm.txtDirPath

    Private Sub txtDirPath_KeyPress(sender As Object, e As KeyPressEventArgs) Handles myTxtDirPath.KeyPress
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            If (e.KeyChar = vbCr) Then
                txtDirPath_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub txtDirPath_Enter(sender As Object, e As EventArgs) Handles myTxtDirPath.Enter
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            _origPath = Trim(myTxtDirPath.Text)
        End If
    End Sub

    Private Sub txtDirPath_Leave(sender As Object, e As EventArgs) Handles myTxtDirPath.Leave
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            Try
                Dim s As String = myTxtDirPath.Text
                UpdateFilePath(Trim(s))
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Data Error")
            End Try
        End If
    End Sub


    Private Sub cbDirFullPath_CheckedChanged(sender As Object, e As EventArgs) Handles myCbDirFullPath.CheckedChanged
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            _useFullPath = myCbDirFullPath.Checked
            UpdateFilePath(_fullPath)
            MainForm.GetPlan().Update()
        End If
    End Sub

    Private Sub btnDirDialog_Click(sender As Object, e As EventArgs) Handles myBtnDirDialog.Click
        Dim ret As Windows.Forms.DialogResult
        Dim mypath As String
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            If (_dialogType = FILE_DIALOG) Then
                ret = MainForm.OpenPathDialog.ShowDialog()
                If (ret = DialogResult.OK) Then
                    mypath = MainForm.OpenPathDialog.FileName
                    UpdateFilePath(mypath)
                End If
            ElseIf (_dialogType = FOLDER_DIALOG) Then
                ret = MainForm.OpenDirDialog.ShowDialog()
                If (ret = DialogResult.OK) Then
                    mypath = MainForm.OpenDirDialog.SelectedPath
                    UpdateFilePath(mypath)
                End If
            End If
        End If
    End Sub

#End Region
End Class

#Region "CDir"
'=============================================================================================================
Public Class CDir : Inherits CDirCmd
    Public Sub New(cmd As String)
        ' must be format #cmd optFilePath   ' optional comment here
        MyBase.New()
        _blankAllowed = True
        _dialogType = FOLDER_DIALOG
        _useFullPath = False

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#dir") Then
            If (cmd.Length > cmdname.Length) Then
                Dim pathname As String = cmd.Substring(cmdname.Length + 1)
                MyBase.UpdateFilePath(Trim(pathname))
            Else
                MyBase.UpdateFilePath("")
            End If
        Else
            Throw New Exception("Invalid #dir command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()

    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase.ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region



#Region "CDark"
'=============================================================================================================
Public Class CDark : Inherits CDirCmd
    Public Sub New(cmd As String)
        ' must be format #cmd optFilePath   ' optional comment here
        MyBase.New()
        _blankAllowed = True
        _dialogType = FILE_DIALOG
        _useFullPath = False

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#dark") Then
            If (cmd.Length > cmdname.Length) Then
                Dim pathname As String = cmd.Substring(cmdname.Length + 1)
                MyBase.UpdateFilePath(Trim(pathname))
            Else
                MyBase.UpdateFilePath("")
            End If
        Else
            Throw New Exception("Invalid #dark command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()

    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase.ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CBias"
'=============================================================================================================
Public Class CBias : Inherits CDirCmd
    Public Sub New(cmd As String)
        ' must be format #cmd optFilePath   ' optional comment here
        MyBase.New()
        _blankAllowed = True
        _dialogType = FILE_DIALOG
        _useFullPath = False

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#bias") Then
            If (cmd.Length > cmdname.Length) Then
                Dim pathname As String = cmd.Substring(cmdname.Length + 1)
                MyBase.UpdateFilePath(Trim(pathname))
            Else
                MyBase.UpdateFilePath("")
            End If
        Else
            Throw New Exception("Invalid #bias command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()

    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase.ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CChain"
'=============================================================================================================
Public Class CChain : Inherits CDirCmd
    Public Sub New(cmd As String)
        ' must be format #cmd optFilePath   ' optional comment here
        MyBase.New()
        _blankAllowed = True
        _dialogType = FILE_DIALOG
        _useFullPath = False

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#chain") Then
            If (cmd.Length > cmdname.Length) Then
                Dim pathname As String = cmd.Substring(cmdname.Length + 1)
                MyBase.UpdateFilePath(Trim(pathname))
            Else
                MyBase.UpdateFilePath("")
            End If
        Else
            Throw New Exception("Invalid #chain command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()

    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase.ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CChainScript"
'=============================================================================================================
Public Class CChainScript : Inherits CDirCmd
    Public Sub New(cmd As String)
        ' must be format #cmd optFilePath   ' optional comment here
        MyBase.New()
        _blankAllowed = True
        _dialogType = FILE_DIALOG
        _useFullPath = False

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#chainscript") Then
            If (cmd.Length > cmdname.Length) Then
                Dim pathname As String = cmd.Substring(cmdname.Length + 1)
                MyBase.UpdateFilePath(Trim(pathname))
            Else
                MyBase.UpdateFilePath("")
            End If
        Else
            Throw New Exception("Invalid #chainscript command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()

    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase.ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CDuskFlats"
'=============================================================================================================
Public Class CDuskFlats : Inherits CDirCmd
    Public Sub New(cmd As String)
        ' must be format #cmd optFilePath   ' optional comment here
        MyBase.New()
        _blankAllowed = True
        _dialogType = FILE_DIALOG
        _useFullPath = False

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#duskflats") Then
            If (cmd.Length > cmdname.Length) Then
                Dim pathname As String = cmd.Substring(cmdname.Length + 1)
                MyBase.UpdateFilePath(Trim(pathname))
            Else
                MyBase.UpdateFilePath("")
            End If
        Else
            Throw New Exception("Invalid #duskflats command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()

    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase.ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CDawnFlats"
'=============================================================================================================
Public Class CDawnFlats : Inherits CDirCmd
    Public Sub New(cmd As String)
        ' must be format #cmd optFilePath   ' optional comment here
        MyBase.New()
        _blankAllowed = True
        _dialogType = FILE_DIALOG
        _useFullPath = False

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#dawnflats") Then
            If (cmd.Length > cmdname.Length) Then
                Dim pathname As String = cmd.Substring(cmdname.Length + 1)
                MyBase.UpdateFilePath(Trim(pathname))
            Else
                MyBase.UpdateFilePath("")
            End If
        Else
            Throw New Exception("Invalid #dawnflats command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()

    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase.ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CScreenFlats"
'=============================================================================================================
Public Class CScreenFlats : Inherits CDirCmd
    Public Sub New(cmd As String)
        ' must be format #cmd optFilePath   ' optional comment here
        MyBase.New()
        _blankAllowed = True
        _dialogType = FILE_DIALOG
        _useFullPath = False

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#screenflats") Then
            If (cmd.Length > cmdname.Length) Then
                Dim pathname As String = cmd.Substring(cmdname.Length + 1)
                MyBase.UpdateFilePath(Trim(pathname))
            Else
                MyBase.UpdateFilePath("")
            End If
        Else
            Throw New Exception("Invalid #screenflats command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()

    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase.ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region


