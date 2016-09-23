Imports System
Imports System.IO

Public Class CPlan
    Private _fileName As String
    Private _filePath As String
    Private _commandList As Collection
    Private _lstBox As ListBox
    Private _planModified As Boolean

    Private _BaseCommand As CommandBaseClass

    Public Property FileName() As String
        Get
            Return _fileName
        End Get
        Set(value As String)
            _fileName = value
        End Set
    End Property

    Public Property ListBox() As ListBox
        Get
            Return _lstBox
        End Get
        Set(value As ListBox)
            _lstBox = value
        End Set
    End Property

    Public Property Modified() As Boolean
        Get
            Return _planModified
        End Get
        Set(value As Boolean)
            _planModified = value
        End Set
    End Property

    Public Sub New()
        _commandList = New Collection
        _fileName = ""
        _lstBox = Nothing
        _planModified = False
        _filePath = ""
        _BaseCommand = New CommandBaseClass()
    End Sub

    Private Function AccumComboCommand(cmd As String, s As String, ByRef comboCommand As String) As Boolean
        ' returns  true if command was accumulated
        ' false if a different command
        Dim ret As Boolean = False
        Dim lcmd As String = LCase(cmd)

        If ((lcmd = "#count") Or (lcmd = "#interval") Or (lcmd = "#binning") Or (lcmd = "#filter")) Then
            ret = True
            If (comboCommand = "") Then
                ' first entry in command
                comboCommand = s
            Else
                ' add to existing command
                comboCommand = comboCommand & "|" & s
            End If
        End If

        Return ret
    End Function

    Private Function CreateCommandObject(s As String, cmdType As String, lineNum As Integer) As String
        ' create the object described by the string
        'create the command object (i.e., a CChill object
        Dim errS As String = ""
        Dim obj As Object

        Try
            obj = Activator.CreateInstance(Type.GetType("BrewPlanEdit." & cmdType), s)
            _lstBox.Items.Add(obj)
        Catch ex As Exception
            errS = "Invalid command line " & lineNum.ToString() & " {" & s & "}" & vbCrLf & ex.Message & vbCrLf
            ' make this a comment object
            s = ";" & s
            cmdType = "CFlatCommand"
            Try
                Dim comtobj As CFlatCommand = Activator.CreateInstance(Type.GetType("BrewPlanEdit." & cmdType), s)
                comtobj.Warning = True
                _lstBox.Items.Add(comtobj)
            Catch e As Exception
                MsgBox("Conversion to comment failed {" & s & "} err " & e.Message)
            End Try

        End Try
        Return errS
    End Function

    Public Function OpenPlan() As String
        ' pop dialog box, open selected file. If file selected,
        ' rreturn the base file name (plan.txt) rather than the complete path
        ' returns "" if no file selected
        
        Dim result As Windows.Forms.DialogResult = MainForm.OpenFileDialog.ShowDialog()
        Dim retFile As String = ""
        If (result = DialogResult.OK) Then
            ' open the file, load the list box
            _filePath = MainForm.OpenFileDialog.FileName
            _fileName = Path.GetFileName(_filePath)
            retFile = _fileName
            Dim reader As StreamReader = My.Computer.FileSystem.OpenTextFileReader(_filePath, Text.Encoding.ASCII)
            Dim s As String
            Dim errS As String = ""
            Dim lineNum As Integer = 0
            Dim comboCommand As String = ""

            _lstBox.Items.Clear()
            s = reader.ReadLine
            lineNum = lineNum + 1
            While (s IsNot Nothing)
                If (Trim(s) = "") Then
                    ' skip blank lines
                ElseIf (s(0) = ";") Then
                    ' comment line
                    Try
                        _lstBox.Items.Add(New CFlatCommand(s))
                    Catch ex As Exception
                        errS = errS & "Invalid comment line {" & s & "}" & vbCrLf & ex.Message & vbCrLf
                    End Try
                ElseIf (s(0) = "#") Then
                    ' command
                    Dim pieces As String() = s.Split(" ")
                    Dim cmdIdx As Integer = -1
                    cmdIdx = _BaseCommand.CmdFindCommand(pieces(0))
                    If (cmdIdx >= 0) Then
                        ' special case - if the file has #count/#interval/#bin/#filter command, combine these together into a
                        ' single "combo" command
                        If (Not AccumComboCommand(pieces(0), s, comboCommand)) Then
                            ' not an accumulated command
                            errS = errS & CreateCommandObject(s, _BaseCommand.cmdParms(cmdIdx, CommandBaseClass.CMDPARMType), lineNum)
                        Else
                            ' command was combined
                        End If

                    Else
                        ' Did not find in list
                        errS = errS & "Command not recognized line " & lineNum.ToString() & "  {" & s & "}" & vbCrLf & "Converting it to a comment " & vbCrLf
                        ' Make it a comment
                        s = ";" & s
                        Try
                            Dim comtobj As CFlatCommand = New CFlatCommand(s)
                            comtobj.Warning = True
                            _lstBox.Items.Add(comtobj)
                        Catch ex As Exception
                            errS = errS & "Invalid comment line {" & s & "}" & vbCrLf & ex.Message & vbCrLf
                        End Try
                    End If
                Else
                    ' target command
                    If (comboCommand <> "") Then
                        ' need to handle combo command
                        errS = errS & CreateCommandObject(comboCommand, "CCombo", lineNum)
                        comboCommand = ""
                    End If
                    ' now do the target command
                    errS = errS & CreateCommandObject(s, "CTarget", lineNum)
                End If
                s = reader.ReadLine
                lineNum = lineNum + 1
            End While

            ' do we still have a comboCommand needing to be done?
            If (comboCommand <> "") Then
                ' need to handle combo command
                errS = errS & CreateCommandObject(comboCommand, "CCombo", lineNum)
                comboCommand = ""
            End If
            If (errS <> "") Then
                MsgBox("Errors found in Plan" & vbCrLf & vbCrLf & errS, MsgBoxStyle.Critical, "Plan Error")
            End If
            _lstBox.SelectedIndex = 0
            reader.Close()
        End If
        OpenPlan = retFile
    End Function

    Public Sub SavePlan()
        ' Rewriting plan to the original file name
        Dim writer As StreamWriter = My.Computer.FileSystem.OpenTextFileWriter(_filePath, False, Text.Encoding.ASCII)   ' false for append mode
        Dim s As String
        Dim i As Integer

        For i = 0 To _lstBox.Items.Count - 1
            s = _lstBox.Items(i).ToString()
            writer.WriteLine(s)
        Next
        writer.Close()

        ResetUpdatedFlag()
    End Sub

    Public Sub SaveAsPlan(filePath As String)
        _filePath = filePath
        _fileName = Path.GetFileName(_filePath)
        SavePlan()
    End Sub


    Public Sub Update()
        ' the plan has been updated.
        ' Refresh the listbox
        ' Change the _planModified flag to True,
        ' change the tab name to have "*"
        If (_lstBox.SelectedIndex >= 0) Then
            _lstBox.Items(_lstBox.SelectedIndex) = _lstBox.SelectedItem
        End If

        If (_planModified = False) Then
            _planModified = True
            Dim tabName As String = MainForm.TabControl1.SelectedTab.Text
            If (tabName.IndexOf("*") < 0) Then
                tabName = Trim(tabName) & "*   "
                MainForm.TabControl1.SelectedTab.Text = tabName
            End If
        End If

    End Sub

    Public Sub AddDefault(cmd As String)
        Dim errS As String = ""
        Dim curIdx As Integer = _lstBox.SelectedIndex

        If (cmd(0) = ";") Then
            ' comment line
            Try
                _lstBox.Items.Insert(curIdx + 1, New CFlatCommand(cmd))
            Catch ex As Exception
                errS = "Invalid comment line {" & cmd & "}" & vbCrLf & ex.Message & vbCrLf
            End Try
        ElseIf (cmd(0) = "#") Then
            ' command
            Dim cmdIdx As Integer = -1
            Dim defCmd As String
            cmdIdx = _BaseCommand.CmdFindCommand(cmd)
            If (cmdIdx >= 0) Then
                defCmd = _BaseCommand.cmdParms(cmdIdx, CommandBaseClass.CMDPARMDefaultString)
                Try
                    _lstBox.Items.Insert(curIdx + 1, Activator.CreateInstance(Type.GetType("BrewPlanEdit." & _BaseCommand.cmdParms(cmdIdx, CommandBaseClass.CMDPARMType)), defCmd))
                Catch ex As Exception
                    errS = "Command failed {" & defCmd & "}" & vbCrLf & ex.Message & vbCrLf
                End Try
            Else
                ' command not in table
                errS = "Command not implemented {" & cmd & "}"
            End If

        Else
            ' must be a target
            Try
                _lstBox.Items.Insert(curIdx + 1, New CTarget(cmd))
            Catch ex As Exception
                errS = "Target command failed {" & cmd & "}" & vbCrLf & ex.Message & vbCrLf
            End Try
        End If

        If (errS <> "") Then
            MsgBox("Command add failed" & vbCrLf & vbCrLf & errS, MsgBoxStyle.Critical, "Command Error")
        Else
            _lstBox.SelectedIndex = curIdx + 1
            Update()
        End If

    End Sub


    Private Sub ResetUpdatedFlag()
        _planModified = False
        Dim tabName As String = MainForm.TabControl1.SelectedTab.Text
        If (tabName.IndexOf("*") > -1) Then
            tabName = tabName.Remove(tabName.IndexOf("*"))
            MainForm.TabControl1.SelectedTab.Text = Trim(tabName) & "   "
        End If
    End Sub

    Public Sub CommentSelectedLine()
        ' we should have a line selected
        Dim idx As Integer = _lstBox.SelectedIndex
        If (idx >= 0) Then
            Dim s As String = ";" & _lstBox.Text
            s = s.Replace(vbCrLf, vbCrLf & ";")               ' CCombo object needs vbcrlf changed to vbcrlf;
            Dim commentObj As CFlatCommand = New CFlatCommand(s)
            _lstBox.Items.RemoveAt(idx)
            _lstBox.Items.Insert(idx, commentObj)
            _lstBox.SelectedIndex = idx
            Update()
        End If
    End Sub

    Public Sub UncommentSelectedLine()
        Dim nextS As String
        Dim nextCmd() As String
        Dim delIdx As Integer

        Dim idx As Integer = _lstBox.SelectedIndex
        If (idx >= 0) Then     ' should have a selected line
            ' what type of comment?
            Dim s As String = _lstBox.Text
            s = s.Substring(1)     ' remove the initial ;
            Dim pieces() As String = s.Split(" ")
            Dim cmdName As String = LCase(Trim(pieces(0)))
            If (cmdName(0) = "#") Then
                ' we have a command
                If ((cmdName = "#count") Or (cmdName = "#interval") Or (cmdName = "#binning") Or (cmdName = "#filter")) Then
                    ' have either
                    ' 1) #count ... ;#interval ...;#binning ... ;#filter ...    (recently commented CCombo)
                    ' or 2) ;#count ...
                    '       ;#interval ...
                    '       ;# binning ...
                    '       ;#filter ...       not necessarily in this order; commented CCombo saved to file
                    'need to append the next commands if option 2
                    Dim deleteList As Collection = New Collection
                    If (s.IndexOf(";#") = -1) Then
                        ' we have option 2
                        Dim i As Integer
                        For i = idx + 1 To Math.Min(idx + 3, _lstBox.Items.Count - 1)
                            nextS = _lstBox.Items(i).ToString()
                            nextS = nextS.Substring(1)     ' remove initial ;
                            nextCmd = nextS.Split(" ")
                            nextCmd(0) = LCase(Trim(nextCmd(0)))
                            If ((nextCmd(0) = "#count") Or (nextCmd(0) = "#interval") Or (nextCmd(0) = "#binning") Or (nextCmd(0) = "#filter")) Then
                                s = s & "|" & nextS
                                ' this lstbox entry needs to be removed
                                deleteList.Add(i)
                            End If
                        Next
                    Else
                        ' option 1
                        s = s.Replace(";", "|")
                    End If
                    ReplaceCommand(idx, cmdName, s)
                    For i = deleteList.Count To 1 Step -1
                        delIdx = deleteList.Item(i)
                        _lstBox.Items.RemoveAt(delIdx)
                    Next
                    _lstBox.SelectedIndex = idx
                    Update()
                Else
                    ' regular 1 line command
                    ReplaceCommand(idx, cmdName, s)
                    Update()
                End If

            ElseIf (cmdName(0) = ";") Then
                ' leave as comment? Just edit the comment!

            Else
                ' OK, this should be a target
                Dim targObj As CTarget = New CTarget(s)
                _lstBox.Items.Insert(idx + 1, targObj)
                _lstBox.Items.RemoveAt(idx)
                _lstBox.SelectedIndex = idx
                Update()
            End If

        End If
    End Sub

    Private Sub ReplaceCommand(idx As Integer, cmdName As String, cmdLine As String)
        Dim cmdIdx As Integer = -1
        Dim errS As String = ""

        cmdIdx = _BaseCommand.CmdFindCommand(cmdName)
        If (cmdIdx >= 0) Then
            Try
                _lstBox.Items.Insert(idx + 1, Activator.CreateInstance(Type.GetType("BrewPlanEdit." & _BaseCommand.cmdParms(cmdIdx, CommandBaseClass.CMDPARMType)), cmdLine))
            Catch ex As Exception
                errS = "ReplaceCommand failed {" & cmdLine & "}" & vbCrLf & ex.Message & vbCrLf
            End Try
            _lstBox.Items.RemoveAt(idx)
            _lstBox.SelectedIndex = idx
        Else
            ' command not in table
            errS = "ReplaceCommand not implemented {" & cmdLine & "}"
        End If

    End Sub
End Class
