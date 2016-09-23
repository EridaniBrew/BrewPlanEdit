Imports System
Imports System.IO

Public Class CStringCommand : Inherits CommandBaseClass
    ' this class will be further subclassed for specific commands:
    ' allows single implementation of the commands using various File and Folder dialog boxes
    
    Protected _dataString As String                   ' string associated with this command
    Protected _origDataString As String

    Protected _blankAllowed As Boolean = False


    Public Sub New()
        ' must be format #cmd int date time   ' optional comment here
        MyBase.New()

        _dataString = ""
        _origDataString = ""
        _blankAllowed = False
        
    End Sub

    Public Function UpdateDataString(curText As String) As Boolean
        ' if the string has changed, update the value
        ' returns True if update occurred
        ' curText is the new string
        Dim updated As Boolean = False

        Dim errMsg As String = ""
        If ((curText = "") And (Not _blankAllowed)) Then
            Throw New Exception("Blank field not allowed ")        ' field is required
        End If
        If (curText <> _origDataString) Then
            _dataString = curText
            updated = True
            MainForm.GetPlan().Update()       ' update the plan
        End If
        Return updated
    End Function


    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.pnlStringCommand.Visible = True
        MainForm.txtStrText.Text = _dataString
        
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase._command & " " & _dataString
        Return lineText
    End Function



#Region "Field Control Events"
    Private WithEvents myTxtStrText As TextBox = MainForm.txtStrText

    Private Sub txtDirPath_KeyPress(sender As Object, e As KeyPressEventArgs) Handles myTxtStrText.KeyPress
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            If (e.KeyChar = vbCr) Then
                txtStrText_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub txtStrText_Enter(sender As Object, e As EventArgs) Handles myTxtStrText.Enter
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            _origDataString = Trim(myTxtStrText.Text)
        End If
    End Sub

    Public Overridable Sub txtStrText_Leave(sender As Object, e As EventArgs) Handles myTxtStrText.Leave
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            Try
                Dim s As String = myTxtStrText.Text
                UpdateDataString(Trim(s))
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Data Error")
            End Try
        End If
    End Sub

#End Region
End Class

#Region "CTag"
'=============================================================================================================
Public Class CTag : Inherits CStringCommand
    Public Sub New(cmd As String)
        ' must be format #cmd optFilePath   ' optional comment here
        MyBase.New()

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#tag") Then
            If (cmd.Length > cmdname.Length) Then
                Dim str As String = cmd.Substring(cmdname.Length + 1)
                UpdateDataString(Trim(str))
            Else
                Throw New Exception("#tag command missing string {" & cmd & "}")
            End If
        Else
            Throw New Exception("Invalid #tag command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.lblStrLabel.Text = "Name = Value pair"
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase.ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ;" & MyBase._comment
        End If
        Return lineText
    End Function

    Private Function LocalUpdateDataString(s As String) As Boolean
        ' Check target name for "=" sign
        Dim updated As Boolean = False
        Dim fields() As String = s.Split("=")
        If (fields.Length <> 2) Then
            Throw New Exception("#tag command should have 2 fields separated by = {" & s & "}")
        End If
        Return MyBase.UpdateDataString(s)
    End Function


    Private WithEvents locTxtStrText As TextBox = MainForm.txtStrText

    Public Overrides Sub txtStrText_Leave(sender As Object, e As EventArgs) Handles locTxtStrText.Leave
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            Try
                Dim s As String = locTxtStrText.Text
                LocalUpdateDataString(Trim(s))
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Data Error")
            End Try
        End If
    End Sub


End Class
#End Region

#Region "CReadoutMode"
'=============================================================================================================
Public Class CReadoutMode : Inherits CStringCommand
    Public Sub New(cmd As String)
        ' must be format #cmd optFilePath   ' optional comment here
        MyBase.New()

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#readoutmode") Then
            If (cmd.Length > cmdname.Length) Then
                Dim str As String = cmd.Substring(cmdname.Length + 1)
                UpdateDataString(Trim(str))
            Else
                Throw New Exception("#readoutmode command missing string {" & cmd & "}")
            End If
        Else
            Throw New Exception("Invalid #readoutmode command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.lblStrLabel.Text = "Readout Mode"
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase.ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ;" & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CMinSetTime"
'=============================================================================================================
Public Class CMinSetTime : Inherits CStringCommand
    Public Sub New(cmd As String)
        ' must be format #cmd optFilePath   ' optional comment here
        MyBase.New()

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#minsettime") Then
            If (cmd.Length > cmdname.Length) Then
                Dim str As String = cmd.Substring(cmdname.Length + 1)
                UpdateDataString(Trim(str))
            Else
                Throw New Exception("#minsettime command missing time {" & cmd & "}")
            End If
        Else
            Throw New Exception("Invalid #minsettime command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.lblStrLabel.Text = "Minimum Time per set (00:05 minutes)"
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase.ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ;" & MyBase._comment
        End If
        Return lineText
    End Function

    Private Function LocalUpdateDataString(s As String) As Boolean
        ' Check target name for 00:00
        Dim updated As Boolean = False
        Dim fields() As String = s.Split(":")
        If (fields.Length <> 2) Then
            Throw New Exception("#minsettime command invalid. Should have 2 integer fields separated by ':' {" & s & "}")
        End If
        Dim field1, field2 As Integer
        If ((Not Integer.TryParse(fields(0), field1)) Or (Not Integer.TryParse(fields(1), field2))) Then
            Throw New Exception("#minsettime command invalid. Should have 2 integer fields separated by ':' {" & s & "}")
        End If
        If (field1 < 0) Then
            Throw New Exception("#minsettime command invalid. First field (hours) must be positive integer {" & s & "}")
        End If
        If ((field2 < 0) Or (field2 > 59)) Then
            Throw New Exception("#minsettime command invalid. 2nd field (minutes) invalid range 0:59 {" & s & "}")
        End If
        Return MyBase.UpdateDataString(s)
    End Function


    Private WithEvents locTxtStrText As TextBox = MainForm.txtStrText

    Public Overrides Sub txtStrText_Leave(sender As Object, e As EventArgs) Handles locTxtStrText.Leave
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            Try
                Dim s As String = locTxtStrText.Text
                LocalUpdateDataString(Trim(s))
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Data Error")
            End Try
        End If
    End Sub


End Class
#End Region

#Region "CManual"
'=============================================================================================================
Public Class CManual : Inherits CStringCommand
    Public Sub New(cmd As String)
        ' must be format #cmd optFilePath   ' optional comment here
        MyBase.New()
        _blankAllowed = True

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#manual") Then
            If (cmd.Length > cmdname.Length) Then
                Dim str As String = cmd.Substring(cmdname.Length + 1)
                UpdateDataString(Trim(str))
            Else
                UpdateDataString("")
            End If
        Else
            Throw New Exception("Invalid #manual command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.lblStrLabel.Text = "Manual Target Name"
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase.ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ;" & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region
