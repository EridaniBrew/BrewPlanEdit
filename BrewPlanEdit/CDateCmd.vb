Imports System.Globalization

Public Class CDateCmd : Inherits CommandBaseClass
    ' this class will be further subclassed for specific commands:
    '        waituntil 0,ddd
    '        quitat ddd
    '        shutdownat ddd
    ' allows single implementation of the date/time controls

    Protected _theDate As String = ""
    Private _origDate As String = ""
    Protected _intVal As Integer = 0
    Private _origIntVal As String = ""
    Protected _minIntVal As Integer = 0         ' subclass needs to set these
    Protected _maxIntVal As Integer = 100
    Protected _intFieldName As String = "Int Field"         ' used for error messages
    Protected _useFullDate As Boolean = False

    ' Format strings for Date control
    Protected Const FULL_DATE_FORMAT = "MM/dd/yy hh:mm tt"
    Protected Const TIME_ONLY_FORMAT = "hh:mm tt"

    Public Function UpdateIntVal(curText As String) As Boolean
        ' if the string has changed, update the value
        ' returns True if update occurred
        Dim updated As Boolean = False
        Dim tempVal As Integer

        If (curText <> _origIntVal) Then
            Dim errMsg As String = ""
            If (curText = "") Then
                Throw New Exception("Blank " & _intFieldName & " field not allowed ")        ' field is required
            End If
            If (Integer.TryParse(curText, tempVal)) Then
                If ((tempVal < _minIntVal) Or (tempVal > _maxIntVal)) Then
                    Throw New Exception("Field " & _intFieldName & " out of range {" & curText & "}" & vbCrLf & "Allowed range is " & _minIntVal & ".." & _maxIntVal)
                Else
                    _intVal = tempVal
                    _origIntVal = curText
                    updated = True
                End If
            Else
                Throw New Exception("Non-integer Field " & _intFieldName & " value {" & curText & "}" & vbCrLf & errMsg)
            End If
        End If
        Return updated
    End Function

    Public Function UpdateDateVal(curText As String) As Boolean
        ' if the string has changed, update the  value
        ' returns True if update occurred
        Dim updated As Boolean = False
        If (curText <> _origDate) Then
            ' verify date is after Now
            Dim curTimeZone As TimeZone = TimeZone.CurrentTimeZone
            Dim locNow As DateTime = curTimeZone.ToLocalTime(DateTime.Now)
            Dim valDate As DateTime
            DateTime.TryParse(curText, CultureInfo.CurrentCulture, DateTimeStyles.None, valDate)
            If ((DateTime.Compare(valDate, locNow) < 0) And (_useFullDate)) Then
                ' curText is earlier than now!
                'MsgBox("Warning: Date selected {" & curText & "} is before Now ", MsgBoxStyle.Exclamation, "Date Warning")
                MainForm.lblDatewarning.Visible = True
                MyBase.Warning = True
            Else
                MainForm.lblDatewarning.Visible = False
                MyBase.Warning = False
            End If
            _theDate = curText
            _origDate = curText
            updated = True

        End If
        Return updated
    End Function

    Public Sub New(cmd As String)
        ' must be format #cmd int datetime   ' optional comment here
        MyBase.New()
        MainForm.lblDatewarning.Visible = False
        If (cmd.IndexOf("/") > -1) Then
            ' we have a full date directive
            _useFullDate = True
            myDtTimeCmd.CustomFormat = FULL_DATE_FORMAT
        Else
            ' we have a full date directive
            _useFullDate = False
            myDtTimeCmd.CustomFormat = TIME_ONLY_FORMAT
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.pnlDateCmd.Visible = True
        MainForm.dtTimeCmd.Text = _theDate.ToString()
        MainForm.txtDateIntVal.Text = _intVal.ToString()
        MainForm.cbUseCalDate.Checked = _useFullDate
        MainForm.txtNote.Text = ""
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase._command
        Return lineText
    End Function

    Public Function DateToUTC(locDate As String) As String
        ' Take the given string, make it into a DateTime, convert that to UTC,
        ' return a string for the result
        Dim utcDate As String
        Dim curTimeZone As TimeZone = TimeZone.CurrentTimeZone
        Dim loc As DateTime
        If (Not DateTime.TryParse(locDate, CultureInfo.CurrentCulture, DateTimeStyles.None, loc)) Then
            loc = curTimeZone.ToLocalTime(DateTime.Now)       ' if string is bad, use current datetime?
            MyBase.Warning = True
        End If
        If (_useFullDate) Then
            utcDate = curTimeZone.ToUniversalTime(loc).ToString()
        Else
            utcDate = curTimeZone.ToUniversalTime(loc).ToShortTimeString
        End If
        Return utcDate
    End Function

    Public Function UTCToLocal(utcDate As String) As String
        ' Take the given string, make it into a DateTime, convert that to UTC,
        ' return a string for the result
        Dim locDate As String
        Dim curTimeZone As TimeZone = TimeZone.CurrentTimeZone
        Dim utc As DateTime
        If (Not DateTime.TryParse(utcDate, CultureInfo.CurrentCulture, DateTimeStyles.None, utc)) Then
            utc = curTimeZone.ToUniversalTime(DateTime.Now)       ' if string is bad, use current datetime?
        End If
        If (_useFullDate) Then
            locDate = curTimeZone.ToLocalTime(utc).ToString()
        Else
            locDate = curTimeZone.ToLocalTime(utc).ToShortTimeString
        End If
        Return locDate
    End Function


#Region "Field Control Events"
    Private WithEvents myTxtDateIntVal As TextBox = MainForm.txtDateIntVal
    Private WithEvents myDtTimeCmd As DateTimePicker = MainForm.dtTimeCmd
    Private WithEvents myUseCalDate As CheckBox = MainForm.cbUseCalDate

    Private Sub txtIntVal_KeyPress(sender As Object, e As KeyPressEventArgs) Handles myTxtDateIntVal.KeyPress
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            If (e.KeyChar = vbCr) Then
                txtIntVal_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub txtIntVal_Enter(sender As Object, e As EventArgs) Handles myTxtDateIntVal.Enter
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            _origIntVal = myTxtDateIntVal.Text
        End If
    End Sub

    Private Sub txtIntVal_Leave(sender As Object, e As EventArgs) Handles myTxtDateIntVal.Leave
        Try
            If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
                If (UpdateIntVal(Trim(myTxtDateIntVal.Text))) Then
                    MainForm.GetPlan().Update()       ' update the plan
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Data Error")
        End Try
    End Sub

    Private Sub dtDateCmd_ValueChanged(sender As Object, e As EventArgs) Handles myDtTimeCmd.ValueChanged
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            Try
                If (UpdateDateVal(Trim(myDtTimeCmd.Text))) Then
                    MainForm.GetPlan().Update()       ' update the plan
                End If
            Catch ex As Exception
                MsgBox(ex.Message, MsgBoxStyle.Critical, "Data Error")
            End Try
        End If


    End Sub

    Private Sub cbUseCalDate_CheckedChanged(sender As Object, e As EventArgs) Handles myUseCalDate.CheckedChanged
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            _useFullDate = myUseCalDate.Checked
            If (_useFullDate) Then
                myDtTimeCmd.CustomFormat = FULL_DATE_FORMAT
            Else
                myDtTimeCmd.CustomFormat = TIME_ONLY_FORMAT
            End If
            MainForm.GetPlan().Update()
        End If
    End Sub

#End Region



End Class

#Region "CWaitUntil"
'=============================================================================================================
Public Class CWaitUntil : Inherits CDateCmd
    Public Sub New(cmd As String)
        ' must be format #cmd int date time   ' optional comment here
        MyBase.New(cmd)
        MyBase._minIntVal = 0
        MyBase._maxIntVal = 10     ' this would be # of sets
        MyBase._intFieldName = "Set Number"

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#waituntil") Then
            Dim cmdNoComment As String = cmd.Substring(cmdname.Length + 1)
            Dim semiIdx As Integer = cmdNoComment.IndexOf(";")
            If (semiIdx > 0) Then
                cmdNoComment = cmdNoComment.Remove(semiIdx)
            End If
            Dim args() As String = cmdNoComment.Split(",")
            If (args.Count <> 2) Then
                Throw New Exception("Invalid #waituntil command arguments {" & cmd & "}")
            End If
            UpdateIntVal(args(0))
            Dim localDate As String = UTCToLocal(args(1))
            UpdateDateVal(localDate)
        Else
            Throw New Exception("Invalid #waituntil command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtDateIntVal.Visible = True
        MainForm.lblDateIntVal.Visible = True
        MainForm.lblDateIntVal.Text = "Set Number"

        ' put tooltip on IntVal. DateCmd has tooltip set in design
        MainForm.ToolTip1.SetToolTip(MainForm.txtDateIntVal, "Set 0 waits on all sets. " & vbCrLf & "Set 1 only waits if in Set 1; restarting a plan which has completed the first of multiple sets will not wait.")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        ' TODO: the date needs to be adjusted to UTC

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & _intVal.ToString() & "," & DateToUTC(_theDate) & "      ; local datetime is " & _theDate
        'If (MyBase._comment <> "") Then
        '    lineText = lineText & "    ; " & MyBase._comment
        'End If
        Return lineText
    End Function

End Class
#End Region


#Region "CQuitAt"
'=============================================================================================================
Public Class CQuitAt : Inherits CDateCmd
    Public Sub New(cmd As String)
        ' must be format #cmd int date time   ' optional comment here
        MyBase.New(cmd)
        
        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#quitat") Then
            Dim cmdNoComment As String = cmd.Substring(cmdname.Length + 1)
            Dim semiIdx As Integer = cmdNoComment.IndexOf(";")
            If (semiIdx > 0) Then
                cmdNoComment = cmdNoComment.Remove(semiIdx)
            End If
            Dim args() As String = cmdNoComment.Split(",")
            If (args.Count <> 1) Then
                Throw New Exception("Invalid #QuitAt command arguments {" & cmd & "}")
            End If
            Dim localDate As String = UTCToLocal(args(0))
            UpdateDateVal(localDate)
        Else
            Throw New Exception("Invalid #QuitAt command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtDateIntVal.Visible = False
        MainForm.lblDateIntVal.Visible = False
        
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        
        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & DateToUTC(_theDate) & "      ; local datetime is " & _theDate
        'If (MyBase._comment <> "") Then
        '    lineText = lineText & "    ; " & MyBase._comment
        'End If
        Return lineText
    End Function

End Class
#End Region

#Region "CShutdownAt"
'=============================================================================================================
Public Class CShutdownAt : Inherits CDateCmd
    Public Sub New(cmd As String)
        ' must be format #cmd int date time   ' optional comment here
        MyBase.New(cmd)

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        Dim cmdname As String = LCase(Trim(pieces(0)))
        If (cmdname = "#shutdownat") Then
            Dim cmdNoComment As String = cmd.Substring(cmdname.Length + 1)
            Dim semiIdx As Integer = cmdNoComment.IndexOf(";")
            If (semiIdx > 0) Then
                cmdNoComment = cmdNoComment.Remove(semiIdx)
            End If
            Dim args() As String = cmdNoComment.Split(",")
            If (args.Count <> 1) Then
                Throw New Exception("Invalid #shutdownat command arguments {" & cmd & "}")
            End If
            Dim localDate As String = UTCToLocal(args(0))
            UpdateDateVal(localDate)
        Else
            Throw New Exception("Invalid #shutdownat command {" & cmd & "}")
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtDateIntVal.Visible = False
        MainForm.lblDateIntVal.Visible = False

    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & DateToUTC(_theDate) & "      ; local datetime is " & _theDate
        'If (MyBase._comment <> "") Then
        '    lineText = lineText & "    ; " & MyBase._comment
        'End If
        Return lineText
    End Function

End Class
#End Region
