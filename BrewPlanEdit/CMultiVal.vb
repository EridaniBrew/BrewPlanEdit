Public Class CMultiVal : Inherits CommandBaseClass
    ' this class will be further subclassed for specific commands:
    ' allows single implementation of the commands using various integer and double text controls
    Protected Const NUMFIELDS = 3       ' num of integer fields, num of double fields

    Public Const FIELD_INT1 = 0       ' index of first Int field
    Public Const FIELD_INT2 = 1       ' index of second Int field
    Public Const FIELD_INT3 = 2       ' index of second Int field
    Public Const FIELD_DBL1 = 3       ' index of first Dbl field
    Public Const FIELD_DBL2 = 4       ' index of second Dbl field
    Public Const FIELD_DBL3 = 5       ' index of second Dbl field

    Protected _intVal(NUMFIELDS - 1) As Integer
    Protected _dblVal(NUMFIELDS - 1) As Double
    Protected _origIntVal(NUMFIELDS - 1) As String
    Protected _origDblVal(NUMFIELDS - 1) As String
    Protected _minIntVal(NUMFIELDS - 1) As Integer          ' subclass needs to set these
    Protected _maxIntVal(NUMFIELDS - 1) As Integer
    Protected _minDblVal(NUMFIELDS - 1) As Double          ' subclass needs to set these
    Protected _maxDblVal(NUMFIELDS - 1) As Double

    Protected _usingInt(NUMFIELDS - 1) As Boolean
    Protected _usingDbl(NUMFIELDS - 1) As Boolean

    Protected _txtMVInt(NUMFIELDS - 1) As TextBox
    Protected _txtMVDbl(NUMFIELDS - 1) As TextBox
    Protected _lblMVInt(NUMFIELDS - 1) As Label
    Protected _lblMVDbl(NUMFIELDS - 1) As Label

    Protected _intFieldName(NUMFIELDS - 1) As String           ' field name for error messages
    Protected _dblFieldName(NUMFIELDS - 1) As String           ' field name for error messages

    Protected _blankAllowedInt(NUMFIELDS - 1) As Boolean        ' if blank is OK for field (optional field)
    Protected _blankAllowedDbl(NUMFIELDS - 1) As Boolean
    Protected _blankValueInt(NUMFIELDS - 1) As Integer          ' value when blank, indicating blank
    Protected _blankValueDbl(NUMFIELDS - 1) As Double          ' value when blank, indicating blank

    Public Sub New()
        ' must be format #cmd int date time   ' optional comment here
        MyBase.New()

        ' init control arrays
        _txtMVInt(0) = myTxtMVInt1
        _txtMVInt(1) = myTxtMVInt2
        _txtMVInt(2) = myTxtMVInt3
        _txtMVDbl(0) = myTxtMVDbl1
        _txtMVDbl(1) = myTxtMVDbl2
        _txtMVDbl(2) = myTxtMVDbl3
        _lblMVInt(0) = MainForm.lblMVInt1
        _lblMVInt(1) = MainForm.lblMVInt2
        _lblMVInt(2) = MainForm.lblMVInt3
        _lblMVDbl(0) = MainForm.lblMVDbl1
        _lblMVDbl(1) = MainForm.lblMVDbl2
        _lblMVDbl(2) = MainForm.lblMVDbl3

        Dim i As Integer
        For i = 0 To NUMFIELDS - 1
            _minIntVal(i) = 0
            _maxIntVal(i) = 1
            _usingInt(i) = False
            _blankAllowedInt(i) = False
            _blankValueInt(i) = 0

            _minDblVal(i) = 0
            _maxDblVal(i) = 1
            _usingDbl(i) = False
            _blankAllowedDbl(i) = False
            _blankValueDbl(i) = -1

        Next
    End Sub

    Public Sub SelectFields(index As Integer, useField As Boolean, minVal As Double, maxVal As Double)
        ' Configure the fields for use
        If ((index = FIELD_INT1) Or (index = FIELD_INT2) Or (index = FIELD_INT3)) Then
            _minIntVal(index) = minVal
            _maxIntVal(index) = maxVal
            _usingInt(index) = useField
        ElseIf ((index = FIELD_DBL1) Or (index = FIELD_DBL2) Or (index = FIELD_DBL3)) Then
            index = index - NUMFIELDS
            _minDblVal(index) = minVal
            _maxDblVal(index) = maxVal
            _usingDbl(index) = useField
        End If
    End Sub

    Public Function UpdateIntVal(idx As Integer, curText As String) As Boolean
        ' if the string has changed, update the value
        ' returns True if update occurred
        Dim updated As Boolean = False
        Dim tempVal As Integer

        If (curText <> _origIntVal(idx)) Then
            Dim errMsg As String = ""
            If (curText = "") Then
                If (_blankAllowedInt(idx)) Then
                    _intVal(idx) = _blankValueInt(idx)
                    _origIntVal(idx) = curText
                    Return True
                Else
                    Throw New Exception("Blank field not allowed " & _intFieldName(idx))        ' field is required
                End If
            End If
            If (Integer.TryParse(curText, tempVal)) Then
                If ((tempVal < _minIntVal(idx)) Or (tempVal > _maxIntVal(idx))) Then
                    Throw New Exception("Field " & _intFieldName(idx) & " out of range {" & curText & "}" & vbCrLf & "Allowed range is " & _minIntVal(idx) & ".." & _maxIntVal(idx))
                Else
                    _intVal(idx) = tempVal
                    _origIntVal(idx) = curText
                    updated = True
                End If
            Else
                Throw New Exception("Invalid Field " & _intFieldName(idx) & " value {" & curText & "}" & vbCrLf & errMsg)
            End If
        End If
        Return updated
    End Function

    Public Function UpdateDblVal(idx As Integer, curText As String) As Boolean
        ' if the string has changed, update the value
        ' returns True if update occurred
        Dim updated As Boolean = False
        Dim tempVal As Double

        If (curText <> _origDblVal(idx)) Then
            Dim errMsg As String = ""
            If (curText = "") Then
                If (_blankAllowedDbl(idx)) Then
                    _dblVal(idx) = _blankValueDbl(idx)
                    _origDblVal(idx) = curText
                    Return True
                Else
                    Throw New Exception("Blank field " & _dblFieldName(idx) & " not allowed " & _dblFieldName(idx))        ' field is required
                End If
            End If
            If (Double.TryParse(curText, tempVal)) Then
                If ((tempVal < _minDblVal(idx)) Or (tempVal > _maxDblVal(idx))) Then
                    Throw New Exception("Field " & _dblFieldName(idx) & " out of range {" & curText & "}" & vbCrLf & "Allowed range is " & _minDblVal(idx) & ".." & _maxDblVal(idx))
                Else
                    _dblVal(idx) = tempVal
                    _origDblVal(idx) = curText
                    updated = True
                End If
            Else
                Throw New Exception("Invalid Field " & _dblFieldName(idx) & " value {" & curText & "}" & vbCrLf & errMsg)
            End If
        End If
        Return updated
    End Function


    Public Overrides Sub Display()
        Dim i As Integer

        MyBase.Display()
        MainForm.pnlMultiVal.Visible = True
        For i = 0 To NUMFIELDS - 1
            _txtMVInt(i).Visible = _usingInt(i)
            _lblMVInt(i).Visible = _usingInt(i)
            _txtMVDbl(i).Visible = _usingDbl(i)
            _lblMVDbl(i).Visible = _usingDbl(i)
            If (_blankAllowedInt(i) And (_blankValueInt(i) = _intVal(i))) Then
                _txtMVInt(i).Text = ""
            Else
                _txtMVInt(i).Text = _intVal(i).ToString()
            End If
            If (_blankAllowedDbl(i) And (_blankValueDbl(i) = _dblVal(i))) Then
                _txtMVDbl(i).Text = ""
            Else
                _txtMVDbl(i).Text = _dblVal(i).ToString()
            End If
            
        Next

    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = MyBase._command
        Return lineText
    End Function



#Region "Field Control Events"
    Private WithEvents myTxtMVInt1 As TextBox = MainForm.txtMVInt1
    Private WithEvents myTxtMVInt2 As TextBox = MainForm.txtMVInt2
    Private WithEvents myTxtMVInt3 As TextBox = MainForm.txtMVInt3
    Private WithEvents myTxtMVDbl1 As TextBox = MainForm.txtMVDbl1
    Private WithEvents myTxtMVDbl2 As TextBox = MainForm.txtMVDbl2
    Private WithEvents myTxtMVDbl3 As TextBox = MainForm.txtMVDbl3

    Private Sub txtIntVal_KeyPress(sender As Object, e As KeyPressEventArgs) Handles myTxtMVInt1.KeyPress, myTxtMVInt2.KeyPress, myTxtMVInt3.KeyPress
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            If (e.KeyChar = vbCr) Then
                txtIntVal_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub txtIntVal_Enter(sender As Object, e As EventArgs) Handles myTxtMVInt1.Enter, myTxtMVInt2.Enter, myTxtMVInt3.Enter
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            Dim idx As Integer = FIELD_INT1
            Dim s As String = myTxtMVInt1.Text
            If (sender.name = "txtMVInt2") Then
                idx = FIELD_INT2
                s = myTxtMVInt2.Text
            ElseIf (sender.name = "txtMVInt3") Then
                idx = FIELD_INT3
                s = myTxtMVInt3.Text
            End If
            _origIntVal(idx) = Trim(s)
        End If
    End Sub

    Private Sub txtIntVal_Leave(sender As Object, e As EventArgs) Handles myTxtMVInt1.Leave, myTxtMVInt2.Leave, myTxtMVInt3.Leave
        Dim idx As Integer = FIELD_INT1
        Dim s As String = myTxtMVInt1.Text
        If (sender.name = "txtMVInt2") Then
            idx = FIELD_INT2
            s = myTxtMVInt2.Text
        ElseIf (sender.name = "txtMVInt3") Then
            idx = FIELD_INT3
            s = myTxtMVInt3.Text
        End If
        Try
            If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
                If (UpdateIntVal(idx, Trim(s))) Then
                    MainForm.GetPlan().Update()       ' update the plan
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Data Error")
        End Try
    End Sub

    Private Sub txtDblVal_KeyPress(sender As Object, e As KeyPressEventArgs) Handles myTxtMVDbl1.KeyPress, myTxtMVDbl2.KeyPress, myTxtMVDbl3.KeyPress
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            If (e.KeyChar = vbCr) Then
                txtDblVal_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub txtDblVal_Enter(sender As Object, e As EventArgs) Handles myTxtMVDbl1.Enter, myTxtMVDbl2.Enter, myTxtMVDbl3.Enter
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            Dim idx As Integer = FIELD_DBL1 - NUMFIELDS
            Dim s As String = myTxtMVDbl1.Text
            If (sender.name = "txtMVDbl2") Then
                idx = FIELD_DBL2 - NUMFIELDS
                s = myTxtMVDbl2.Text
            ElseIf (sender.name = "txtMVDbl3") Then
                idx = FIELD_DBL3 - NUMFIELDS
                s = myTxtMVDbl3.Text
            End If
            _origDblVal(idx) = Trim(s)
        End If
    End Sub

    Private Sub txtDblVal_Leave(sender As Object, e As EventArgs) Handles myTxtMVDbl1.Leave, myTxtMVDbl2.Leave, myTxtMVDbl3.Leave
        Dim idx As Integer = FIELD_DBL1 - NUMFIELDS
        Dim s As String = myTxtMVDbl1.Text
        If (sender.name = "txtMVDbl2") Then
            idx = FIELD_DBL2 - NUMFIELDS
            s = myTxtMVDbl2.Text
        ElseIf (sender.name = "txtMVDbl3") Then
            idx = FIELD_DBL3 - NUMFIELDS
            s = myTxtMVDbl3.Text
        End If
        Try
            If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
                If (UpdateDblVal(idx, Trim(s))) Then
                    MainForm.GetPlan().Update()       ' update the plan
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Data Error")
        End Try
    End Sub

#End Region
End Class


#Region "CSubFrame"
'=============================================================================================================
Public Class CSubFrame : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #subframe percent   ' optional comment here
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_DBL1, True, 0.1, 1.0)
        _dblFieldName(0) = "Subframe"

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If ((pieces.Length = 2) And (LCase(Trim(pieces(0))) = "#subframe")) Then
            UpdateDblVal(0, Trim(pieces(1)))
        Else
            Throw New Exception("Invalid subframe command " & cmd)
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVDbl1.Visible = True
        MainForm.lblMVDbl1.Visible = True
        MainForm.lblMVDbl1.Text = "SubFrame Percentage"

        ' put tooltip on IntVal
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVDbl1, "Select the fraction of the chip used for the subframe.")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & MyBase._dblVal(0).ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CChill"
'=============================================================================================================
Public Class CChill : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #Chill temp, precision   ' optional comment here
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_DBL1, True, -50, 50.0)
        _dblFieldName(0) = "Chill Temperature"
        MyBase.SelectFields(CMultiVal.FIELD_DBL2, True, 0.01, 3.0)
        _dblFieldName(1) = "Precision"
        _blankAllowedDbl(1) = True
        _blankValueDbl(1) = -1

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If ((pieces.Length = 2) And (LCase(Trim(pieces(0))) = "#chill")) Then
            Dim data() As String = pieces(1).Split(",")
            UpdateDblVal(0, Trim(data(0)))
            If (data.Length > 1) Then
                ' optional second value
                UpdateDblVal(1, Trim(data(1)))
            Else
                _dblVal(1) = _blankValueDbl(1)
                _origDblVal(1) = ""
            End If
        Else
            Throw New Exception("Invalid chill command " & cmd)
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVDbl1.Visible = True
        MainForm.lblMVDbl1.Visible = True
        MainForm.lblMVDbl1.Text = "Chill Temperature"
        MainForm.txtMVDbl2.Visible = True
        MainForm.lblMVDbl2.Visible = True
        MainForm.lblMVDbl2.Text = "Precision"

        ' put tooltip on IntVal
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVDbl1, "Target temperature for cooling the CCD.")
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVDbl2, "(Optional) Required precision for target temperature.")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & MyBase._dblVal(0).ToString()
        If (_dblVal(1) <> _blankValueDbl(1)) Then
            lineText = lineText & "," & MyBase._dblVal(1).ToString()
        End If
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region


#Region "CWaitFor"
'=============================================================================================================
Public Class CWaitFor : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #WaitFor seconds   ' optional comment here
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_INT1, True, 1, 10000)
        MyBase.SelectFields(CMultiVal.FIELD_INT2, False, 0, 1)
        MyBase.SelectFields(CMultiVal.FIELD_DBL1, False, 0.0, 1)
        MyBase.SelectFields(CMultiVal.FIELD_DBL2, False, 0, 1)
        _intFieldName(0) = "WaitFor"

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If ((pieces.Length = 2) And (LCase(Trim(pieces(0))) = "#waitfor")) Then
            UpdateIntVal(0, Trim(pieces(1)))
        Else
            Throw New Exception("Invalid waitfor command " & cmd)
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVInt1.Visible = True
        MainForm.lblMVInt1.Visible = True
        MainForm.lblMVInt1.Text = "Wait For (Seconds)"

        ' put tooltip on IntVal
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVInt1, "Select the number of seconds to wait.")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & MyBase._intVal(0).ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CPosAng"
'=============================================================================================================
Public Class CPosAng : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #posang angle   ' optional comment here
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_DBL1, True, 0.0, 360.0)
        _dblFieldName(0) = "Position Angle"

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If ((pieces.Length = 2) And (LCase(Trim(pieces(0))) = "#posang")) Then
            UpdateDblVal(0, Trim(pieces(1)))
        Else
            Throw New Exception("Invalid posang command " & cmd)
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVDbl1.Visible = True
        MainForm.lblMVDbl1.Visible = True
        MainForm.lblMVDbl1.Text = "Position Angle"

        ' put tooltip on IntVal
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVDbl1, "Select the position angle for the Rotator.")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & MyBase._dblVal(0).ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CDither"
'=============================================================================================================
Public Class CDither : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #CDither pixels   ' optional comment here
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_DBL1, True, 0.0, 20.0)
        _dblFieldName(0) = "Dither"
        _blankAllowedDbl(0) = True
        _blankValueDbl(0) = -1

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If (LCase(Trim(pieces(0))) = "#dither") Then
            If (pieces.Length = 2) Then
                UpdateDblVal(0, Trim(pieces(1)))
            Else
                _dblVal(0) = _blankValueDbl(0)
                _origDblVal(0) = ""
            End If
        Else
            Throw New Exception("Invalid dither command " & cmd)
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVDbl1.Visible = True
        MainForm.lblMVDbl1.Visible = True
        MainForm.lblMVDbl1.Text = "Dither"

        ' put tooltip on IntVal
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVDbl1, "Select the amountof dithering, pixels or arcseconds (See Help).")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString()
        If (_dblVal(0) <> _blankValueDbl(0)) Then
            lineText = lineText & " " & MyBase._dblVal(0).ToString()
        End If
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CWaitInLimits"
'=============================================================================================================
Public Class CWaitInLimits : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #WaitInLimits minutes   ' optional comment here
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_INT1, True, 1, 10000)
        _intFieldName(0) = "WaitInLimits"

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If ((pieces.Length = 2) And (LCase(Trim(pieces(0))) = "#waitinlimits")) Then
            UpdateIntVal(0, Trim(pieces(1)))
        Else
            Throw New Exception("Invalid waitinlimits command " & cmd)
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVInt1.Visible = True
        MainForm.lblMVInt1.Visible = True
        MainForm.lblMVInt1.Text = "Wait this many minutes for observatory to be within limits"

        ' put tooltip on IntVal
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVInt1, "Select the number of minutes to wait.")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & MyBase._intVal(0).ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CDefocus"
'=============================================================================================================
Public Class CDefocus : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #WaitFor seconds   ' optional comment here
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_INT1, True, -1000, 1000)
        _intFieldName(0) = "Defocus"

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If ((pieces.Length = 2) And (LCase(Trim(pieces(0))) = "#defocus")) Then
            UpdateIntVal(0, Trim(pieces(1)))
        Else
            Throw New Exception("Invalid defocus command " & cmd)
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVInt1.Visible = True
        MainForm.lblMVInt1.Visible = True
        MainForm.lblMVInt1.Text = "Shift focuser this many tics"

        ' put tooltip on IntVal
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVInt1, "Select the number of tics to jog focuser.")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & MyBase._intVal(0).ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CSets"
'=============================================================================================================
Public Class CSets : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #Sets num   ' optional comment here
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_INT1, True, 1, 100)
        _intFieldName(0) = "Sets"

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If ((pieces.Length = 2) And (LCase(Trim(pieces(0))) = "#sets")) Then
            UpdateIntVal(0, Trim(pieces(1)))
        Else
            Throw New Exception("Invalid sets command " & cmd)
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVInt1.Visible = True
        MainForm.lblMVInt1.Visible = True
        MainForm.lblMVInt1.Text = "Number of sets"

        ' put tooltip on IntVal
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVInt1, "Select the number of Sets for the plan.")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & MyBase._intVal(0).ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CRepeat"
'=============================================================================================================
Public Class CRepeat : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #Repeat num   ' optional comment here
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_INT1, True, 1, 100)
        _intFieldName(0) = "Repeat"

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If ((pieces.Length = 2) And (LCase(Trim(pieces(0))) = "#repeat")) Then
            UpdateIntVal(0, Trim(pieces(1)))
        Else
            Throw New Exception("Invalid repeat command " & cmd)
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVInt1.Visible = True
        MainForm.lblMVInt1.Visible = True
        MainForm.lblMVInt1.Text = "Repeat count"

        ' put tooltip on IntVal
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVInt1, "The number of times a filter set is performed.")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & MyBase._intVal(0).ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CAfInterval"
'=============================================================================================================
Public Class CAfInterval : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #AfInterval seconds   ' optional comment here
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_INT1, True, 1, 1000)
        _intFieldName(0) = "AfInterval"

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If ((pieces.Length = 2) And (LCase(Trim(pieces(0))) = "#afinterval")) Then
            UpdateIntVal(0, Trim(pieces(1)))
        Else
            Throw New Exception("Invalid afinterval command " & cmd)
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVInt1.Visible = True
        MainForm.lblMVInt1.Visible = True
        MainForm.lblMVInt1.Text = "Autofocus Interval (minutes)"

        ' put tooltip on IntVal
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVInt1, "The number of minutes between focusing runs")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & MyBase._intVal(0).ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CStartSetNum"
'=============================================================================================================
Public Class CStartSetNum : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #WaitFor seconds   ' optional comment here
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_INT1, True, 1, 100)
        _intFieldName(0) = "StartSetNum"

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If ((pieces.Length = 2) And (LCase(Trim(pieces(0))) = "#startsetnum")) Then
            UpdateIntVal(0, Trim(pieces(1)))
        Else
            Throw New Exception("Invalid startsetnum command " & cmd)
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVInt1.Visible = True
        MainForm.lblMVInt1.Visible = True
        MainForm.lblMVInt1.Text = "Starting Set Number" & vbCrLf & "See Help - use caution changing this!"

        ' put tooltip on IntVal
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVInt1, "Starting set number for this plan restart")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & MyBase._intVal(0).ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CWaitZenDist"
'=============================================================================================================
Public Class CWaitZenDist : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #WaitFor seconds   ' optional comment here
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_INT1, True, 0, 90)
        MyBase.SelectFields(CMultiVal.FIELD_INT2, True, 1, 120)
        _intFieldName(0) = "ZenithDegrees"
        _intFieldName(1) = "WaitMinutes"

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If (LCase(Trim(pieces(0))) = "#waitzendist") Then
            Dim data As String = cmd.Substring(13)
            Dim fields() As String = data.Split(",")
            If (fields.Length = 2) Then
                UpdateIntVal(0, Trim(fields(0)))
                UpdateIntVal(1, Trim(fields(1)))
            Else
                Throw New Exception("Invalid waitzendist command: needs 2 data fields {" & cmd & "}")
            End If
        Else
            Throw New Exception("Invalid waitzendist command " & cmd)
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVInt1.Visible = True
        MainForm.lblMVInt1.Visible = True
        MainForm.lblMVInt1.Text = "Zenith Target Degrees"

        MainForm.txtMVInt2.Visible = True
        MainForm.lblMVInt2.Visible = True
        MainForm.lblMVInt2.Text = "Minutes to wait"

        ' put tooltip on IntVal
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVInt1, "Target degrees of zenith")
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVInt2, "Number of minutes to wait")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & MyBase._intVal(0).ToString() & "," & MyBase._intVal(1).ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CWaitAirMass"
'=============================================================================================================
Public Class CWaitAirMass : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #WaitFor seconds   ' optional comment here
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_INT1, True, 1, 120)
        MyBase.SelectFields(CMultiVal.FIELD_DBL1, True, 1.0, 4.0)
        _intFieldName(0) = "WaitMinutes"
        _intFieldName(1) = "AirMass"

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If (LCase(Trim(pieces(0))) = "#waitairmass") Then
            Dim data As String = cmd.Substring(13)
            Dim fields() As String = data.Split(",")
            If (fields.Length = 2) Then
                UpdateDblVal(0, Trim(fields(0)))
                UpdateIntVal(0, Trim(fields(1)))
            Else
                Throw New Exception("Invalid waitairmass command: needs 2 data fields {" & cmd & "}")
            End If
        Else
            Throw New Exception("Invalid waitairmass command " & cmd)
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVInt1.Visible = True
        MainForm.lblMVInt1.Visible = True
        MainForm.lblMVInt1.Text = "Minutes to wait"

        MainForm.txtMVDbl1.Visible = True
        MainForm.lblMVDbl1.Visible = True
        MainForm.lblMVDbl1.Text = "Air Mass target"

        ' put tooltip on IntVal
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVInt1, "Minutes to wait until air mass reached")
        MainForm.ToolTip1.SetToolTip(MainForm.txtMVDbl1, "Number of Air masses to wait for")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & MyBase._dblVal(0).ToString() & "," & MyBase._intVal(0).ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region

#Region "CCompletionState"
'=============================================================================================================
Public Class CCompletionState : Inherits CMultiVal
    Public Sub New(cmd As String)
        ' must be format #completionstate 1,2,3,4,5   ' optional comment here
        ' this is not typically added by the user. ACP adds this command while the
        ' plan is being executed. If the plan is restarted, this directive allows ACP
        ' to continue where it left off
        MyBase.New()

        MyBase.SelectFields(CMultiVal.FIELD_INT1, True, 0, 20)
        _intFieldName(0) = "Sets completed"
        MyBase.SelectFields(CMultiVal.FIELD_INT2, True, 0, 20)
        _intFieldName(1) = "Targets in current Set"
        MyBase.SelectFields(CMultiVal.FIELD_INT3, True, 0, 20)
        _intFieldName(2) = "Repeats in current Target"
        MyBase.SelectFields(CMultiVal.FIELD_DBL1, True, 0, 4)
        _dblFieldName(0) = "Filter Groups in current Repeat"
        MyBase.SelectFields(CMultiVal.FIELD_DBL2, True, 0, 100)
        _dblFieldName(1) = "Images in Filter Group"

        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        If ((pieces.Length = 2) And (LCase(Trim(pieces(0))) = "#completionstate")) Then
            Dim data() As String = pieces(1).Split(",")
            If (data.Length = 5) Then
                UpdateIntVal(0, Trim(data(0)))
                UpdateIntVal(1, Trim(data(1)))
                UpdateIntVal(2, Trim(data(2)))
                UpdateDblVal(0, Trim(data(3)))
                UpdateDblVal(1, Trim(data(4)))
            Else
                Throw New Exception("Invalid completionstate command does not have 5 fields {" & cmd & "}")
            End If
        Else
            Throw New Exception("Invalid completionstate command {" & cmd & "}")
        End If

    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.txtMVInt1.Visible = True
        MainForm.lblMVInt1.Visible = True
        MainForm.lblMVInt1.Text = "Sets Completed"
        MainForm.txtMVInt2.Visible = True
        MainForm.lblMVInt2.Visible = True
        MainForm.lblMVInt2.Text = "Targets in current Set"
        MainForm.txtMVInt3.Visible = True
        MainForm.lblMVInt3.Visible = True
        MainForm.lblMVInt3.Text = "Repeats in current Target"
        MainForm.txtMVDbl1.Visible = True
        MainForm.lblMVDbl1.Visible = True
        MainForm.lblMVDbl1.Text = "Filter Groups in current Repeat"
        MainForm.txtMVDbl2.Visible = True
        MainForm.lblMVDbl2.Visible = True
        MainForm.lblMVDbl2.Text = "Images in Filter Group"

        ' put tooltip on IntVal
        'MainForm.ToolTip1.SetToolTip(MainForm.txtMVDbl1, "Select the fraction of the chip used for the subframe.")
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label

        Dim lineText As String = ""
        lineText = MyBase.ToString() & " " & MyBase._intVal(0).ToString() & "," _
            & MyBase._intVal(1).ToString() & "," _
            & MyBase._intVal(2).ToString() & "," _
            & MyBase._dblVal(0).ToString() & "," _
            & MyBase._dblVal(1).ToString()
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function

End Class
#End Region





