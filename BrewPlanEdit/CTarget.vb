Public Class CTarget : Inherits CommandBaseClass
    Private _TargetName As String
    Private _TargetRA As String
    Private _TargetDec As String

    Public Property TargetName() As String
        Get
            Return _TargetName
        End Get
        Set(value As String)
            _TargetName = value
        End Set
    End Property

    Public Property TargetRA() As String
        Get
            Return _TargetRA
        End Get
        Set(value As String)
            _TargetRA = value
        End Set
    End Property

    Public Property TargetDec() As String
        Get
            Return _TargetDec
        End Get
        Set(value As String)
            _TargetDec = value
        End Set
    End Property


    Public Function UpdateTargetName(curText As String) As Boolean
        ' if the string has changed, update the  value
        ' returns True if update occurred
        Dim updated As Boolean = False
        If (curText <> _origTargetName) Then
            If (curText = "") Then
                ' must have a name
                Throw New Exception("Target Name cannot be blank")
            End If
            _TargetName = Trim(curText)
            _origTargetName = Trim(curText)
            updated = True
        End If
        Return updated
    End Function

    Public Function UpdateTargetRA(curText As String) As Boolean
        ' if the string has changed, update the  value
        ' returns True if update occurred
        Dim updated As Boolean = False
        If (curText <> _origTargetRA) Then
            Dim errMsg As String = ""
            If (curText <> "") Then
                errMsg = TargetAddressIsOK(curText, True)
            End If
            If (errMsg = "") Then
                _TargetRA = curText
                _origTargetRA = curText
                updated = True
            Else
                Throw New Exception("Invalid Target RA value {" & curText & "}" & vbCrLf & errMsg)
            End If
        End If
        Return updated
    End Function

    Public Function UpdateTargetDec(curText As String) As Boolean
        ' if the string has changed, update the  value
        ' returns True if update occurred
        Dim updated As Boolean = False
        If (curText <> _origTargetDec) Then
            Dim errMsg As String = ""
            If (curText <> "") Then
                errMsg = TargetAddressIsOK(curText, False)
            End If
            If (errMsg = "") Then
                ' check for both filled in or both blank
                If (((curText = "") And (_TargetRA = "")) Or ((curText <> "") And (_TargetRA <> ""))) Then
                    'ok
                Else
                    errMsg = "Must have both RA and Dec filled in, or both blank"
                End If
            End If
            If (errMsg = "") Then
                _TargetDec = curText
                _origTargetDec = curText
                updated = True
            Else
                Throw New Exception("Invalid Target Dec value {" & curText & "}" & vbCrLf & errMsg)
            End If
        End If
        Return updated
    End Function

    Private Function TargetAddressIsOK(val As String, RAformat As Boolean) As String
        ' return true if OK
        ' needs to be format 99 99 99.99
        Dim pieces() As String = val.Split(" ")
        Dim i As Integer = 0
        Dim d As Double
        
        If (pieces.Count <> 3) Then
            Return "Invalid Target Address value {" & val & "}" & vbCrLf & "Needs 3 sub fields"
        End If
        If (Not Integer.TryParse(pieces(0), i)) Then
            Return "Invalid Target Address degrees {" & pieces(0) & "}"
        End If
        If (RAformat) Then
            If ((i < 0) Or (i > 24)) Then
                Return "Invalid Target RA Address degrees {" & i & "}"
            End If
        Else
            If ((i < -90) Or (i > 90)) Then
                Return "Invalid Target Dec Address degrees {" & i & "}"
            End If
        End If
        If ((Not Integer.TryParse(pieces(1), i)) Or (i < 0) Or (i > 59)) Then
            Return "Invalid Target Address minutes {" & pieces(1) & "}"
        End If
        If ((Not Double.TryParse(pieces(2), d)) Or (d < 0) Or (d > 59.9)) Then
            Return "Invalid Target Address seconds {" & pieces(2) & "}"
        End If
        Return ""
    End Function

    Public Sub New(cmd As String)
        ' must be format targetname \tRA \tDec   ' optional comment here
        MyBase.New()
        Dim pieces() As String = MyBase.CleanCommandString(cmd)
        ' for target, pieces have been split by Tab
        If (pieces.Count = 1) Then
            _TargetName = pieces(0)
            _TargetRA = ""
            _TargetDec = ""
        ElseIf (pieces.Count = 3) Then
            _TargetName = pieces(0)
            _TargetRA = pieces(1)
            _TargetDec = pieces(2)
        Else
            Throw New Exception("Invalid target command " & cmd)
        End If
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.pnlTarget.Visible = True
        MainForm.txtTargetName.Text = _TargetName
        MainForm.txtTargetRA.Text = _TargetRA
        MainForm.txtTargetDec.Text = _TargetDec
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        lineText = _TargetName
        If (_TargetRA <> "") Then
            lineText = lineText & vbTab & _TargetRA & " " & vbTab & _TargetDec
        End If
        If (MyBase._comment <> "") Then
            lineText = lineText & "    ; " & MyBase._comment
        End If
        Return lineText
    End Function



#Region "Field Control Events"
    Private WithEvents myTxtTargetName As TextBox = MainForm.txtTargetName
    Private WithEvents myTxtTargetRA As TextBox = MainForm.txtTargetRA
    Private WithEvents myTxtTargetDec As TextBox = MainForm.txtTargetDec
    Private _origTargetName As String
    Private _origTargetRA As String
    Private _origTargetDec As String
    Private Sub txtTargetName_KeyPress(sender As Object, e As KeyPressEventArgs) Handles myTxtTargetName.KeyPress
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            If (e.KeyChar = vbCr) Then
                txtTargetName_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub txtTargetName_Enter(sender As Object, e As EventArgs) Handles myTxtTargetName.Enter
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            _origTargetName = myTxtTargetName.Text
        End If
    End Sub

    Private Sub txtTargetName_Leave(sender As Object, e As EventArgs) Handles myTxtTargetName.Leave
        Try
            If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
                If (UpdateTargetName(Trim(myTxtTargetName.Text))) Then
                    MainForm.GetPlan().Update()       ' update the plan
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Data Error")
        End Try
    End Sub

    Private Sub txtTargetRA_KeyPress(sender As Object, e As KeyPressEventArgs) Handles myTxtTargetRA.KeyPress
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            If (e.KeyChar = vbCr) Then
                txtTargetRA_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub txtTargetRA_Enter(sender As Object, e As EventArgs) Handles myTxtTargetRA.Enter
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            _origTargetRA = myTxtTargetRA.Text
        End If
    End Sub

    Private Sub txtTargetRA_Leave(sender As Object, e As EventArgs) Handles myTxtTargetRA.Leave
        Try
            If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
                If (UpdateTargetRA(Trim(myTxtTargetRA.Text))) Then
                    MainForm.GetPlan().Update()       ' update the plan
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Data Error")
        End Try
    End Sub

    Private Sub txtTargetDec_KeyPress(sender As Object, e As KeyPressEventArgs) Handles myTxtTargetDec.KeyPress
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            If (e.KeyChar = vbCr) Then
                txtTargetDec_Leave(sender, e)
            End If
        End If
    End Sub

    Private Sub txtTargetDec_Enter(sender As Object, e As EventArgs) Handles myTxtTargetDec.Enter
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            _origTargetDec = myTxtTargetDec.Text
        End If
    End Sub

    Private Sub txtTargetDec_Leave(sender As Object, e As EventArgs) Handles myTxtTargetDec.Leave
        Try
            If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
                If (UpdateTargetDec(Trim(myTxtTargetDec.Text))) Then
                    MainForm.GetPlan().Update()       ' update the plan
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Data Error")
        End Try
    End Sub
#End Region


End Class
