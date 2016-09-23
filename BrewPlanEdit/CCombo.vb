Public Class CCombo : Inherits CommandBaseClass
    Private _numCols As Integer    ' how many columns are active
    
    Public Const MAXCOLS = 8
    Private _count(MAXCOLS - 1) As Integer
    Private _interval(MAXCOLS - 1) As Double
    Private _bin(MAXCOLS - 1) As String
    Private _filter(MAXCOLS - 1) As String

    Private _txtCount As Collection
    Private _txtInterval As Collection
    Private _lstFilter As Collection
    Private _lstBin As Collection

    Private _useRow(MAXCOLS - 1) As Boolean
    Private Const USE_COUNT = 0
    Private Const USE_INTERVAL = 1
    Private Const USE_BIN = 2
    Private Const USE_FILTER = 3

    Private Function ValidInteger(s As String, fieldName As String) As String
        ' check text string: must be non-blank, valid integer, positive
        Dim retS As String = ""
        Dim ival As Integer = 0
        If (Trim(s) = "") Then
            retS = "Invalid " & fieldName & " (blank) in {" & s & "}" & vbCrLf
        ElseIf (Not Integer.TryParse(s, ival)) Then
            retS = "Invalid " & fieldName & " in {" & s & "}" & vbCrLf
        ElseIf (ival < 1) Then
            retS = "Invalid " & fieldName & " in {" & s & "}" & vbCrLf
        End If
        Return retS
    End Function

    Private Function ValidDouble(s As String, fieldName As String) As String
        ' check text string: must be non-blank, valid integer, positive
        Dim retS As String = ""
        Dim ival As Double = 0
        If (Trim(s) = "") Then
            retS = "Invalid " & fieldName & " (blank) in {" & s & "}" & vbCrLf
        ElseIf (Not Double.TryParse(s, ival)) Then
            retS = "Invalid " & fieldName & " in {" & s & "}" & vbCrLf
        ElseIf (ival < 0) Then
            retS = "Invalid " & fieldName & " in {" & s & "}" & vbCrLf
        End If
        Return retS
    End Function

    Private Function UpdateFields() As Boolean
        ' verify the fields are OK, then update the internal variables
        ' return True if OK,  false if error
        ' Count cannot be blank, negative, or non-integer
        ' interval cannot be blank, negative, or non-double
        ' only need to check field < number of columns (hidden columns/rows don't count)
        ' listboxes should be valid
        Dim i As Integer
        Dim errS As String = ""
        Dim txtBox As TextBox
        Dim lstBox As ComboBox
        Dim ival As Integer = 0
        Dim dval As Double = 0

        ' Note - _numCols and _useRow() have already been changed in the callbacks

        For i = 0 To _numCols - 1
            If (_useRow(USE_COUNT)) Then
                txtBox = _txtCount.Item(i + 1)
                errS = errS & ValidInteger(txtBox.Text, "Count")
            End If
            If (_useRow(USE_INTERVAL)) Then
                txtBox = _txtInterval.Item(i + 1)
                errS = errS & ValidDouble(txtBox.Text, "Interval")
            End If
        Next

        If (errS <> "") Then
            MsgBox("Invalid fields" & vbCrLf & vbCrLf & errS, MsgBoxStyle.Critical, "Data Errors")
            Return False
        End If

        ' Update the fields - they should be good
        For i = 0 To _numCols - 1
            If (_useRow(USE_COUNT)) Then
                txtBox = _txtCount.Item(i + 1)
                Integer.TryParse(txtBox.Text, ival)
                _count(i) = ival
            End If

            If (_useRow(USE_INTERVAL)) Then
                txtBox = _txtInterval.Item(i + 1)
                Double.TryParse(txtBox.Text, dval)
                _interval(i) = dval
            End If

            If (_useRow(USE_BIN)) Then
                lstBox = _lstBin.Item(i + 1)
                _bin(i) = lstBox.SelectedItem
            End If

            If (_useRow(USE_FILTER)) Then
                lstBox = _lstFilter.Item(i + 1)
                _filter(i) = lstBox.SelectedItem
            End If
        Next

        Return True
    End Function


    Public Sub New(cmd As String)
        ' must be format 
        MyBase.New()
        Dim i As Integer
        Dim obj() As Control
        Dim tb As TextBox
        Dim cb As ComboBox

        ' fill in controls
        _txtCount = New Collection
        _txtInterval = New Collection
        _lstBin = New Collection
        _lstFilter = New Collection
        For i = 1 To MAXCOLS
            obj = MainForm.pnlCombo.Controls.Find("txtComboCount" & i.ToString(), True)
            tb = obj(0)
            _txtCount.Add(tb, tb.Name)
            obj = MainForm.pnlCombo.Controls.Find("txtComboInterval" & i.ToString(), True)
            tb = obj(0)
            _txtInterval.Add(tb, tb.Name)
            obj = MainForm.pnlCombo.Controls.Find("lstComboBin" & i.ToString(), True)
            cb = obj(0)
            _lstBin.Add(cb, cb.Name)
            obj = MainForm.pnlCombo.Controls.Find("lstComboFilter" & i.ToString(), True)
            cb = obj(0)
            _lstFilter.Add(cb, cb.Name)
        Next
        
        ' my default command has piped commands for count, interval, etc
        ' need to break up the pipes into individual commands, which then direct the
        ' initialization of the widgets
        Dim cmds() As String = cmd.Split("|")
        Dim numCountEntries As Integer = -1
        Dim numIntervalEntries As Integer = -1
        Dim numFilterEntries As Integer = -1
        Dim numBinEntries As Integer = -1
        Dim data() As String
        Dim pieces() As String
        Dim countData As String = ""
        Dim intervalData As String = ""
        Dim binData As String = ""
        Dim filterData As String = ""

        For i = 0 To cmds.Count - 1
            pieces = cmds(i).Split(" ")
            Dim s As String = cmds(i).Substring(pieces(0).Length)
            s = s.Replace(" ", "")
            s = s.Replace(vbCr, "").Replace(vbLf, "")
            data = s.Split(",")
            If (LCase(pieces(0)) = "#count") Then
                numCountEntries = data.Count
                countData = Trim(s)
                _useRow(USE_COUNT) = True
            ElseIf (LCase(pieces(0)) = "#interval") Then
                numIntervalEntries = data.Count
                intervalData = Trim(s)
                _useRow(USE_INTERVAL) = True
            ElseIf (LCase(pieces(0)) = "#binning") Then
                numBinEntries = data.Count
                binData = Trim(s)
                _useRow(USE_BIN) = True
            ElseIf (LCase(pieces(0)) = "#filter") Then
                numFilterEntries = data.Count
                filterData = Trim(s)
                _useRow(USE_FILTER) = True
            End If
        Next
        ' are all choice counts the same? i.e., #count 1,1 | #interval 1,1,1  is bad
        Dim maxEntries As Integer = Math.Max(numCountEntries, numIntervalEntries)
        maxEntries = Math.Max(maxEntries, numBinEntries)
        maxEntries = Math.Max(maxEntries, numFilterEntries)
        If ((numCountEntries = maxEntries Or numCountEntries < 0) And _
            (numIntervalEntries = maxEntries Or numIntervalEntries < 0) And _
            (numBinEntries = maxEntries Or numBinEntries < 0) And _
            (numFilterEntries = maxEntries Or numFilterEntries < 0)) Then
            ' ok
        Else
            MsgBox("Count/Interval/Filter/Binning count mismatch")

        End If
        If ((numCountEntries > 0) And (numCountEntries < maxEntries)) Then
            ' count data needs to be extended
            For i = numCountEntries To maxEntries - 1
                countData = countData & ",1"
            Next
        End If
        If ((numIntervalEntries > 0) And (numIntervalEntries < maxEntries)) Then
            ' Interval data needs to be extended
            For i = numIntervalEntries To maxEntries - 1
                intervalData = intervalData & ",1"
            Next
        End If
        If ((numBinEntries > 0) And (numBinEntries < maxEntries)) Then
            ' Bin data needs to be extended
            For i = numBinEntries To maxEntries - 1
                binData = binData & ",1"
            Next
        End If
        If ((numFilterEntries > 0) And (numFilterEntries < maxEntries)) Then
            ' Filter data needs to be extended
            For i = numFilterEntries To maxEntries - 1
                filterData = filterData & ",1"
            Next
        End If

        _numCols = maxEntries
        EnableColumns()
        EnableRows()

        For i = 0 To cmds.Count - 1
            pieces = MyBase.CleanCommandString(cmds(i))    ' pieces(0) is like #count  pieces(1) is 1,1,1,1
            If (LCase(pieces(0)) = "#count") Then
                InitCount(countData)
            ElseIf (LCase(pieces(0)) = "#interval") Then
                InitInterval(intervalData)
            ElseIf (LCase(pieces(0)) = "#binning") Then
                InitBin(binData)
            ElseIf (LCase(pieces(0)) = "#filter") Then
                InitFilter(filterData)
            End If
        Next


    End Sub

    Private Sub InitCount(s As String)
        ' s is something like 1,1,1
        Dim data() As String = s.Split(",")

        For i = 0 To data.Count - 1
            Integer.TryParse(data(i), _count(i))
        Next

    End Sub

    Private Sub InitInterval(s As String)
        ' s is something like 1,1,1
        Dim data() As String = s.Split(",")

        For i = 0 To data.Count - 1
            Double.TryParse(data(i), _interval(i))
        Next
    End Sub

    Private Sub InitFilter(s As String)
        ' s is something like Clear,Red
        Dim data() As String = s.Split(",")
        Dim lstBox As ComboBox

        For i = 0 To data.Count - 1
            _filter(i) = data(i)
            ' Need to set the lstBox
            lstBox = _lstFilter.Item(i + 1)
            lstBox.SelectedIndex = lstBox.FindStringExact(Trim(data(i)))
            If (lstBox.SelectedIndex < 0) Then
                MsgBox("Filter " & data(i) & " not found in choices." & vbCrLf & "Either the plan has a bad filter, or" & vbCrLf & "the Filter Preferences need to changed.", , "Bad Filter Selection")
            End If
        Next

    End Sub

    Private Sub InitBin(s As String)
        ' s is something like 1,1,1
        Dim data() As String = s.Split(",")
        Dim lstBox As ComboBox

        For i = 0 To data.Count - 1
            _bin(i) = Trim(data(i))
            lstBox = _lstBin.Item(i + 1)
            lstBox.SelectedIndex = lstBox.FindStringExact(_bin(i))
            If (lstBox.SelectedIndex < 0) Then
                MsgBox("Bin " & data(i) & " not found in choices.", , "Bad Binning Selection")
            End If
        Next

    End Sub

    Private Sub EnableColumns()
        ' making 1-4 columns of values visible
        Dim i As Integer

        If (_numCols > MAXCOLS) Then _numCols = MAXCOLS

        ' Set the numCols combobox
        myLstComboColCount.SelectedIndex = _numCols - 1

        For i = 1 To _numCols
            If (_useRow(USE_COUNT)) Then _txtCount.Item(i).Visible = True
            If (_useRow(USE_INTERVAL)) Then _txtInterval.Item(i).Visible = True
            If (_useRow(USE_BIN)) Then _lstBin.Item(i).Visible = True
            If (_useRow(USE_FILTER)) Then _lstFilter.Item(i).Visible = True
        Next
        For i = _numCols + 1 To MAXCOLS
            _txtCount.Item(i).Visible = False
            _txtInterval.Item(i).Visible = False
            _lstBin.Item(i).Visible = False
            _lstFilter.Item(i).Visible = False
        Next
    End Sub

    Private Sub EnableRows()
        Dim i As Integer
        Dim txt As TextBox
        Dim lst As ComboBox

        For i = 0 To _numCols - 1
            txt = _txtCount.Item(i + 1)
            txt.Visible = _useRow(USE_COUNT)
            myChkComboCount.Checked = _useRow(USE_COUNT)
            txt = _txtInterval.Item(i + 1)
            txt.Visible = _useRow(USE_INTERVAL)
            myChkComboInterval.Checked = _useRow(USE_INTERVAL)
            lst = _lstBin.Item(i + 1)
            lst.Visible = _useRow(USE_BIN)
            myChkComboBin.Checked = _useRow(USE_BIN)
            lst = _lstFilter.Item(i + 1)
            lst.Visible = _useRow(USE_FILTER)
            myChkComboFilter.Checked = _useRow(USE_FILTER)
        Next
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        ' turn off comment field
        MainForm.txtNote.Visible = False
        MainForm.lblNote.Visible = False

        Dim i As Integer
        EnableColumns()
        EnableRows()
        MainForm.pnlCombo.Visible = True
        For i = 0 To _numCols - 1
            If (_useRow(USE_COUNT)) Then _txtCount.Item(i + 1).Text = _count(i).ToString()
            If (_useRow(USE_INTERVAL)) Then _txtInterval.Item(i + 1).Text = _interval(i).ToString()
            If (_useRow(USE_FILTER)) Then _lstFilter.Item(i + 1).SelectedIndex = _lstFilter.Item(i + 1).FindStringExact(_filter(i))
            If (_lstFilter.Item(i + 1).SelectedIndex < 0) Then
                ' bad filter - use first filter
                _lstFilter.Item(i + 1).SelectedIndex = 0
            End If
            If (_useRow(USE_BIN)) Then _lstBin.Item(i + 1).SelectedIndex = _lstBin.Item(i + 1).FindStringExact(_bin(i))
            If (_lstBin.Item(i + 1).SelectedIndex < 0) Then
                ' bad bin - use first filter
                _lstBin.Item(i + 1).SelectedIndex = 0
            End If
        Next
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the lstCommands listbox will display the correct label
        Dim lineText As String = ""
        Dim i As Integer
        If (_useRow(USE_COUNT)) Then
            lineText = "#Count "
            For i = 0 To _numCols - 1
                If i = 0 Then
                    lineText = lineText & _count(i).ToString()
                Else
                    lineText = lineText & ", " & _count(i).ToString()
                End If
            Next
            lineText = lineText & Environment.NewLine
        End If
        If (_useRow(USE_INTERVAL)) Then
            lineText = lineText & "#Interval "
            For i = 0 To _numCols - 1
                If i = 0 Then
                    lineText = lineText & _interval(i).ToString()
                Else
                    lineText = lineText & ", " & _interval(i).ToString()
                End If
            Next
            lineText = lineText & Environment.NewLine
        End If
        If (_useRow(USE_BIN)) Then
            lineText = lineText & "#Binning "
            For i = 0 To _numCols - 1
                If i = 0 Then
                    lineText = lineText & _bin(i)
                Else
                    lineText = lineText & ", " & _bin(i)
                End If
            Next
            lineText = lineText & Environment.NewLine
        End If
        If (_useRow(USE_FILTER)) Then
            lineText = lineText & "#Filter "
            For i = 0 To _numCols - 1
                If i = 0 Then
                    lineText = lineText & _filter(i)
                Else
                    lineText = lineText & ", " & _filter(i)
                End If
            Next
            lineText = lineText & Environment.NewLine
        End If

        ' remove the last newline
        lineText = lineText.Remove(lineText.Length - 1, 1)
        Return lineText
    End Function


#Region "Field Control Events"
    Private WithEvents myBtnComboApply As Button = MainForm.btnComboApply
    Private WithEvents myLstComboColCount As ComboBox = MainForm.lstComboColCount
    Private WithEvents myChkComboCount As CheckBox = MainForm.chkComboCount
    Private WithEvents myChkComboInterval As CheckBox = MainForm.chkComboInterval
    Private WithEvents myChkComboBin As CheckBox = MainForm.chkComboBin
    Private WithEvents myChkComboFilter As CheckBox = MainForm.chkComboFilter

    Private Sub btnComboApply_Click(sender As Object, e As EventArgs) Handles myBtnComboApply.Click
        Try
            If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
                If (UpdateFields()) Then
                    MainForm.GetPlan().Update()       ' update the plan
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Critical, "Data Error")
        End Try

    End Sub

    Private Sub lstComboColCount_SelectedIndexChanged(sender As Object, e As EventArgs) Handles myLstComboColCount.SelectedIndexChanged
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            Integer.TryParse(myLstComboColCount.Text, _numCols)
            EnableColumns()
        End If
    End Sub

    Private Sub chkComboRow_CheckedChanged(sender As Object, e As EventArgs) Handles myChkComboCount.CheckedChanged, myChkComboInterval.CheckedChanged, myChkComboBin.CheckedChanged, myChkComboFilter.CheckedChanged
        If ((MainForm._activeCommand IsNot Nothing) And (ReferenceEquals(Me, MainForm._activeCommand))) Then
            If (ReferenceEquals(sender, myChkComboCount)) Then
                _useRow(USE_COUNT) = myChkComboCount.Checked
            ElseIf (ReferenceEquals(sender, myChkComboInterval)) Then
                _useRow(USE_INTERVAL) = myChkComboInterval.Checked
            ElseIf (ReferenceEquals(sender, myChkComboBin)) Then
                _useRow(USE_BIN) = myChkComboBin.Checked
            ElseIf (ReferenceEquals(sender, myChkComboFilter)) Then
                _useRow(USE_FILTER) = myChkComboFilter.Checked
            End If
            EnableRows()
        End If
        
    End Sub


#End Region


End Class
