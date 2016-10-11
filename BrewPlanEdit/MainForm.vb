'   BrewPlanEdit
' Program to edit ACP plans.
' 
' Allows multiple tabs.
' Provides preformatted ACP commands. 
' Times are handled In local time, converts To UTC.
' Help button provides help info copied from ACP User manual.
' Groups Count/Interval/Bin/Filter commands together so they are in sync
'
' version 1.1.0.0 - Fixed a couple of bugs, I forget
'         1.2.0.0 - Added ability to change font for old men. Also saves program size and position.

Public Class MainForm

    Public panelList As Collection      ' all of the command panels

    Public _activeCommand As Object        ' this is the command being edited. Holds a CChill, CTarget, whatever with
    ' fields needing to be updated

    Private Sub lstCommands_SelectedIndexChanged(sender As Object, e As EventArgs) Handles lstCommands.SelectedIndexChanged
        ' this callback is used by the various listboxes in the tabs
        Dim lstBox As ListBox = sender
        If (lstBox.SelectedItem IsNot Nothing) Then
            ConfigureContextMenu(True, lstBox)
            _activeCommand = lstBox.SelectedItem
            _activeCommand.Display()
        Else
            ' nothing selected - should be an empty plan
            _activeCommand = Nothing
            HidePanels()
            ConfigureContextMenu(False, lstBox)
        End If

    End Sub

    Private Sub lstCommands_MeasureItem(sender As Object, e As MeasureItemEventArgs) Handles lstCommands.MeasureItem
        Dim lstBox As ListBox = sender
        e.ItemHeight = e.Graphics.MeasureString(lstBox.Items(e.Index).ToString, lstBox.Font).Height
    End Sub

    Private Sub lstCommands_DrawItem(sender As Object, e As DrawItemEventArgs) Handles lstCommands.DrawItem
        ' special routine to allow colored font for commands with errors/warnings
        'e.DrawBackground()      Sample code from somewhere

        'If ListBox1.Items(e.Index).ToString() = "herp" Then

        '    e.Graphics.FillRectangle(Brushes.LightGreen, e.Bounds)
        'End If
        'e.Graphics.DrawString(ListBox1.Items(e.Index).ToString(), e.Font, Brushes.Black, New System.Drawing.PointF(e.Bounds.X, e.Bounds.Y))
        'e.DrawFocusRectangle()
        '----------------
        Dim lstBox As ListBox = sender
        e.DrawBackground()
        If (e.Index >= 0) Then
            Dim obj As CommandBaseClass = lstBox.Items(e.Index)
            Dim s As String = obj.Command
            If (obj.Warning) Then
                e.Graphics.DrawString(lstBox.Items(e.Index).ToString, lstBox.Font, Brushes.Red, e.Bounds)
            Else
                e.Graphics.DrawString(lstBox.Items(e.Index).ToString, lstBox.Font, Brushes.Black, e.Bounds)
            End If
            e.DrawFocusRectangle()
        End If
    End Sub

    Private lastMousePosition As Point
    Private Sub lstCommands_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles lstCommands.MouseMove

        Dim MousePositionInClientCoords As Point = sender.PointToClient(MousePosition)
        'Check if the mouse has moved a reasonable distance  
        If MousePositionInClientCoords.Y > lastMousePosition.Y - 5 And
            MousePositionInClientCoords.Y < lastMousePosition.Y + 5 And
            MousePositionInClientCoords.X > lastMousePosition.X - 10 And
            MousePositionInClientCoords.X < lastMousePosition.X + 10 Then
            'do nothing  

        Else
            'Save the current mouse position as the last mouse position  
            lastMousePosition = MousePositionInClientCoords

            Dim indexUnderTheMouse As Integer = sender.IndexFromPoint(MousePositionInClientCoords)
            Dim lb As ListBox = sender
            If ((indexUnderTheMouse > -1) And (indexUnderTheMouse < 1000) And (indexUnderTheMouse = lb.SelectedIndex)) Then     ' somehow index is being 65535, not -1
                Dim s As String = sender.items(indexUnderTheMouse).ToString()
                Dim g As Graphics = sender.CreateGraphics
                If g.MeasureString(s, sender.Font).Width > sender.ClientRectangle.Width Then
                    ToolTip1.SetToolTip(sender, s)
                Else
                    ToolTip1.SetToolTip(sender, "")
                End If
                g.Dispose()
            Else
                ToolTip1.SetToolTip(sender, "")
            End If
        End If
    End Sub


    Private Sub ConfigureContextMenu(enable As Boolean, lstBox As ListBox)
        ' enable is true if a listItem has been selected
        DeleteToolStripMenuItem.Enabled = False
        MoveDownToolStripMenuItem.Enabled = False
        MoveUpToolStripMenuItem.Enabled = False
        CommentLineToolStripMenuItem.Enabled = False
        UncommentLineToolStripMenuItem.Enabled = False
        If enable Then
            ' list item is selected
            DeleteToolStripMenuItem.Enabled = enable
            If (lstBox.SelectedIndex > 0) Then
                MoveUpToolStripMenuItem.Enabled = True
            End If
            If (lstBox.SelectedIndex < lstBox.Items.Count - 1) Then
                MoveDownToolStripMenuItem.Enabled = True
            End If

            ' enable comment/uncomment based on whether the current item is a comment line
            Dim selection As CommandBaseClass = lstBox.SelectedItem
            If (selection.Command = ";") Then
                UncommentLineToolStripMenuItem.Enabled = True
            Else
                CommentLineToolStripMenuItem.Enabled = True
            End If
        End If
    End Sub

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ReadSettings()

        ' build the panel list
        panelList = New Collection
        panelList.Add(pnlFlatCommand, "pnlFlatCommand")
        panelList.Add(pnlChill, "pnlChill")
        panelList.Add(pnlTarget, "pnlTarget")
        panelList.Add(pnlCombo, "pnlCombo")
        panelList.Add(pnlDateCmd, "pnlDateCmd")
        panelList.Add(pnlMultiVal, "pnlMultiVal")
        panelList.Add(pnlDirCommand, "pnlDirCommand")
        panelList.Add(pnlStringCommand, "pnlStringCommand")

        Dim pnl As Panel
        For Each pnl In panelList
            pnl.Dock = Windows.Forms.DockStyle.Fill
        Next
        HidePanels()
        TabControl1.SendToBack()            ' so documentation panel is on top
        WebBrowser1.Visible = False
        WebBrowser1.DocumentText = "<html><body>No command selected</body></html>"

        lblCommandName.Text = ""
        LoadFilters()

        ' Put version info into heading
        Me.Text = Me.Text & " " & String.Format("Version {0}", My.Application.Info.Version.ToString)

        ' create plan for initial tab
        Dim plan As CPlan = New CPlan
        plan.FileName = "NewPlan.txt"
        plan.ListBox = lstCommands
        Dim tabPage As Windows.Forms.TabPage = TabControl1.TabPages(0)
        tabPage.Tag = plan
        ConfigureContextMenu(False, lstCommands)

        ' set up help
        Dim strHelpPath As String = System.IO.Path.Combine(Application.StartupPath, "BrewPlanEdit.chm")
        HelpProvider1.HelpNamespace = strHelpPath

    End Sub

    Private Sub HidePanels()
        Dim pnl As Panel
        For Each pnl In panelList
            pnl.Visible = False
        Next
        pnlBase.Visible = False
        txtNote.Text = ""
    End Sub

    Private Function AreAnyPlansModified() As Boolean
        ' step through the tabs, see if any have been modified
        ' if so, return true
        ' return false if none have been modified
        Dim modified As Boolean = False
        Dim theTabPage As TabPage
        For Each theTabPage In TabControl1.TabPages
            Dim plan As CPlan = theTabPage.Tag
            If (plan.Modified) Then
                modified = True
            End If
        Next
        Return modified
    End Function
#Region "Main Menu Callbacks"
    Private Sub OpenToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles OpenToolStripMenuItem.Click
        ' open the plan into the existing tab

        ' First check whether current plan needs saving
        Dim plan As CPlan = GetPlan()
        If (plan.Modified) Then
            Dim res As MsgBoxResult = MsgBox("Plan has been modified. Do you wish to save the plan?", MsgBoxStyle.YesNoCancel, "Plan has been modified")
            If (res = MsgBoxResult.Yes) Then
                'save first
                SaveAsToolStripMenuItem_Click(Nothing, Nothing)
            ElseIf (res = MsgBoxResult.No) Then
                ' no save, continue with opening
            Else
                'Cancel
                Exit Sub
            End If
        End If

        plan.Modified = False
        Dim filename As String = plan.OpenPlan()
        If (filename <> "") Then
            TabControl1.TabPages(TabControl1.SelectedIndex).Text = filename & "   "
        End If
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        GetPlan().SavePlan()
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        If (SaveFileDialog.ShowDialog = DialogResult.OK) Then
            GetPlan().SaveAsPlan(SaveFileDialog.FileName)
            TabControl1.TabPages(TabControl1.SelectedIndex).Text = GetPlan().FileName & "   "
        End If
    End Sub

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Dim plan As CPlan = GetPlan()
        If (plan.Modified) Then
            Dim res As MsgBoxResult = MsgBox("Plan has been modified. Do you wish to save the plan?", MsgBoxStyle.YesNoCancel, "Plan has been modified")
            If (res = MsgBoxResult.Yes) Then
                'save first
                SaveAsToolStripMenuItem_Click(Nothing, Nothing)
            ElseIf (res = MsgBoxResult.No) Then
                ' no save, close it
            Else
                'Cancel
                Exit Sub
            End If
        End If

        ' now do the close
        _activeCommand = Nothing
        Dim lstBox As ListBox = plan.ListBox
        lstBox.Items.Clear()
        If (TabControl1.SelectedIndex > 0) Then
            ' closing a secondary tab
            lstBox.Items.Clear()
            lstBox.Dispose()                ' destroy listbox
            plan.ListBox = Nothing
            plan = Nothing                  ' destroy plan

            ' remove tab
            Dim idx As Integer = TabControl1.SelectedIndex
            TabControl1.TabPages.RemoveAt(idx)
            idx = idx - 1
            If (idx >= 0) Then
                TabControl1.SelectedIndex = idx
            End If
        Else
            ' primary tab
            plan.Modified = False
            plan.FileName = "NewPlan.txt"
            TabControl1.TabPages(0).Text = "NewPlan.txt   "
            TabControl1_Selected(Nothing, Nothing)
        End If
    End Sub

    Private Sub ExitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ExitToolStripMenuItem.Click
        If (AreAnyPlansModified()) Then
            If MessageBox.Show("Some plans have been modified." & vbCrLf & "Are you sure you wish to close BrewPlanEdit?", "Close", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                SaveSettings()
                End
            End If
        Else
            SaveSettings()
            End
        End If
    End Sub

    Private Sub Form_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        If (AreAnyPlansModified()) Then
            If MessageBox.Show("Some plans have been modified." & vbCrLf & "Are you sure you wish to close BrewPlanEdit?", "Close", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = Windows.Forms.DialogResult.Yes Then
                SaveSettings()
            Else
                e.Cancel = True
            End If
        Else
            SaveSettings()
        End If
    End Sub

    Private Sub AddTabToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddTabToolStripMenuItem.Click
        ' Add a new tab to the tab control
        Dim plan As CPlan = New CPlan
        Dim newfilename As String = "NewPlan.txt"
        Dim num As Integer = 1
        Dim newTabName As String = "TabPage" & num.ToString()
        While (TabControl1.TabPages.ContainsKey(newTabName))
            num = num + 1
            newTabName = "TabPage" & num.ToString()
        End While
        newfilename = "NewPlan" & num.ToString() & ".txt"
        TabControl1.TabPages.Add(newTabName, newfilename & "   ")
        Dim newPage As Windows.Forms.TabPage
        newPage = TabControl1.TabPages(newTabName)

        ' now put the list control into tab
        Dim lstBox As ListBox = New ListBox()
        newPage.Controls.Add(lstBox)
        newPage.Tag = plan                ' save plan for when tabs change
        lstBox.Name = "lstCommands" & num.ToString()
        lstBox.Dock = DockStyle.Fill
        lstBox.Font = lstCommands.Font          ' use the larger font
        lstBox.DrawMode = DrawMode.OwnerDrawVariable
        lstBox.ContextMenuStrip = menuCommand
        ' Set this listbox to use the same event handler?
        AddHandler lstBox.SelectedIndexChanged, AddressOf lstCommands_SelectedIndexChanged
        AddHandler lstBox.MeasureItem, AddressOf lstCommands_MeasureItem
        AddHandler lstBox.DrawItem, AddressOf lstCommands_DrawItem
        AddHandler lstBox.MouseMove, AddressOf lstCommands_MouseMove
        plan.ListBox = lstBox
        plan.FileName = newfilename

        ' empty tabPage, so hide the panel fields
        HidePanels()

        TabControl1.SelectTab(newTabName)
    End Sub
#End Region

#Region "Context Menu Callbacks"
    Private Sub MoveUpToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MoveUpToolStripMenuItem.Click
        MoveItem(-1)
    End Sub

    Private Sub MoveDownToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MoveDownToolStripMenuItem.Click
        MoveItem(1)
    End Sub

    Public Sub MoveItem(direction As Integer)
        Dim lstBox As ListBox = GetPlan().ListBox
        ' Checking selected item
        If (lstBox.SelectedIndex < 0) Then
            Exit Sub ' No selected item - nothing to do
        End If

        ' Calculate new index using move direction
        Dim newIndex As Integer = lstBox.SelectedIndex + direction

        ' Checking bounds of the range
        If ((newIndex < 0) Or (newIndex >= lstBox.Items.Count)) Then
            Exit Sub     ' Index out of range - nothing to do
        End If

        Dim selected As Object = lstBox.SelectedItem

        ' Removing removable element
        lstBox.Items.Remove(selected)
        ' Insert it in new position
        lstBox.Items.Insert(newIndex, selected)
        ' Restore selection
        lstBox.SetSelected(newIndex, True)
    End Sub

    Private Sub DeleteToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteToolStripMenuItem.Click
        ' delete the currently selected command 
        If (_activeCommand IsNot Nothing) Then
            Dim lstBox As ListBox = GetPlan().ListBox
            Dim idx As Integer = lstBox.SelectedIndex
            lstBox.Items.RemoveAt(idx)
            _activeCommand = Nothing
            If (idx < lstBox.Items.Count) Then
                lstBox.SetSelected(idx, True)
            ElseIf ((idx = lstBox.Items.Count) And (idx > 0)) Then
                lstBox.SetSelected(idx - 1, True)
            End If
            ' mark plan as updated
            GetPlan().Update()
        End If
    End Sub

#End Region

    Private Sub TabControl1_Selected(sender As Object, e As EventArgs) Handles TabControl1.SelectedIndexChanged
        ' New tab has been clicked. Need to display the lstBox fields for the tab
        Dim plan As CPlan = GetPlan()           ' retrieve plan
        lstCommands_SelectedIndexChanged(plan.ListBox, e)
    End Sub

    Private Sub TabControl1_DrawItem(ByVal sender As Object, ByVal e As System.Windows.Forms.DrawItemEventArgs) Handles TabControl1.DrawItem
        ' This needs the tab control to have DrawMode set to OwnerDrawFixed
        Dim g As Graphics = e.Graphics
        Dim f As Font = New Font(e.Font, FontStyle.Bold)
        Dim b As New SolidBrush(Color.LightBlue)
        If e.Index = Me.TabControl1.SelectedIndex Then
            b.Color = Color.Aquamarine
            g.FillRectangle(b, e.Bounds)
            b.Color = Color.Black    ' font color
            g.DrawString(Me.TabControl1.TabPages(e.Index).Text, f, b, e.Bounds.X + 2, e.Bounds.Y + 2)
        Else
            b.Color = Color.CadetBlue
            g.FillRectangle(b, e.Bounds)
            b.Color = Color.Black     ' font color
            g.DrawString(Me.TabControl1.TabPages(e.Index).Text, e.Font, b, e.Bounds.X + 2, e.Bounds.Y + 2)
        End If
        b.Dispose()
        f.Dispose()
    End Sub


    ' these routines handle updating of the various fields on the commands
    Public Function GetPlan() As CPlan
        GetPlan = TabControl1.SelectedTab.Tag
    End Function




#Region "Context Menu Add command callbacks"
    ' routines to handle Add from context menu
    Private Sub ChillToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ChillToolStripMenuItem.Click,
                                        NoWeatherToolStripMenuItem.Click,
                                        DomeOpenToolStripMenuItem.Click,
                                        DomeCloseToolStripMenuItem.Click,
                                        AutofocusToolStripMenuItem.Click,
                                        CountToolStripMenuItem.Click,
                                        DitherToolStripMenuItem.Click,
                                        SubFrameToolStripMenuItem.Click,
                                        DirectoryToolStripMenuItem.Click,
                                        TagToolStripMenuItem.Click,
                                        PosAngToolStripMenuItem.Click,
                                        ReadoutModeToolStripMenuItem.Click,
                                        ManualToolStripMenuItem.Click,
                                        AfIntervalToolStripMenuItem.Click,
                                        CalibrateToolStripMenuItem.Click,
                                        StackToolStripMenuItem.Click,
                                        StackAlignToolStripMenuItem.Click,
                                        AutoguideToolStripMenuItem.Click,
                                        PointingToolStripMenuItem.Click,
                                        NoPointingToolStripMenuItem.Click,
                                        NoPreviewToolStripMenuItem.Click,
                                        NoSolveToolStripMenuItem.Click,
                                        TrackOnToolStripMenuItem.Click,
                                        TrackOffToolStripMenuItem.Click,
                                        AlwaysSolveToolStripMenuItem.Click,
                                        WaitForToolStripMenuItem.Click,
                                        WaitUntilToolStripMenuItem.Click,
                                        WaitZenDistToolStripMenuItem.Click,
                                        WaitInLimitsToolStripMenuItem.Click,
                                        WaitAirMassToolStripMenuItem.Click,
                                        DuskFlatsToolStripMenuItem.Click,
                                        DawnFlatsToolStripMenuItem.Click,
                                        ScreenFlatsToolStripMenuItem.Click,
                                        DarkToolStripMenuItem.Click,
                                        BiasToolStripMenuItem.Click,
                                        SetsToolStripMenuItem.Click,
                                        RepeatToolStripMenuItem.Click,
                                        MinSetTimeToolStripMenuItem.Click,
                                        QuitAtToolStripMenuItem.Click,
                                        ShutdownToolStripMenuItem.Click,
                                        ShutdownAtToolStripMenuItem.Click,
                                        ChainToolStripMenuItem.Click,
                                        ChainScriptToolStripMenuItem.Click,
                                        DefocusToolStripMenuItem.Click, _
 _
                                        ChillToolStripMenuItem1.Click,
                                        NoWeatherToolStripMenuItem1.Click,
                                        DomeOpenToolStripMenuItem1.Click,
                                        DomeCloseToolStripMenuItem1.Click,
                                        AutofocusToolStripMenuItem1.Click,
                                        CountToolStripMenuItem1.Click,
                                        DitherToolStripMenuItem1.Click,
                                        SubframeToolStripMenuItem1.Click,
                                        DirectoryToolStripMenuItem1.Click,
                                        TagToolStripMenuItem1.Click,
                                        PosAngToolStripMenuItem1.Click,
                                        ReadoutmodeToolStripMenuItem1.Click,
                                        ManualToolStripMenuItem1.Click,
                                        AfIntervalToolStripMenuItem1.Click,
                                        CalibrateToolStripMenuItem1.Click,
                                        StackToolStripMenuItem1.Click,
                                        StackAlignToolStripMenuItem1.Click,
                                        AutoGuideToolStripMenuItem1.Click,
                                        PointingToolStripMenuItem1.Click,
                                        NoPointingToolStripMenuItem1.Click,
                                        NoPreviewToolStripMenuItem1.Click,
                                        NoSolveToolStripMenuItem1.Click,
                                        TrackOnToolStripMenuItem1.Click,
                                        TrackOffToolStripMenuItem1.Click,
                                        AlwaysSolveToolStripMenuItem1.Click,
                                        WaitForToolStripMenuItem1.Click,
                                        WaitForToolStripMenuItem1.Click,
                                        WaitZenDistToolStripMenuItem1.Click,
                                        WaitInLimitsToolStripMenuItem1.Click,
                                        WaitAirMassToolStripMenuItem1.Click,
                                        DuskFlatsToolStripMenuItem1.Click,
                                        DawnFlatsToolStripMenuItem1.Click,
                                        ScreenFlatsToolStripMenuItem1.Click,
                                        DarkToolStripMenuItem1.Click,
                                        BiasToolStripMenuItem1.Click,
                                        SetsToolStripMenuItem1.Click,
                                        RepeatToolStripMenuItem1.Click,
                                        MinSetTimeToolStripMenuItem1.Click,
                                        QuitAtToolStripMenuItem1.Click,
                                        ShutdownToolStripMenuItem1.Click,
                                        ShutdownAtToolStripMenuItem1.Click,
                                        ChainToolStripMenuItem1.Click,
                                        ChainScriptToolStripMenuItem1.Click,
                                        DefocusToolStripMenuItem1.Click

        Dim it As ToolStripMenuItem = sender
        GetPlan().AddDefault(it.Text)
    End Sub

    Private Sub CommentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CommentToolStripMenuItem.Click, CommentToolStripMenuItem1.Click
        GetPlan().AddDefault("; Comment")
    End Sub


    Private Sub TargetToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TargetToolStripMenuItem.Click, TargetToolStripMenuItem1.Click
        GetPlan().AddDefault("MyTarget")
    End Sub
    Private Sub FilterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FilterToolStripMenuItem.Click, FilterToolStripMenuItem1.Click,
        IntervalToolStripMenuItem.Click, IntervalToolStripMenuItem1.Click,
        BinningToolStripMenuItem.Click, BinningToolStripMenuItem1.Click
        GetPlan().AddDefault("#Count")
    End Sub
#End Region

    Private Sub btnPopDocumentation_Click(sender As Object, e As EventArgs) Handles btnPopDocumentation.Click
        WebBrowser1.Visible = Not WebBrowser1.Visible
        WebBrowser1.BringToFront()
    End Sub

    Private Sub AboutBrewPlanEditToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutBrewPlanEditToolStripMenuItem.Click
        AboutBox.ShowDialog()
    End Sub

    Private Sub ProfilesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ProfilesToolStripMenuItem.Click
        If (PrefDialog.ShowDialog() = Windows.Forms.DialogResult.OK) Then
            LoadFilters()
        End If
    End Sub

    Private Sub LoadFilters()
        ' Update the comboboxes in CCombo class
        Dim i, f, n As Integer
        Dim tb() As Control
        Dim cb As ComboBox

        ' which set is being used?
        i = My.Settings.FilterSetInUse
        Dim filterSet As Collections.Specialized.StringCollection = My.Settings.Item("FilterSet" & i)

        For n = 1 To CCombo.MAXCOLS
            tb = Me.Controls.Find("lstComboFilter" & n, True)
            If (tb.Length > 0) Then
                cb = tb(0)
                cb.Items.Clear()
                For f = 1 To filterSet.Count
                    cb.Items.Add(filterSet.Item(f - 1))
                Next
            End If
        Next

    End Sub

    Private Sub CommentLineToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CommentLineToolStripMenuItem.Click
        GetPlan().CommentSelectedLine()
    End Sub

    Private Sub UncommentLineToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles UncommentLineToolStripMenuItem.Click
        GetPlan().UncommentSelectedLine()
    End Sub

    Private Sub DocumentationNotesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DocumentationNotesToolStripMenuItem.Click
        Help.ShowHelp(Me, HelpProvider1.HelpNamespace, HelpNavigator.TableOfContents)
    End Sub

    Public Sub ReadSettings()
        ' Font.tostring gives [Font: Name=Microsoft Sans Serif, Size=8.25, Units=3, GdiCharSet=0, GdiVerticalFont=False]
        Dim font As Font
        Dim myFontStyle As FontStyle = My.Settings.FontStyle
        font = New Font(My.Settings.FontFamily, My.Settings.FontSize, myFontStyle)
        Me.Font = font
        PrefDialog.ActiveForm.Font = font

        Me.Width = My.Settings.MainWidth
        Me.Height = My.Settings.MainHeight
        Me.Top = My.Settings.MainTop
        Me.Left = My.Settings.MainLeft
    End Sub

    Public Sub SaveSettings()
        My.Settings.FontFamily = Me.Font.FontFamily.ToString()
        My.Settings.FontSize = Me.Font.Size
        My.Settings.FontStyle = Me.Font.Style
        My.Settings.MainWidth = Me.Width
        My.Settings.MainHeight = Me.Height
        My.Settings.MainTop = Me.Top
        My.Settings.MainLeft = Me.Left
        My.Settings.Save()
    End Sub
End Class
