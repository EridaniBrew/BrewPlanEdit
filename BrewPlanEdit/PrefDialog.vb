Imports System.Windows.Forms

Public Class PrefDialog
    Public Const MAX_FILTERS = 10          ' number of filters per set
    Public Const MAX_FILTER_SETS = 3       ' number of filter sets

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK

        ' Save the settings
        Dim f, n As Integer
        Dim tb() As Control
        Dim rb As RadioButton

        For n = 1 To MAX_FILTER_SETS
            Dim filterSet As Collections.Specialized.StringCollection = My.Settings.Item("FilterSet" & n)
            filterSet.Clear()
            For f = 1 To MAX_FILTERS
                tb = Me.Controls.Find("txtFilter" & n & "Num" & f, True)
                If (tb.Length > 0) Then
                    If (tb(0).Text <> "") Then
                        filterSet.Add(tb(0).Text)
                    Else
                        Exit For
                    End If
                End If

            Next
        Next

        ' get the currently selected filterset num
        For n = 1 To MAX_FILTER_SETS
            tb = Me.Controls.Find("radFilter" & n, True)     'tb is radioButton
            rb = tb(0)
            If (rb.Checked) Then
                My.Settings.FilterSetInUse = n
            End If
        Next

        My.Settings.Save()
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub PrefDialog_Load(sender As Object, e As EventArgs) Handles Me.Load
        ' load up the filter values from Settings
        Dim f, n As Integer
        Dim tb() As Control
        Dim rb As RadioButton

        For n = 1 To MAX_FILTER_SETS
            Dim filterSet As Collections.Specialized.StringCollection = My.Settings.Item("FilterSet" & n)
            For f = 1 To MAX_FILTERS
                tb = Me.Controls.Find("txtFilter" & n & "Num" & f, True)
                If (tb.Length > 0) Then
                    If (f <= filterSet.Count) Then
                        tb(0).Text = filterSet.Item(f - 1)
                    Else
                        tb(0).Text = ""
                    End If
                End If

            Next
        Next

        ' Select the current FilterSet being used
        tb = Me.Controls.Find("radFilter" & My.Settings.FilterSetInUse, True)    'rb is radioButton
        If (tb.Length > 0) Then
            rb = tb(0)
            rb.Checked = True
        End If

    End Sub

End Class
