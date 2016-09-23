Public Class CFlatCommand : Inherits CommandBaseClass
    ' Flat Command is a command with only the command itself, no parameters
    Public Sub New(cmd As String)
        MyBase.New()
        Dim pieces As String() = MyBase.CleanCommandString(cmd)
    End Sub

    Public Overrides Sub Display()
        MyBase.Display()
        MainForm.pnlFlatCommand.Visible = True
    End Sub

    Public Overrides Function ToString() As String
        'Required so that the listbox will display the correct label
        Dim lineText As String = ""
        If (MyBase._commandDescription = "Comment") Then
            lineText = ";" & MyBase._comment
        Else
            ' regular command
            lineText = MyBase._command
            If (MyBase._comment <> "") Then
                lineText = lineText & "    ; " & MyBase._comment
            End If
        End If
        Return lineText
    End Function
End Class
