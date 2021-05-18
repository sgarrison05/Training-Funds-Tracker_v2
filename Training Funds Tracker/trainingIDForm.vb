Public Class trainingIDForm

    Private rootfileLocation As String = "C:\Training\"
    Private activityFileLocation As String = "C:\Training\trainingrun.txt"
    Private bankFileLocation As String = "C:\Training\training.txt"


    Private Sub btnDone_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDone.Click

        'declares variables
        Dim name As String
        Dim location As String
        Dim startDate As String
        Dim endDate As String


        'get input
        name = Me.txtTrainingName.Text
        location = Me.txtTrainingLocation.Text
        startDate = Me.dtpTrainingStart.Text
        endDate = Me.dtpTrainingEnd.Text

        'Enters Initial text to the trainingrun.txt file
        If My.Computer.FileSystem.DirectoryExists(rootfileLocation) Then
            My.Computer.FileSystem.WriteAllText(activityFileLocation, _
            "Name:" & Strings.Space(7) & name & ControlChars.NewLine & "Location:" & _
            Strings.Space(3) & location & ControlChars.NewLine & "Dates:" & _
            Strings.Space(6) & startDate & Strings.Space(2) & "-" & _
            Strings.Space(2) & endDate & ControlChars.NewLine, _
            True)
            My.Computer.FileSystem.WriteAllText(activityFileLocation, "".PadLeft(105, "-") & _
            ControlChars.NewLine, True)

        Else 'it creates the director and trainingrun.txt file with initial information
            My.Computer.FileSystem.CreateDirectory(rootfileLocation)
            My.Computer.FileSystem.WriteAllText(activityFileLocation, _
            "Name:" & Strings.Space(7) & name & ControlChars.NewLine & "Location:" & _
            Strings.Space(3) & location & ControlChars.NewLine & "Dates:" & _
            Strings.Space(6) & startDate & Strings.Space(2) & "-" & _
            Strings.Space(2) & endDate & ControlChars.NewLine, _
            True)
            My.Computer.FileSystem.WriteAllText(activityFileLocation, "".PadLeft(105, "-") & _
            ControlChars.NewLine, True)

        End If

        'closes the form
        Me.Close()

    End Sub

    Private Sub trainingIDForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'sets focus to first text box
        Me.txtTrainingName.Focus()

    End Sub
End Class