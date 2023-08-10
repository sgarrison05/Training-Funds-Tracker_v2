Public Class trainingIDForm

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
        If My.Computer.FileSystem.DirectoryExists(mainForm.rdirectory) Then
            Me.Close()

        Else

            'it creates the directory and trainingrun.txt file with initial information heading
            My.Computer.FileSystem.CreateDirectory(mainForm.rdirectory)

            My.Computer.FileSystem.WriteAllText(mainForm.rfile,
            "Name:" & Strings.Space(7) &
            name & ControlChars.NewLine &
            "Location:" & Strings.Space(3) &
            location & ControlChars.NewLine &
            "Dates:" & Strings.Space(6) &
            startDate & Strings.Space(2) &
            "-" & Strings.Space(2) &
            endDate & ControlChars.NewLine, True)

            mainForm.Separation()

            mainForm.CreateEntry(mainForm.payee)
            mainForm.Show()

            Me.Close()

        End If

    End Sub

    Private Sub trainingIDForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'sets focus to first text box
        Me.txtTrainingName.Focus()

    End Sub
End Class