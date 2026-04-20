Public Class trainingIDForm

    Private Sub btnDone_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnDone.Click

        'declares variables
        Dim name As String
        Dim location As String
        Dim startDate As String
        Dim endDate As String

        'get input
        name = Me.txtTrainingName.Text.Trim()
        location = Me.txtTrainingLocation.Text.Trim()
        startDate = Me.dtpTrainingStart.Value.ToShortDateString()
        endDate = Me.dtpTrainingEnd.Value.ToShortDateString()

        ' Validate input before proceeding
        If String.IsNullOrWhiteSpace(name) OrElse String.IsNullOrWhiteSpace(location) Then
            MessageBox.Show("Please enter both a training name and location.",
                            "Missing Information", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            Return
        End If

        ' If file already exists, skip setup and return to main form
        If My.Computer.FileSystem.FileExists(mainForm.rfile) Then
            Me.Close()
            Return
        End If

        ' Create directory only if it doesn't already exist
        If Not My.Computer.FileSystem.DirectoryExists(mainForm.rdirectory) Then
            My.Computer.FileSystem.CreateDirectory(mainForm.rdirectory)
        End If

        ' Write header to training run file
        My.Computer.FileSystem.WriteAllText(mainForm.rfile,
                                                "Training Name:" & Strings.Space(7) &
                                                name & ControlChars.NewLine &
                                                "Location:" & Strings.Space(12) &
                                                 location & ControlChars.NewLine &
                                                "Dates:" & Strings.Space(15) &
                                                startDate & Strings.Space(2) &
                                                "-" & Strings.Space(2) &
                                                endDate & ControlChars.NewLine &
                                                ControlChars.NewLine, True)
            mainForm.Separation()

        mainForm.reason = "Initial Balance"
        mainForm.payee = "N/A"
        mainForm.CreateEntry(mainForm.payee, mainForm.reason)
        mainForm.Show()
        Me.Close()

    End Sub

    Private Sub trainingIDForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'sets focus to first text box
        Me.txtTrainingName.Focus()

    End Sub
End Class