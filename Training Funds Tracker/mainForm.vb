'Title                  Training Funds Tracker
'Purpose                To help keep track of funds available for
'                       Training
'Created                December 2009
'Last Updated           Updated March 2014


Option Explicit On
Imports System.Globalization

Public Class mainForm

    Private title As String = "Training Funds Tracker"
    Public rdirectory As String = "C:\Training"
    Public rfile As String = "C:\Training\trainingrun.txt"
    Private newDailyBalance As Decimal
    Private previousBalance As Decimal
    Private heading As String = "Date Entered" & Strings.Space(5) & "Type" & Strings.Space(7) & "Payee" & Strings.Space(22) &
                "Debit(-)" & Strings.Space(7) & "Credit(+)" & Strings.Space(7) & "Balance"


    Private Sub mainForm_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'makes the preview txt box read only
        Me.txtPreview.ReadOnly = True

        'Disable apply calc button till preview is seen
        Me.btnApply.Enabled = False
        Me.ApplyToolStripMenuItem.Enabled = False

        'puts numeric values in the credit and debit txt boxes
        Me.txtDebit.Text = "0.00"
        Me.txtCredit.Text = "0.00"

        'loads the typeComboBox
        Me.cmboxType.Items.Add("Enter No.")
        Me.cmboxType.Items.Add("ATM")
        Me.cmboxType.Items.Add("Debit")
        Me.cmboxType.Items.Add("Dep")
        Me.cmboxType.Items.Add("EFT")
        Me.cmboxType.Items.Add("Wthdrw")
        Me.cmboxType.Items.Add("Trxns")

        Me.cmboxType.SelectedIndex = 0


        'initializes diaglog variables
        Dim button As DialogResult


        'varifies the trainingrun.txt exists else asks if the user wants to create it
        If My.Computer.FileSystem.FileExists(rfile) Then

            'declare block variables
            Dim text As String
            Dim nameIndex As Integer 'numer of chars accessed or what you are putting in txt box
            Dim name As String  'variable to put accessed text into
            Dim NewLineIndex As Integer = 0 'beginning of line
            Dim colonIndex As Integer
            Dim dashIndex As Integer
            Dim location As String
            Dim startDate As String
            Dim endDate As String
            Dim entryIndex As Integer = 0  'length of Line
            Dim entry As String
            Dim myentry As String

            Me.lblPrevBal.Text = My.Computer.FileSystem.ReadAllText(rfile)
            Me.lblNewBal.Text = "0.00"

            'pulls heading information from text file if the file exists, otherwise the program
            'recongnizes it as a new project
            If My.Computer.FileSystem.FileExists(rfile) Then


                text = My.Computer.FileSystem.ReadAllText(rfile) 'pulls and reads entire file

                NewLineIndex = text.IndexOf(ControlChars.NewLine, nameIndex) 'read the beginning of first line 
                colonIndex = text.IndexOf(":", nameIndex) 'gets indexOf ":" to use as marker to find other variables

                name = text.Substring(colonIndex + 1, NewLineIndex - colonIndex) 'loads name variable with substring text before the ":" in first line
                name = name.TrimStart(" ")

                nameIndex = NewLineIndex + 2  'moves to second line

                Me.lblName.Text = name
                Me.lblName.ForeColor = Color.Maroon

                NewLineIndex = text.IndexOf(ControlChars.NewLine, nameIndex)  'read current line
                colonIndex = text.IndexOf(":", nameIndex)

                location = text.Substring(colonIndex + 1, NewLineIndex - colonIndex) 'loads location variable with substring text before the ":" in second line
                location = location.TrimStart(" ")

                nameIndex = NewLineIndex + 2 'moves to the third line

                Me.lblLocation.Text = location
                Me.lblLocation.ForeColor = Color.Maroon

                NewLineIndex = text.IndexOf(ControlChars.NewLine, nameIndex)  'read current line
                colonIndex = text.IndexOf(":", nameIndex)  'get location of ":" for "Date"
                dashIndex = text.IndexOf("-", nameIndex)  'get location of "-" for actual date reads

                startDate = text.Substring(colonIndex + 1, (dashIndex - colonIndex) - 1)  'loads start date with substring after ":" but before "-" in third line
                startDate = startDate.TrimStart(" ")

                Me.lblStartDate.Text = startDate
                Me.lblStartDate.ForeColor = Color.Maroon

                endDate = text.Substring(dashIndex, NewLineIndex - dashIndex)  'loads end date with substring after "-" in third line
                endDate = endDate.TrimStart("-", " ")

                Me.lblEndDate.Text = endDate
                Me.lblEndDate.ForeColor = Color.Maroon

                NewLineIndex = 0 'reset variable before primer and loop before re-run

                NewLineIndex = text.IndexOf(ControlChars.NewLine, entryIndex) 'primer before loop to find balance

                Do Until NewLineIndex = -1

                    entry = text.Substring(entryIndex, NewLineIndex - entryIndex) 'get each line

                    'if line is found with a date, add to entry variable to find bank bal at end
                    If entry.Contains("/") Then
                        myentry = Trim(Microsoft.VisualBasic.Right(entry, 7))
                    End If

                    'Retrieve Current Bank Bal
                    lblPrevBal.Text = myentry

                    'if not found, updates entryindex with next line
                    entryIndex = NewLineIndex + 2
                    NewLineIndex = text.IndexOf(ControlChars.NewLine, entryIndex)
                Loop

            End If


        Else 'sets up the Training folder and creates the trainingrun.txt file
            button = MessageBox.Show _
            ("The current training file does not exist.  This is your bank, would you like to create it?",
            title, MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1)

            'declares another button result and asks the user to enter a beginning balance
            If button = DialogResult.Yes Then

                'loops till the user enters a beginning balance and verifies it is numeric
                Do Until IsNumeric(Me.lblPrevBal.Text) And Me.lblPrevBal.Text <> String.Empty
                    Me.lblPrevBal.Text = InputBox("Please enter starting balance formatted as 0.00.", title, "0.00")
                    If Not IsNumeric(Me.lblPrevBal.Text) Then
                        MessageBox.Show("Number must be numeric.", title, MessageBoxButtons.OK)
                    End If
                Loop

                Me.Show()
                Me.txtPreview.Text = "Ready"
                Me.lblNewBal.Text = "0.00"
                Me.lblName.Text = "New Training"
                Me.lblName.ForeColor = Color.Blue

                'Opens the trainingID form for basic info
                trainingIDForm.ShowDialog()

                'user has decided not to enter a intial balance and closes the form
            Else : button = Windows.Forms.DialogResult.No
                Me.Close()
            End If
        End If
    End Sub

    Private Sub btnClear_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClear.Click

        ClearForm()

    End Sub

    Private Sub btnCalc_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnCalc.Click

        PreviewCalculations()

    End Sub

    Private Sub btnApply_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnApply.Click

        ApplyCalculation()

    End Sub

    Private Sub btnClose_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClose.Click

        CloseApp()

    End Sub

    Private Sub cmboxType_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmboxType.Enter
        Me.cmboxType.SelectAll()
    End Sub

    Private Sub txtPayee_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPayee.Enter
        Me.txtPayee.SelectAll()

    End Sub

    Private Sub txtDebit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDebit.Enter
        Me.txtDebit.SelectAll()

    End Sub

    Private Sub txtCredit_Enter(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCredit.Enter
        Me.txtCredit.SelectAll()

    End Sub

    Private Sub cmboxType_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmboxType.GotFocus
        Me.cmboxType.BackColor = Color.LightBlue
        Me.txtPayee.BackColor = Color.White
        Me.txtDebit.BackColor = Color.White
        Me.txtCredit.BackColor = Color.White
    End Sub

    Private Sub txtPayee_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtPayee.GotFocus
        Me.cmboxType.BackColor = Color.White
        Me.txtPayee.BackColor = Color.LightBlue
        Me.txtDebit.BackColor = Color.White
        Me.txtCredit.BackColor = Color.White
    End Sub

    Private Sub txtDebit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDebit.GotFocus
        Me.cmboxType.BackColor = Color.White
        Me.txtPayee.BackColor = Color.White
        Me.txtDebit.BackColor = Color.LightBlue
        Me.txtCredit.BackColor = Color.White
    End Sub

    Private Sub txtCredit_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtCredit.GotFocus
        Me.cmboxType.BackColor = Color.White
        Me.txtPayee.BackColor = Color.White
        Me.txtDebit.BackColor = Color.White
        Me.txtCredit.BackColor = Color.LightBlue
    End Sub

    Private Sub dtpEntryDate_GotFocus(ByVal sender As Object, ByVal e As System.EventArgs) Handles dtpEntryDate.GotFocus
        Me.cmboxType.BackColor = Color.White
        Me.txtPayee.BackColor = Color.White
        Me.txtDebit.BackColor = Color.White
        Me.txtCredit.BackColor = Color.White
    End Sub

    Private Sub ExitToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ExitToolStripMenuItem.Click

        Me.Close()

    End Sub

    Private Sub ActivitySheetToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ActivitySheetToolStripMenuItem.Click

        If My.Computer.FileSystem.FileExists(rfile) Then
            Dim proc As New System.Diagnostics.Process
            proc.StartInfo.FileName = "notepad.exe"
            proc.StartInfo.Arguments = rfile
            proc.Start()

        Else
            MessageBox.Show("File is in the process of being created on first run.  Please make a calculation and press reply before re-opening.",
            title, MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Me.Show()
            Me.dtpEntryDate.Focus()

        End If
    End Sub

    Private Sub ClearFormToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearFormToolStripMenuItem.Click

        ClearForm()

    End Sub

    Private Sub CalculateToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CalculateToolStripMenuItem.Click

        PreviewCalculations()

    End Sub

    Private Sub ApplyToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ApplyToolStripMenuItem.Click

        ApplyCalculation()

    End Sub

    Private Sub AboutToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles AboutToolStripMenuItem.Click

        trainingAboutBox.ShowDialog()


    End Sub

    Private Sub cmboxType_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmboxType.TextChanged

        If Me.cmboxType.SelectedIndex = 0 Or Me.cmboxType.SelectedIndex = 6 Then
            Me.txtDebit.Enabled = True
            Me.txtCredit.Enabled = True
        End If
        If IsNumeric(Me.cmboxType.Text) Or Me.cmboxType.SelectedIndex = 1 Or Me.cmboxType.SelectedIndex = 2 Or
        Me.cmboxType.SelectedIndex = 4 Or Me.cmboxType.SelectedIndex = 5 Then
            Me.txtCredit.Enabled = False
            Me.txtDebit.Enabled = True
        End If
        If Me.cmboxType.SelectedIndex = 3 Then
            Me.txtDebit.Enabled = False
            Me.txtCredit.Enabled = True
        End If
    End Sub

    Private Sub ReadMeToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ReadMeToolStripMenuItem.Click

        If My.Computer.FileSystem.FileExists("C:\Training\trainingreadme.txt") Then
            Dim proc As New System.Diagnostics.Process
            proc.StartInfo.FileName = "notepad.exe"
            proc.StartInfo.Arguments = "C:\Training\trainingreadme.txt"
            proc.Start()

        Else
            MessageBox.Show("File was not copied manually at install.  Check with deloyment file for trainingreadme.txt file and copy to (C:\Training\) root directory.",
            title, MessageBoxButtons.OK, MessageBoxIcon.Stop)
            Me.Show()
            Me.dtpEntryDate.Focus()

        End If


    End Sub

    Private Sub InfoFormToolStripMenuItem_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        trainingIDForm.ShowDialog()

    End Sub

    Private Sub ClearForm()

        'delcare proceedure variables
        Dim button As DialogResult

        button = MessageBox.Show("Do you wish to add to the new balance to the bank?", title, MessageBoxButtons.YesNo,
        MessageBoxIcon.Question)

        If button = Windows.Forms.DialogResult.Yes Then
            'declare block variables
            Dim curdate As String
            Dim type As String
            Dim payee As String
            Dim previous As String 'formally line2


            'make calculations
            newDailyBalance = Convert.ToDecimal(Me.lblNewBal.Text)
            newDailyBalance = Math.Round(newDailyBalance, 2)
            previousBalance = newDailyBalance
            newDailyBalance = 0D
            Me.lblPrevBal.Text = Convert.ToString(previousBalance)

            'convert data
            type = Me.cmboxType.Text
            payee = Me.txtPayee.Text
            curdate = Me.dtpEntryDate.Text
            previous = Convert.ToString(previousBalance)

            'Write the current balance to the txt file
            If button = Windows.Forms.DialogResult.Yes And My.Computer.FileSystem.FileExists(rfile) Then
                My.Computer.FileSystem.WriteAllText(rfile, curdate _
                & Strings.Space(7) & type.PadRight(7, " ") & Strings.Space(4) & payee.PadRight(20, " ") & Strings.Space(7) & Me.txtDebit.Text.PadRight(6, " ") &
                Strings.Space(9) & Me.txtCredit.Text.PadRight(6, " ") & Strings.Space(10) & Convert.ToString(previousBalance) &
                ControlChars.NewLine, True)
                My.Computer.FileSystem.WriteAllText(rfile, "".PadLeft(105, "-") &
                ControlChars.NewLine, True)



            Else 'set up for the first run
                My.Computer.FileSystem.CreateDirectory(rdirectory)
                My.Computer.FileSystem.WriteAllText(rfile, heading & ControlChars.NewLine _
                & "-------------" & Strings.Space(5) & "-----" & Strings.Space(6) &
                "-------" & Strings.Space(20) & "----------" & Strings.Space(5) &
                "----------" & Strings.Space(6) & "--------" & ControlChars.NewLine, True)
                My.Computer.FileSystem.WriteAllText(rfile, curdate _
                & Strings.Space(7) & type.PadRight(7, " ") & Strings.Space(4) & payee.PadRight(20, " ") & Strings.Space(7) & Me.txtDebit.Text.PadRight(6, " ") &
                Strings.Space(9) & Me.txtCredit.Text.PadRight(6, " ") & Strings.Space(10) & Convert.ToString(previousBalance) &
                ControlChars.NewLine, True)
                My.Computer.FileSystem.WriteAllText(rfile, "".PadLeft(105, "-") &
                ControlChars.NewLine, True)

            End If

            'clears and returns to form
            Me.txtPreview.Text = "Ready"
            Me.lblTransAction.Text = ""
            Me.cmboxType.SelectedIndex = 0
            Me.txtPayee.Clear()
            Me.txtCredit.Text = "0.00"
            Me.txtDebit.Text = "0.00"
            Me.lblNewBal.Text = ""
            Me.dtpEntryDate.Focus()

        Else
            'clears and returns to form
            Me.Show()
            newDailyBalance = 0D
            Me.txtPreview.Text = "Ready"
            Me.lblTransAction.Text = ""
            Me.cmboxType.SelectedIndex = 0
            Me.txtPayee.Clear()
            Me.txtCredit.Text = "0.00"
            Me.txtDebit.Text = "0.00"
            Me.lblNewBal.Text = ""
            Me.dtpEntryDate.Focus()


        End If

    End Sub

    Private Sub ApplyCalculation()

        'saves current balance to the text file

        'declares variables
        Dim addTextButton As DialogResult
        Dim returnButton As DialogResult

        'program asks user if they would like to save the current transaction
        addTextButton = MessageBox.Show("Do you wish to add the new daily balance to the bank?", title,
        MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If addTextButton = Windows.Forms.DialogResult.Yes Then

            'make calculations
            newDailyBalance = Convert.ToDecimal(Me.lblNewBal.Text)
            newDailyBalance = Math.Round(newDailyBalance, 2)
            previousBalance = newDailyBalance
            newDailyBalance = 0D
            Me.lblPrevBal.Text = Convert.ToString(previousBalance)

            'declare text writing variables
            Dim previous As String ' formally line
            Dim curdate As String
            Dim type As String
            Dim payee As String

            'convert the data from the lblPrevBal and store it in the previous variable
            previous = Convert.ToString(previousBalance)
            type = Me.cmboxType.Text
            curdate = Me.dtpEntryDate.Text
            payee = Me.txtPayee.Text

            'If the training file exists, the prog writes current balance to the text file
            If My.Computer.FileSystem.FileExists(rfile) Then
                My.Computer.FileSystem.WriteAllText(rfile, curdate _
                & Strings.Space(7) & type.PadRight(7, " ") & Strings.Space(4) & payee.PadRight(20, " ") & Strings.Space(7) & Me.txtDebit.Text.PadRight(6, " ") &
                Strings.Space(9) & Me.txtCredit.Text.PadRight(6, " ") & Strings.Space(10) & Convert.ToString(previousBalance) &
                ControlChars.NewLine, True)
                My.Computer.FileSystem.WriteAllText(rfile, "".PadLeft(105, "-") &
                ControlChars.NewLine, True)

                MessageBox.Show("Processing complete. The form will be cleared.",
                                title, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Show()
                newDailyBalance = 0D
                Me.txtPreview.Text = "Ready"
                Me.lblTransAction.Text = ""
                Me.lblNewBal.Text = ""
                Me.cmboxType.SelectedIndex = 0
                Me.txtPayee.Clear()
                Me.txtCredit.Text = "0.00"
                Me.txtDebit.Text = "0.00"
                Me.dtpEntryDate.Focus()
                Me.btnApply.Enabled = False

            Else 'set up for the first run
                My.Computer.FileSystem.CreateDirectory(rfile)
                My.Computer.FileSystem.WriteAllText(rfile, heading & ControlChars.NewLine _
                & "------------" & Strings.Space(5) & "-----" & Strings.Space(6) &
                "-------" & Strings.Space(20) & "----------" & Strings.Space(5) &
                "----------" & Strings.Space(6) & "--------" & ControlChars.NewLine, True)
                My.Computer.FileSystem.WriteAllText(rfile, curdate _
                & Strings.Space(7) & type.PadRight(7, " ") & Strings.Space(4) & payee.PadRight(20, " ") & Strings.Space(7) & Me.txtDebit.Text.PadRight(6, " ") &
                Strings.Space(9) & Me.txtCredit.Text.PadRight(6, " ") & Strings.Space(10) & Convert.ToString(previousBalance) &
                ControlChars.NewLine, True)
                My.Computer.FileSystem.WriteAllText(rfile, "".PadLeft(105, "-") &
                ControlChars.NewLine, True)

                MessageBox.Show("Processing complete. The form will be cleared.",
                                title, MessageBoxButtons.OK, MessageBoxIcon.Information)
                Me.Show()
                newDailyBalance = 0D
                Me.txtPreview.Text = "Ready"
                Me.lblTransAction.Text = ""
                Me.lblNewBal.Text = ""
                Me.cmboxType.SelectedIndex = 0
                Me.txtPayee.Clear()
                Me.txtCredit.Text = "0.00"
                Me.txtDebit.Text = "0.00"
                Me.dtpEntryDate.Focus()
                Me.btnApply.Enabled = False

            End If


        Else 'If the user does not want to make a calculation, the program will ask if the user
            'wants to return to the program to perform another calculation
            'if the user does not, then the program will direct the user to exit
            returnButton = MessageBox.Show("Do you wan to make another calculation?", title,
            MessageBoxButtons.YesNo, MessageBoxIcon.Question)
            If returnButton = Windows.Forms.DialogResult.Yes Then

                Me.Show()
                newDailyBalance = 0D
                Me.txtPreview.Text = "Ready"
                Me.lblTransAction.Text = ""
                Me.lblNewBal.Text = ""
                Me.cmboxType.SelectedIndex = 0
                Me.txtPayee.Clear()
                Me.txtCredit.Text = "0.00"
                Me.txtDebit.Text = "0.00"
                Me.dtpEntryDate.Focus()
                Me.btnApply.Enabled = False
            Else
                returnButton = Windows.Forms.DialogResult.No
                MessageBox.Show("No calculation will be made and the form will be reset. You may exit the program.",
                title, MessageBoxButtons.OK, MessageBoxIcon.Information)

                Me.Show()
                newDailyBalance = 0D
                Me.txtPreview.Text = "Ready"
                Me.lblTransAction.Text = ""
                Me.lblNewBal.Text = ""
                Me.cmboxType.SelectedIndex = 0
                Me.txtPayee.Clear()
                Me.txtCredit.Text = "0.00"
                Me.txtDebit.Text = "0.00"
                Me.dtpEntryDate.Focus()
                Me.btnApply.Enabled = False

            End If
        End If

    End Sub

    Private Sub PreviewCalculations()

        'Declare Procedure Level Variables for Calculations
        Dim isAdded As Boolean
        Dim isSubtracted As Boolean
        Dim credit As Decimal
        Dim debit As Decimal
        Dim calcTransactionBal As Decimal
        Dim previewBankBal As Decimal

        'Convert Input
        isAdded = Decimal.TryParse(Me.txtCredit.Text, credit)
        isSubtracted = Decimal.TryParse(Me.txtDebit.Text, debit)


        If IsNumeric(Me.txtCredit.Text) And IsNumeric(Me.txtDebit.Text) Then
            If isAdded And isSubtracted Then

                'Make calculations
                calcTransactionBal = credit - debit
                calcTransactionBal = Math.Round(calcTransactionBal, 2)
                newDailyBalance = Convert.ToDecimal(Me.lblPrevBal.Text) + calcTransactionBal
                Me.lblTransAction.Text = Convert.ToString(calcTransactionBal)
                Me.lblNewBal.Text = Convert.ToString(newDailyBalance)

                'fills preview pane (txtPreview)
                previewBankBal = calcTransactionBal + Convert.ToDecimal(Me.lblPrevBal.Text)
                Me.txtPreview.Text = "Preview of Entry to Activity Sheet:" & ControlChars.NewLine &
                ControlChars.NewLine & "Date Entered" & Strings.Space(6) & "Type" & Strings.Space(10) & "Payee" & Strings.Space(47) &
                "Debit(-)" & Strings.Space(7) & "Credit(+)" & Strings.Space(7) & "Balance" &
                ControlChars.NewLine & "-------------------" & Strings.Space(7) & "----------" & Strings.Space(7) &
                "----------------------------" & Strings.Space(30) & "----------" & Strings.Space(9) &
                "----------" & Strings.Space(10) & "------------" & ControlChars.NewLine & Me.dtpEntryDate.Text.PadRight(10, " ") & Strings.Space(10) &
                Me.cmboxType.Text.PadRight(8, " ") & Strings.Space(5) & Me.txtPayee.Text.PadRight(20, " ") & Strings.Space(26) &
                Me.txtDebit.Text.PadLeft(6, " ") & Strings.Space(10) &
                Me.txtCredit.Text.PadLeft(6, " ") & Strings.Space(10) &
                Convert.ToString(previewBankBal).PadLeft(7)


            Else
                'show error message and highlight the problematic area
                MessageBox.Show("Number for calculations must be numeric.", title,
                MessageBoxButtons.OK, MessageBoxIcon.Information)

                If Not isSubtracted Then

                    Me.txtDebit.Focus()

                ElseIf Not isAdded Then

                    Me.txtCredit.Focus()

                End If

            End If


        Else 'show error message and highlight the problematic area
            MessageBox.Show("Number for calculations must be numeric.", title,
            MessageBoxButtons.OK, MessageBoxIcon.Information)
            If Not IsNumeric(Me.txtDebit.Text) Then

                Me.txtDebit.Focus()

            ElseIf Not IsNumeric(Me.txtCredit.Text) Then

                Me.txtCredit.Focus()

            End If


        End If

        Me.btnApply.Enabled = True
        Me.ApplyToolStripMenuItem.Enabled = True

    End Sub

    Private Sub CloseApp()

        'Exits the Program

        'declare variable
        Dim exitButton As DialogResult

        exitButton = MessageBox.Show("Are you sure that you are ready to exit?", title,
        MessageBoxButtons.YesNo, MessageBoxIcon.Question)
        If exitButton = Windows.Forms.DialogResult.No Then

            'Returns to the form
            Me.Show()
            Me.txtPayee.Clear()
            Me.lblNewBal.Text = "0.00"
            Me.lblTransAction.Text = ""
            Me.txtPreview.Text = "Ready"
            Me.cmboxType.SelectedIndex = 0
            Me.txtDebit.Text = "0.00"
            Me.txtCredit.Text = "0.00"
            newDailyBalance = 0D
            Me.dtpEntryDate.Focus()
            Me.btnApply.Enabled = False

        Else 'Exits the program
            exitButton = Windows.Forms.DialogResult.Yes
            If My.Computer.FileSystem.FileExists(rfile) Then
                Me.Close()

            Else : CreateMyPaths()
                Me.Close()

            End If
        End If
    End Sub
    Public Sub CreateMyPaths()

        'Only used as a placeholder on first run if no prior transactions were completed

        'Declare text writing variables
        Dim entryno As String = "Placeholder"
        Dim curdate As String = dtpEntryDate.Text
        Dim previous As String

        previous = lblPrevBal.Text

        'TODO: Add 2nd Placeholder in order to line up everything. 

        My.Computer.FileSystem.CreateDirectory(rdirectory)
        My.Computer.FileSystem.WriteAllText(rfile, curdate &
                                            Strings.Space(12) & entryno.PadRight(15, " ") & Strings.Space(6) & "0.00".PadLeft(5, " ") &
                                            Strings.Space(13) & "0.00".PadLeft(5, " ") & Strings.Space(13) & Convert.ToString(previous).PadLeft(5, " ") &
                                            ControlChars.NewLine, True)

        My.Computer.FileSystem.WriteAllText(rfile, "".PadLeft(105, "_") & ControlChars.NewLine, True)







    End Sub
End Class
