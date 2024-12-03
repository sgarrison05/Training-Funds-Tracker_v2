<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class mainForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.lbltransDate = New System.Windows.Forms.Label()
        Me.dtpEntryDate = New System.Windows.Forms.DateTimePicker()
        Me.lblPay = New System.Windows.Forms.Label()
        Me.lblDebit = New System.Windows.Forms.Label()
        Me.lblCredit = New System.Windows.Forms.Label()
        Me.txtPayee = New System.Windows.Forms.TextBox()
        Me.txtDebit = New System.Windows.Forms.TextBox()
        Me.txtCredit = New System.Windows.Forms.TextBox()
        Me.btnClear = New System.Windows.Forms.Button()
        Me.btnCalc = New System.Windows.Forms.Button()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.lblPrevBal = New System.Windows.Forms.Label()
        Me.lblTransAction = New System.Windows.Forms.Label()
        Me.lblNewBal = New System.Windows.Forms.Label()
        Me.lblLine = New System.Windows.Forms.Label()
        Me.lblOperator = New System.Windows.Forms.Label()
        Me.txtPreview = New System.Windows.Forms.TextBox()
        Me.lblPreview = New System.Windows.Forms.Label()
        Me.trainingMenuStrip = New System.Windows.Forms.MenuStrip()
        Me.FileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.OpenToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ActivitySheetToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ExitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.EditToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ClearFormToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolsToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.CalculateToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ApplyToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ReadMeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.AboutToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.lblType = New System.Windows.Forms.Label()
        Me.lblPreviousID = New System.Windows.Forms.Label()
        Me.cmboxType = New System.Windows.Forms.ComboBox()
        Me.lblProjectID = New System.Windows.Forms.Label()
        Me.lblName = New System.Windows.Forms.Label()
        Me.lblLocationID = New System.Windows.Forms.Label()
        Me.lblStartDateID = New System.Windows.Forms.Label()
        Me.lblEndDateID = New System.Windows.Forms.Label()
        Me.lblLineID = New System.Windows.Forms.Label()
        Me.lblLocation = New System.Windows.Forms.Label()
        Me.lblStartDate = New System.Windows.Forms.Label()
        Me.lblEndDate = New System.Windows.Forms.Label()
        Me.btnReconcile = New System.Windows.Forms.Button()
        Me.trainingMenuStrip.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbltransDate
        '
        Me.lbltransDate.AutoSize = True
        Me.lbltransDate.Location = New System.Drawing.Point(9, 84)
        Me.lbltransDate.Name = "lbltransDate"
        Me.lbltransDate.Size = New System.Drawing.Size(33, 13)
        Me.lbltransDate.TabIndex = 0
        Me.lbltransDate.Text = "Date:"
        '
        'dtpEntryDate
        '
        Me.dtpEntryDate.Checked = False
        Me.dtpEntryDate.CustomFormat = "MM/dd/yyyy"
        Me.dtpEntryDate.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpEntryDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpEntryDate.Location = New System.Drawing.Point(12, 100)
        Me.dtpEntryDate.Name = "dtpEntryDate"
        Me.dtpEntryDate.Size = New System.Drawing.Size(111, 26)
        Me.dtpEntryDate.TabIndex = 1
        '
        'lblPay
        '
        Me.lblPay.AutoSize = True
        Me.lblPay.Location = New System.Drawing.Point(154, 144)
        Me.lblPay.Name = "lblPay"
        Me.lblPay.Size = New System.Drawing.Size(99, 13)
        Me.lblPay.TabIndex = 4
        Me.lblPay.Text = "Pay to the Order of:"
        '
        'lblDebit
        '
        Me.lblDebit.AutoSize = True
        Me.lblDebit.Location = New System.Drawing.Point(476, 144)
        Me.lblDebit.Name = "lblDebit"
        Me.lblDebit.Size = New System.Drawing.Size(47, 13)
        Me.lblDebit.TabIndex = 6
        Me.lblDebit.Text = "Debit (-):"
        '
        'lblCredit
        '
        Me.lblCredit.AutoSize = True
        Me.lblCredit.Location = New System.Drawing.Point(564, 144)
        Me.lblCredit.Name = "lblCredit"
        Me.lblCredit.Size = New System.Drawing.Size(52, 13)
        Me.lblCredit.TabIndex = 8
        Me.lblCredit.Text = "Credit (+):"
        '
        'txtPayee
        '
        Me.txtPayee.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPayee.Location = New System.Drawing.Point(157, 160)
        Me.txtPayee.Name = "txtPayee"
        Me.txtPayee.Size = New System.Drawing.Size(300, 26)
        Me.txtPayee.TabIndex = 5
        '
        'txtDebit
        '
        Me.txtDebit.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtDebit.Location = New System.Drawing.Point(479, 160)
        Me.txtDebit.Name = "txtDebit"
        Me.txtDebit.Size = New System.Drawing.Size(74, 26)
        Me.txtDebit.TabIndex = 7
        '
        'txtCredit
        '
        Me.txtCredit.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtCredit.Location = New System.Drawing.Point(567, 160)
        Me.txtCredit.Name = "txtCredit"
        Me.txtCredit.Size = New System.Drawing.Size(74, 26)
        Me.txtCredit.TabIndex = 9
        '
        'btnClear
        '
        Me.btnClear.BackColor = System.Drawing.Color.Cyan
        Me.btnClear.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClear.Location = New System.Drawing.Point(12, 403)
        Me.btnClear.Name = "btnClear"
        Me.btnClear.Size = New System.Drawing.Size(86, 41)
        Me.btnClear.TabIndex = 10
        Me.btnClear.Text = "Clear"
        Me.btnClear.UseVisualStyleBackColor = False
        '
        'btnCalc
        '
        Me.btnCalc.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnCalc.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnCalc.Location = New System.Drawing.Point(180, 403)
        Me.btnCalc.Name = "btnCalc"
        Me.btnCalc.Size = New System.Drawing.Size(86, 41)
        Me.btnCalc.TabIndex = 11
        Me.btnCalc.Text = "Calculate"
        Me.btnCalc.UseVisualStyleBackColor = False
        '
        'btnApply
        '
        Me.btnApply.BackColor = System.Drawing.Color.FromArgb(CType(CType(0, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(0, Byte), Integer))
        Me.btnApply.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnApply.Location = New System.Drawing.Point(361, 403)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(86, 41)
        Me.btnApply.TabIndex = 12
        Me.btnApply.Text = "Apply"
        Me.btnApply.UseVisualStyleBackColor = False
        '
        'btnClose
        '
        Me.btnClose.BackColor = System.Drawing.Color.FromArgb(CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.btnClose.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnClose.Location = New System.Drawing.Point(717, 403)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(86, 41)
        Me.btnClose.TabIndex = 13
        Me.btnClose.Text = "Close"
        Me.btnClose.UseVisualStyleBackColor = False
        '
        'lblPrevBal
        '
        Me.lblPrevBal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblPrevBal.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPrevBal.Location = New System.Drawing.Point(703, 122)
        Me.lblPrevBal.Name = "lblPrevBal"
        Me.lblPrevBal.Size = New System.Drawing.Size(100, 23)
        Me.lblPrevBal.TabIndex = 17
        '
        'lblTransAction
        '
        Me.lblTransAction.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblTransAction.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTransAction.Location = New System.Drawing.Point(703, 162)
        Me.lblTransAction.Name = "lblTransAction"
        Me.lblTransAction.Size = New System.Drawing.Size(100, 23)
        Me.lblTransAction.TabIndex = 18
        '
        'lblNewBal
        '
        Me.lblNewBal.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblNewBal.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNewBal.Location = New System.Drawing.Point(703, 210)
        Me.lblNewBal.Name = "lblNewBal"
        Me.lblNewBal.Size = New System.Drawing.Size(100, 23)
        Me.lblNewBal.TabIndex = 19
        '
        'lblLine
        '
        Me.lblLine.BackColor = System.Drawing.SystemColors.ActiveCaption
        Me.lblLine.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.lblLine.Location = New System.Drawing.Point(704, 194)
        Me.lblLine.Name = "lblLine"
        Me.lblLine.Size = New System.Drawing.Size(100, 2)
        Me.lblLine.TabIndex = 20
        '
        'lblOperator
        '
        Me.lblOperator.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblOperator.Location = New System.Drawing.Point(647, 163)
        Me.lblOperator.Name = "lblOperator"
        Me.lblOperator.Size = New System.Drawing.Size(50, 23)
        Me.lblOperator.TabIndex = 15
        Me.lblOperator.Text = "(+ / -)"
        Me.lblOperator.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        '
        'txtPreview
        '
        Me.txtPreview.BackColor = System.Drawing.Color.Black
        Me.txtPreview.Font = New System.Drawing.Font("Arial", 9.75!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtPreview.ForeColor = System.Drawing.Color.White
        Me.txtPreview.Location = New System.Drawing.Point(12, 267)
        Me.txtPreview.Multiline = True
        Me.txtPreview.Name = "txtPreview"
        Me.txtPreview.Size = New System.Drawing.Size(791, 113)
        Me.txtPreview.TabIndex = 22
        '
        'lblPreview
        '
        Me.lblPreview.AutoSize = True
        Me.lblPreview.Location = New System.Drawing.Point(9, 251)
        Me.lblPreview.Name = "lblPreview"
        Me.lblPreview.Size = New System.Drawing.Size(119, 13)
        Me.lblPreview.TabIndex = 21
        Me.lblPreview.Text = "Preview of Transaction:"
        '
        'trainingMenuStrip
        '
        Me.trainingMenuStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.FileToolStripMenuItem, Me.EditToolStripMenuItem, Me.ToolsToolStripMenuItem, Me.HelpToolStripMenuItem})
        Me.trainingMenuStrip.Location = New System.Drawing.Point(0, 0)
        Me.trainingMenuStrip.Name = "trainingMenuStrip"
        Me.trainingMenuStrip.Size = New System.Drawing.Size(825, 24)
        Me.trainingMenuStrip.TabIndex = 14
        Me.trainingMenuStrip.Text = "trainingMenuStrip"
        '
        'FileToolStripMenuItem
        '
        Me.FileToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewToolStripMenuItem, Me.OpenToolStripMenuItem, Me.ExitToolStripMenuItem})
        Me.FileToolStripMenuItem.Name = "FileToolStripMenuItem"
        Me.FileToolStripMenuItem.Size = New System.Drawing.Size(37, 20)
        Me.FileToolStripMenuItem.Text = "File"
        '
        'NewToolStripMenuItem
        '
        Me.NewToolStripMenuItem.Name = "NewToolStripMenuItem"
        Me.NewToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.NewToolStripMenuItem.Text = "New Training..."
        '
        'OpenToolStripMenuItem
        '
        Me.OpenToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ActivitySheetToolStripMenuItem})
        Me.OpenToolStripMenuItem.Name = "OpenToolStripMenuItem"
        Me.OpenToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.OpenToolStripMenuItem.Text = "Open..."
        '
        'ActivitySheetToolStripMenuItem
        '
        Me.ActivitySheetToolStripMenuItem.Name = "ActivitySheetToolStripMenuItem"
        Me.ActivitySheetToolStripMenuItem.Size = New System.Drawing.Size(146, 22)
        Me.ActivitySheetToolStripMenuItem.Text = "Activity Sheet"
        '
        'ExitToolStripMenuItem
        '
        Me.ExitToolStripMenuItem.Name = "ExitToolStripMenuItem"
        Me.ExitToolStripMenuItem.Size = New System.Drawing.Size(152, 22)
        Me.ExitToolStripMenuItem.Text = "Exit"
        '
        'EditToolStripMenuItem
        '
        Me.EditToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ClearFormToolStripMenuItem})
        Me.EditToolStripMenuItem.Name = "EditToolStripMenuItem"
        Me.EditToolStripMenuItem.Size = New System.Drawing.Size(39, 20)
        Me.EditToolStripMenuItem.Text = "Edit"
        '
        'ClearFormToolStripMenuItem
        '
        Me.ClearFormToolStripMenuItem.Name = "ClearFormToolStripMenuItem"
        Me.ClearFormToolStripMenuItem.Size = New System.Drawing.Size(132, 22)
        Me.ClearFormToolStripMenuItem.Text = "Clear Form"
        '
        'ToolsToolStripMenuItem
        '
        Me.ToolsToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CalculateToolStripMenuItem, Me.ApplyToolStripMenuItem})
        Me.ToolsToolStripMenuItem.Name = "ToolsToolStripMenuItem"
        Me.ToolsToolStripMenuItem.Size = New System.Drawing.Size(46, 20)
        Me.ToolsToolStripMenuItem.Text = "Tools"
        '
        'CalculateToolStripMenuItem
        '
        Me.CalculateToolStripMenuItem.Name = "CalculateToolStripMenuItem"
        Me.CalculateToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
        Me.CalculateToolStripMenuItem.Text = "Calculate"
        '
        'ApplyToolStripMenuItem
        '
        Me.ApplyToolStripMenuItem.Name = "ApplyToolStripMenuItem"
        Me.ApplyToolStripMenuItem.Size = New System.Drawing.Size(123, 22)
        Me.ApplyToolStripMenuItem.Text = "Apply"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ReadMeToolStripMenuItem, Me.AboutToolStripMenuItem})
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(44, 20)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'ReadMeToolStripMenuItem
        '
        Me.ReadMeToolStripMenuItem.Name = "ReadMeToolStripMenuItem"
        Me.ReadMeToolStripMenuItem.Size = New System.Drawing.Size(120, 22)
        Me.ReadMeToolStripMenuItem.Text = "Read Me"
        '
        'AboutToolStripMenuItem
        '
        Me.AboutToolStripMenuItem.Name = "AboutToolStripMenuItem"
        Me.AboutToolStripMenuItem.Size = New System.Drawing.Size(120, 22)
        Me.AboutToolStripMenuItem.Text = "About"
        '
        'lblType
        '
        Me.lblType.AutoSize = True
        Me.lblType.Location = New System.Drawing.Point(9, 144)
        Me.lblType.Name = "lblType"
        Me.lblType.Size = New System.Drawing.Size(87, 13)
        Me.lblType.TabIndex = 2
        Me.lblType.Text = "Type/Check No:"
        '
        'lblPreviousID
        '
        Me.lblPreviousID.AutoSize = True
        Me.lblPreviousID.Location = New System.Drawing.Point(701, 101)
        Me.lblPreviousID.Name = "lblPreviousID"
        Me.lblPreviousID.Size = New System.Drawing.Size(35, 13)
        Me.lblPreviousID.TabIndex = 16
        Me.lblPreviousID.Text = "Bank:"
        '
        'cmboxType
        '
        Me.cmboxType.Font = New System.Drawing.Font("Arial", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.cmboxType.FormattingEnabled = True
        Me.cmboxType.Location = New System.Drawing.Point(12, 160)
        Me.cmboxType.Name = "cmboxType"
        Me.cmboxType.Size = New System.Drawing.Size(121, 26)
        Me.cmboxType.TabIndex = 3
        '
        'lblProjectID
        '
        Me.lblProjectID.AutoSize = True
        Me.lblProjectID.Location = New System.Drawing.Point(9, 33)
        Me.lblProjectID.Name = "lblProjectID"
        Me.lblProjectID.Size = New System.Drawing.Size(80, 13)
        Me.lblProjectID.TabIndex = 23
        Me.lblProjectID.Text = "Current Project:"
        '
        'lblName
        '
        Me.lblName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblName.Location = New System.Drawing.Point(12, 49)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(241, 23)
        Me.lblName.TabIndex = 24
        '
        'lblLocationID
        '
        Me.lblLocationID.AutoSize = True
        Me.lblLocationID.Location = New System.Drawing.Point(274, 32)
        Me.lblLocationID.Name = "lblLocationID"
        Me.lblLocationID.Size = New System.Drawing.Size(51, 13)
        Me.lblLocationID.TabIndex = 25
        Me.lblLocationID.Text = "Location:"
        '
        'lblStartDateID
        '
        Me.lblStartDateID.AutoSize = True
        Me.lblStartDateID.Location = New System.Drawing.Point(476, 33)
        Me.lblStartDateID.Name = "lblStartDateID"
        Me.lblStartDateID.Size = New System.Drawing.Size(58, 13)
        Me.lblStartDateID.TabIndex = 26
        Me.lblStartDateID.Text = "Start Date:"
        '
        'lblEndDateID
        '
        Me.lblEndDateID.AutoSize = True
        Me.lblEndDateID.Location = New System.Drawing.Point(617, 33)
        Me.lblEndDateID.Name = "lblEndDateID"
        Me.lblEndDateID.Size = New System.Drawing.Size(55, 13)
        Me.lblEndDateID.TabIndex = 27
        Me.lblEndDateID.Text = "End Date:"
        '
        'lblLineID
        '
        Me.lblLineID.BackColor = System.Drawing.SystemColors.ControlDarkDark
        Me.lblLineID.Location = New System.Drawing.Point(595, 60)
        Me.lblLineID.Name = "lblLineID"
        Me.lblLineID.Size = New System.Drawing.Size(10, 1)
        Me.lblLineID.TabIndex = 28
        '
        'lblLocation
        '
        Me.lblLocation.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblLocation.Location = New System.Drawing.Point(277, 49)
        Me.lblLocation.Name = "lblLocation"
        Me.lblLocation.Size = New System.Drawing.Size(170, 24)
        Me.lblLocation.TabIndex = 29
        '
        'lblStartDate
        '
        Me.lblStartDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblStartDate.Location = New System.Drawing.Point(479, 49)
        Me.lblStartDate.Name = "lblStartDate"
        Me.lblStartDate.Size = New System.Drawing.Size(103, 23)
        Me.lblStartDate.TabIndex = 30
        '
        'lblEndDate
        '
        Me.lblEndDate.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblEndDate.Location = New System.Drawing.Point(620, 50)
        Me.lblEndDate.Name = "lblEndDate"
        Me.lblEndDate.Size = New System.Drawing.Size(103, 23)
        Me.lblEndDate.TabIndex = 31
        '
        'btnReconcile
        '
        Me.btnReconcile.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnReconcile.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnReconcile.Location = New System.Drawing.Point(541, 403)
        Me.btnReconcile.Name = "btnReconcile"
        Me.btnReconcile.Size = New System.Drawing.Size(86, 41)
        Me.btnReconcile.TabIndex = 32
        Me.btnReconcile.Text = "Reconcile"
        Me.btnReconcile.UseVisualStyleBackColor = False
        '
        'mainForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(825, 477)
        Me.Controls.Add(Me.btnReconcile)
        Me.Controls.Add(Me.lblEndDate)
        Me.Controls.Add(Me.lblStartDate)
        Me.Controls.Add(Me.lblLocation)
        Me.Controls.Add(Me.lblLineID)
        Me.Controls.Add(Me.lblEndDateID)
        Me.Controls.Add(Me.lblStartDateID)
        Me.Controls.Add(Me.lblLocationID)
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.lblProjectID)
        Me.Controls.Add(Me.cmboxType)
        Me.Controls.Add(Me.lblPreviousID)
        Me.Controls.Add(Me.lblType)
        Me.Controls.Add(Me.lblPreview)
        Me.Controls.Add(Me.txtPreview)
        Me.Controls.Add(Me.lblOperator)
        Me.Controls.Add(Me.lblLine)
        Me.Controls.Add(Me.lblNewBal)
        Me.Controls.Add(Me.lblTransAction)
        Me.Controls.Add(Me.lblPrevBal)
        Me.Controls.Add(Me.btnClose)
        Me.Controls.Add(Me.btnApply)
        Me.Controls.Add(Me.btnCalc)
        Me.Controls.Add(Me.btnClear)
        Me.Controls.Add(Me.txtCredit)
        Me.Controls.Add(Me.txtDebit)
        Me.Controls.Add(Me.txtPayee)
        Me.Controls.Add(Me.lblCredit)
        Me.Controls.Add(Me.lblDebit)
        Me.Controls.Add(Me.lblPay)
        Me.Controls.Add(Me.dtpEntryDate)
        Me.Controls.Add(Me.lbltransDate)
        Me.Controls.Add(Me.trainingMenuStrip)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MainMenuStrip = Me.trainingMenuStrip
        Me.Name = "mainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Training Funds Tracker"
        Me.trainingMenuStrip.ResumeLayout(False)
        Me.trainingMenuStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lbltransDate As System.Windows.Forms.Label
    Friend WithEvents dtpEntryDate As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblPay As System.Windows.Forms.Label
    Friend WithEvents lblDebit As System.Windows.Forms.Label
    Friend WithEvents lblCredit As System.Windows.Forms.Label
    Friend WithEvents txtPayee As System.Windows.Forms.TextBox
    Friend WithEvents txtDebit As System.Windows.Forms.TextBox
    Friend WithEvents txtCredit As System.Windows.Forms.TextBox
    Friend WithEvents btnClear As System.Windows.Forms.Button
    Friend WithEvents btnCalc As System.Windows.Forms.Button
    Friend WithEvents btnApply As System.Windows.Forms.Button
    Friend WithEvents btnClose As System.Windows.Forms.Button
    Friend WithEvents lblPrevBal As System.Windows.Forms.Label
    Friend WithEvents lblTransAction As System.Windows.Forms.Label
    Friend WithEvents lblNewBal As System.Windows.Forms.Label
    Friend WithEvents lblLine As System.Windows.Forms.Label
    Friend WithEvents lblOperator As System.Windows.Forms.Label
    Friend WithEvents txtPreview As System.Windows.Forms.TextBox
    Friend WithEvents lblPreview As System.Windows.Forms.Label
    Friend WithEvents trainingMenuStrip As System.Windows.Forms.MenuStrip
    Friend WithEvents FileToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents EditToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolsToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblType As System.Windows.Forms.Label
    Friend WithEvents lblPreviousID As System.Windows.Forms.Label
    Friend WithEvents cmboxType As System.Windows.Forms.ComboBox
    Friend WithEvents OpenToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ExitToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ClearFormToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents CalculateToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ApplyToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ReadMeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents AboutToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ActivitySheetToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents lblProjectID As System.Windows.Forms.Label
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents lblLocationID As System.Windows.Forms.Label
    Friend WithEvents lblStartDateID As System.Windows.Forms.Label
    Friend WithEvents lblEndDateID As System.Windows.Forms.Label
    Friend WithEvents lblLineID As System.Windows.Forms.Label
    Friend WithEvents lblLocation As System.Windows.Forms.Label
    Friend WithEvents lblStartDate As System.Windows.Forms.Label
    Friend WithEvents lblEndDate As System.Windows.Forms.Label
    Friend WithEvents NewToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents btnReconcile As Button
End Class
