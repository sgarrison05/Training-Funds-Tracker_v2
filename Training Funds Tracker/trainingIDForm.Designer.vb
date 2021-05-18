<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class trainingIDForm
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
        Me.dtpTrainingStart = New System.Windows.Forms.DateTimePicker
        Me.dtpTrainingEnd = New System.Windows.Forms.DateTimePicker
        Me.lblTrainStart = New System.Windows.Forms.Label
        Me.lblTrainEnd = New System.Windows.Forms.Label
        Me.lblTrainingName = New System.Windows.Forms.Label
        Me.lblTrainLocation = New System.Windows.Forms.Label
        Me.txtTrainingName = New System.Windows.Forms.TextBox
        Me.txtTrainingLocation = New System.Windows.Forms.TextBox
        Me.btnDone = New System.Windows.Forms.Button
        Me.SuspendLayout()
        '
        'dtpTrainingStart
        '
        Me.dtpTrainingStart.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpTrainingStart.Checked = False
        Me.dtpTrainingStart.CustomFormat = "MM/dd/yyyy"
        Me.dtpTrainingStart.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpTrainingStart.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTrainingStart.Location = New System.Drawing.Point(125, 117)
        Me.dtpTrainingStart.Name = "dtpTrainingStart"
        Me.dtpTrainingStart.Size = New System.Drawing.Size(122, 26)
        Me.dtpTrainingStart.TabIndex = 5
        '
        'dtpTrainingEnd
        '
        Me.dtpTrainingEnd.CalendarFont = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpTrainingEnd.Checked = False
        Me.dtpTrainingEnd.CustomFormat = "MM/dd/yyyy"
        Me.dtpTrainingEnd.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dtpTrainingEnd.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.dtpTrainingEnd.Location = New System.Drawing.Point(125, 158)
        Me.dtpTrainingEnd.Name = "dtpTrainingEnd"
        Me.dtpTrainingEnd.Size = New System.Drawing.Size(122, 26)
        Me.dtpTrainingEnd.TabIndex = 7
        '
        'lblTrainStart
        '
        Me.lblTrainStart.AutoSize = True
        Me.lblTrainStart.Location = New System.Drawing.Point(27, 130)
        Me.lblTrainStart.Name = "lblTrainStart"
        Me.lblTrainStart.Size = New System.Drawing.Size(58, 13)
        Me.lblTrainStart.TabIndex = 4
        Me.lblTrainStart.Text = "Start Date:"
        '
        'lblTrainEnd
        '
        Me.lblTrainEnd.AutoSize = True
        Me.lblTrainEnd.Location = New System.Drawing.Point(27, 171)
        Me.lblTrainEnd.Name = "lblTrainEnd"
        Me.lblTrainEnd.Size = New System.Drawing.Size(55, 13)
        Me.lblTrainEnd.TabIndex = 6
        Me.lblTrainEnd.Text = "End Date:"
        '
        'lblTrainingName
        '
        Me.lblTrainingName.AutoSize = True
        Me.lblTrainingName.Location = New System.Drawing.Point(27, 37)
        Me.lblTrainingName.Name = "lblTrainingName"
        Me.lblTrainingName.Size = New System.Drawing.Size(79, 13)
        Me.lblTrainingName.TabIndex = 0
        Me.lblTrainingName.Text = "Training Name:"
        '
        'lblTrainLocation
        '
        Me.lblTrainLocation.AutoSize = True
        Me.lblTrainLocation.Location = New System.Drawing.Point(27, 74)
        Me.lblTrainLocation.Name = "lblTrainLocation"
        Me.lblTrainLocation.Size = New System.Drawing.Size(51, 13)
        Me.lblTrainLocation.TabIndex = 2
        Me.lblTrainLocation.Text = "Location:"
        '
        'txtTrainingName
        '
        Me.txtTrainingName.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTrainingName.Location = New System.Drawing.Point(125, 34)
        Me.txtTrainingName.Name = "txtTrainingName"
        Me.txtTrainingName.Size = New System.Drawing.Size(221, 26)
        Me.txtTrainingName.TabIndex = 1
        '
        'txtTrainingLocation
        '
        Me.txtTrainingLocation.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtTrainingLocation.Location = New System.Drawing.Point(125, 74)
        Me.txtTrainingLocation.Name = "txtTrainingLocation"
        Me.txtTrainingLocation.Size = New System.Drawing.Size(199, 26)
        Me.txtTrainingLocation.TabIndex = 3
        '
        'btnDone
        '
        Me.btnDone.Location = New System.Drawing.Point(255, 228)
        Me.btnDone.Name = "btnDone"
        Me.btnDone.Size = New System.Drawing.Size(91, 33)
        Me.btnDone.TabIndex = 8
        Me.btnDone.Text = "Done"
        Me.btnDone.UseVisualStyleBackColor = True
        '
        'trainingIDForm
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(371, 273)
        Me.Controls.Add(Me.btnDone)
        Me.Controls.Add(Me.txtTrainingLocation)
        Me.Controls.Add(Me.txtTrainingName)
        Me.Controls.Add(Me.lblTrainLocation)
        Me.Controls.Add(Me.lblTrainingName)
        Me.Controls.Add(Me.lblTrainEnd)
        Me.Controls.Add(Me.lblTrainStart)
        Me.Controls.Add(Me.dtpTrainingEnd)
        Me.Controls.Add(Me.dtpTrainingStart)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Name = "trainingIDForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Training Tracker Basic Info"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents dtpTrainingStart As System.Windows.Forms.DateTimePicker
    Friend WithEvents dtpTrainingEnd As System.Windows.Forms.DateTimePicker
    Friend WithEvents lblTrainStart As System.Windows.Forms.Label
    Friend WithEvents lblTrainEnd As System.Windows.Forms.Label
    Friend WithEvents lblTrainingName As System.Windows.Forms.Label
    Friend WithEvents lblTrainLocation As System.Windows.Forms.Label
    Friend WithEvents txtTrainingName As System.Windows.Forms.TextBox
    Friend WithEvents txtTrainingLocation As System.Windows.Forms.TextBox
    Friend WithEvents btnDone As System.Windows.Forms.Button
End Class
