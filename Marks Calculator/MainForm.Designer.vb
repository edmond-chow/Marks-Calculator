<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class FrmMain
    Inherits MetroFramework.Forms.MetroForm

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmMain))
        Me.PrbMain = New MetroFramework.Controls.MetroProgressBar()
        Me.LblName = New MetroFramework.Controls.MetroLabel()
        Me.TxtName = New MetroFramework.Controls.MetroTextBox()
        Me.GrpInput = New System.Windows.Forms.GroupBox()
        Me.LblInputMain = New MetroFramework.Controls.MetroLabel()
        Me.LblInputTest = New MetroFramework.Controls.MetroLabel()
        Me.TxtInputTest = New MetroFramework.Controls.MetroTextBox()
        Me.LblInputQuizzes = New MetroFramework.Controls.MetroLabel()
        Me.TxtInputQuizzes = New MetroFramework.Controls.MetroTextBox()
        Me.LblInputProject = New MetroFramework.Controls.MetroLabel()
        Me.TxtInputProject = New MetroFramework.Controls.MetroTextBox()
        Me.LblInputExam = New MetroFramework.Controls.MetroLabel()
        Me.TxtInputExam = New MetroFramework.Controls.MetroTextBox()
        Me.GrpResult = New System.Windows.Forms.GroupBox()
        Me.LblResultCA = New MetroFramework.Controls.MetroLabel()
        Me.TxtResultCA = New MetroFramework.Controls.MetroTextBox()
        Me.LblResultModule = New MetroFramework.Controls.MetroLabel()
        Me.TxtResultModule = New MetroFramework.Controls.MetroTextBox()
        Me.LblResultGrade = New MetroFramework.Controls.MetroLabel()
        Me.TxtReusltGrade = New MetroFramework.Controls.MetroTextBox()
        Me.LblReusltRemarks = New MetroFramework.Controls.MetroLabel()
        Me.TxtReusltRemarks = New MetroFramework.Controls.MetroTextBox()
        Me.GrpStatistics = New System.Windows.Forms.GroupBox()
        Me.LblStatisticsMd = New MetroFramework.Controls.MetroLabel()
        Me.TxtStatisticsMd = New MetroFramework.Controls.MetroTextBox()
        Me.LblStatisticsNo = New MetroFramework.Controls.MetroLabel()
        Me.TxtStatisticsNo = New MetroFramework.Controls.MetroTextBox()
        Me.LblStatisticsA = New MetroFramework.Controls.MetroLabel()
        Me.TxtStatisticsA = New MetroFramework.Controls.MetroTextBox()
        Me.LblStatisticsB = New MetroFramework.Controls.MetroLabel()
        Me.TxtStatisticsB = New MetroFramework.Controls.MetroTextBox()
        Me.LblStatisticsC = New MetroFramework.Controls.MetroLabel()
        Me.TxtStatisticsC = New MetroFramework.Controls.MetroTextBox()
        Me.LblStatisticsF = New MetroFramework.Controls.MetroLabel()
        Me.TxtStatisticsF = New MetroFramework.Controls.MetroTextBox()
        Me.LblStatisticsAv = New MetroFramework.Controls.MetroLabel()
        Me.TxtStatisticsAv = New MetroFramework.Controls.MetroTextBox()
        Me.LblStatisticsSd = New MetroFramework.Controls.MetroLabel()
        Me.TxtStatisticsSd = New MetroFramework.Controls.MetroTextBox()
        Me.GrpRecords = New System.Windows.Forms.GroupBox()
        Me.ChkRecords = New MetroFramework.Controls.MetroCheckBox()
        Me.BtnRecordsAdd = New MetroFramework.Controls.MetroButton()
        Me.BtnRecordsRemove = New MetroFramework.Controls.MetroButton()
        Me.LstRecords = New System.Windows.Forms.ListBox()
        Me.TxtRecordsSearch = New MetroFramework.Controls.MetroTextBox()
        Me.ChkRecordsSearch = New MetroFramework.Controls.MetroCheckBox()
        Me.TmrMain = New System.Windows.Forms.Timer(Me.components)
        Me.GrpInput.SuspendLayout()
        Me.GrpResult.SuspendLayout()
        Me.GrpStatistics.SuspendLayout()
        Me.GrpRecords.SuspendLayout()
        Me.SuspendLayout()
        '
        'PrbMain
        '
        Me.PrbMain.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.PrbMain.Location = New System.Drawing.Point(0, 0)
        Me.PrbMain.Name = "PrbMain"
        Me.PrbMain.Size = New System.Drawing.Size(778, 5)
        Me.PrbMain.TabIndex = 0
        '
        'LblName
        '
        Me.LblName.AutoSize = True
        Me.LblName.Location = New System.Drawing.Point(29, 63)
        Me.LblName.Name = "LblName"
        Me.LblName.Size = New System.Drawing.Size(45, 19)
        Me.LblName.TabIndex = 0
        Me.LblName.Text = "Name"
        '
        'TxtName
        '
        Me.TxtName.Enabled = False
        Me.TxtName.Location = New System.Drawing.Point(80, 63)
        Me.TxtName.Name = "TxtName"
        Me.TxtName.ReadOnly = True
        Me.TxtName.Size = New System.Drawing.Size(200, 21)
        Me.TxtName.TabIndex = 1
        '
        'GrpInput
        '
        Me.GrpInput.Controls.Add(Me.LblInputMain)
        Me.GrpInput.Controls.Add(Me.LblInputTest)
        Me.GrpInput.Controls.Add(Me.TxtInputTest)
        Me.GrpInput.Controls.Add(Me.LblInputQuizzes)
        Me.GrpInput.Controls.Add(Me.TxtInputQuizzes)
        Me.GrpInput.Controls.Add(Me.LblInputProject)
        Me.GrpInput.Controls.Add(Me.TxtInputProject)
        Me.GrpInput.Controls.Add(Me.LblInputExam)
        Me.GrpInput.Controls.Add(Me.TxtInputExam)
        Me.GrpInput.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.GrpInput.Location = New System.Drawing.Point(23, 90)
        Me.GrpInput.Name = "GrpInput"
        Me.GrpInput.Size = New System.Drawing.Size(500, 100)
        Me.GrpInput.TabIndex = 2
        Me.GrpInput.TabStop = False
        Me.GrpInput.Text = "Input CA and Exam Marks"
        '
        'LblInputMain
        '
        Me.LblInputMain.AutoSize = True
        Me.LblInputMain.Location = New System.Drawing.Point(6, 19)
        Me.LblInputMain.Name = "LblInputMain"
        Me.LblInputMain.Size = New System.Drawing.Size(114, 19)
        Me.LblInputMain.TabIndex = 3
        Me.LblInputMain.Text = "[CA Components]"
        '
        'LblInputTest
        '
        Me.LblInputTest.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblInputTest.AutoSize = True
        Me.LblInputTest.Location = New System.Drawing.Point(6, 47)
        Me.LblInputTest.Name = "LblInputTest"
        Me.LblInputTest.Size = New System.Drawing.Size(30, 19)
        Me.LblInputTest.TabIndex = 0
        Me.LblInputTest.Text = "Test"
        '
        'TxtInputTest
        '
        Me.TxtInputTest.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtInputTest.Enabled = False
        Me.TxtInputTest.Location = New System.Drawing.Point(62, 46)
        Me.TxtInputTest.Name = "TxtInputTest"
        Me.TxtInputTest.ReadOnly = True
        Me.TxtInputTest.Size = New System.Drawing.Size(125, 21)
        Me.TxtInputTest.TabIndex = 1
        Me.TxtInputTest.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblInputQuizzes
        '
        Me.LblInputQuizzes.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblInputQuizzes.AutoSize = True
        Me.LblInputQuizzes.Location = New System.Drawing.Point(193, 47)
        Me.LblInputQuizzes.Name = "LblInputQuizzes"
        Me.LblInputQuizzes.Size = New System.Drawing.Size(54, 19)
        Me.LblInputQuizzes.TabIndex = 0
        Me.LblInputQuizzes.Text = "Quizzes"
        '
        'TxtInputQuizzes
        '
        Me.TxtInputQuizzes.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtInputQuizzes.Enabled = False
        Me.TxtInputQuizzes.Location = New System.Drawing.Point(253, 46)
        Me.TxtInputQuizzes.Name = "TxtInputQuizzes"
        Me.TxtInputQuizzes.ReadOnly = True
        Me.TxtInputQuizzes.Size = New System.Drawing.Size(125, 21)
        Me.TxtInputQuizzes.TabIndex = 2
        Me.TxtInputQuizzes.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblInputProject
        '
        Me.LblInputProject.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblInputProject.AutoSize = True
        Me.LblInputProject.Location = New System.Drawing.Point(6, 74)
        Me.LblInputProject.Name = "LblInputProject"
        Me.LblInputProject.Size = New System.Drawing.Size(50, 19)
        Me.LblInputProject.TabIndex = 0
        Me.LblInputProject.Text = "Project"
        '
        'TxtInputProject
        '
        Me.TxtInputProject.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtInputProject.Enabled = False
        Me.TxtInputProject.Location = New System.Drawing.Point(62, 73)
        Me.TxtInputProject.Name = "TxtInputProject"
        Me.TxtInputProject.ReadOnly = True
        Me.TxtInputProject.Size = New System.Drawing.Size(125, 21)
        Me.TxtInputProject.TabIndex = 3
        Me.TxtInputProject.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblInputExam
        '
        Me.LblInputExam.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.LblInputExam.AutoSize = True
        Me.LblInputExam.Location = New System.Drawing.Point(193, 74)
        Me.LblInputExam.Name = "LblInputExam"
        Me.LblInputExam.Size = New System.Drawing.Size(41, 19)
        Me.LblInputExam.TabIndex = 0
        Me.LblInputExam.Text = "Exam"
        '
        'TxtInputExam
        '
        Me.TxtInputExam.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.TxtInputExam.Enabled = False
        Me.TxtInputExam.Location = New System.Drawing.Point(253, 73)
        Me.TxtInputExam.Name = "TxtInputExam"
        Me.TxtInputExam.ReadOnly = True
        Me.TxtInputExam.Size = New System.Drawing.Size(125, 21)
        Me.TxtInputExam.TabIndex = 4
        Me.TxtInputExam.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'GrpResult
        '
        Me.GrpResult.Controls.Add(Me.LblResultCA)
        Me.GrpResult.Controls.Add(Me.TxtResultCA)
        Me.GrpResult.Controls.Add(Me.LblResultModule)
        Me.GrpResult.Controls.Add(Me.TxtResultModule)
        Me.GrpResult.Controls.Add(Me.LblResultGrade)
        Me.GrpResult.Controls.Add(Me.TxtReusltGrade)
        Me.GrpResult.Controls.Add(Me.LblReusltRemarks)
        Me.GrpResult.Controls.Add(Me.TxtReusltRemarks)
        Me.GrpResult.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.GrpResult.Location = New System.Drawing.Point(23, 196)
        Me.GrpResult.Name = "GrpResult"
        Me.GrpResult.Size = New System.Drawing.Size(500, 69)
        Me.GrpResult.TabIndex = 3
        Me.GrpResult.TabStop = False
        Me.GrpResult.Text = "Module Result"
        '
        'LblResultCA
        '
        Me.LblResultCA.AutoSize = True
        Me.LblResultCA.Location = New System.Drawing.Point(6, 16)
        Me.LblResultCA.Name = "LblResultCA"
        Me.LblResultCA.Size = New System.Drawing.Size(66, 19)
        Me.LblResultCA.TabIndex = 0
        Me.LblResultCA.Text = "CA Marks"
        '
        'TxtResultCA
        '
        Me.TxtResultCA.Location = New System.Drawing.Point(106, 15)
        Me.TxtResultCA.Name = "TxtResultCA"
        Me.TxtResultCA.ReadOnly = True
        Me.TxtResultCA.Size = New System.Drawing.Size(125, 21)
        Me.TxtResultCA.TabIndex = 1
        '
        'LblResultModule
        '
        Me.LblResultModule.AutoSize = True
        Me.LblResultModule.Location = New System.Drawing.Point(237, 16)
        Me.LblResultModule.Name = "LblResultModule"
        Me.LblResultModule.Size = New System.Drawing.Size(93, 19)
        Me.LblResultModule.TabIndex = 0
        Me.LblResultModule.Text = "Module Marks"
        '
        'TxtResultModule
        '
        Me.TxtResultModule.Location = New System.Drawing.Point(336, 15)
        Me.TxtResultModule.Name = "TxtResultModule"
        Me.TxtResultModule.ReadOnly = True
        Me.TxtResultModule.Size = New System.Drawing.Size(125, 21)
        Me.TxtResultModule.TabIndex = 2
        '
        'LblResultGrade
        '
        Me.LblResultGrade.AutoSize = True
        Me.LblResultGrade.Location = New System.Drawing.Point(6, 43)
        Me.LblResultGrade.Name = "LblResultGrade"
        Me.LblResultGrade.Size = New System.Drawing.Size(94, 19)
        Me.LblResultGrade.TabIndex = 0
        Me.LblResultGrade.Text = "Module Grade"
        '
        'TxtReusltGrade
        '
        Me.TxtReusltGrade.Location = New System.Drawing.Point(106, 42)
        Me.TxtReusltGrade.Name = "TxtReusltGrade"
        Me.TxtReusltGrade.ReadOnly = True
        Me.TxtReusltGrade.Size = New System.Drawing.Size(125, 21)
        Me.TxtReusltGrade.TabIndex = 3
        '
        'LblReusltRemarks
        '
        Me.LblReusltRemarks.AutoSize = True
        Me.LblReusltRemarks.Location = New System.Drawing.Point(237, 43)
        Me.LblReusltRemarks.Name = "LblReusltRemarks"
        Me.LblReusltRemarks.Size = New System.Drawing.Size(59, 19)
        Me.LblReusltRemarks.TabIndex = 0
        Me.LblReusltRemarks.Text = "Remarks"
        '
        'TxtReusltRemarks
        '
        Me.TxtReusltRemarks.Location = New System.Drawing.Point(336, 42)
        Me.TxtReusltRemarks.Name = "TxtReusltRemarks"
        Me.TxtReusltRemarks.ReadOnly = True
        Me.TxtReusltRemarks.Size = New System.Drawing.Size(125, 21)
        Me.TxtReusltRemarks.TabIndex = 4
        '
        'GrpStatistics
        '
        Me.GrpStatistics.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GrpStatistics.Controls.Add(Me.LblStatisticsMd)
        Me.GrpStatistics.Controls.Add(Me.TxtStatisticsMd)
        Me.GrpStatistics.Controls.Add(Me.LblStatisticsNo)
        Me.GrpStatistics.Controls.Add(Me.TxtStatisticsNo)
        Me.GrpStatistics.Controls.Add(Me.LblStatisticsA)
        Me.GrpStatistics.Controls.Add(Me.TxtStatisticsA)
        Me.GrpStatistics.Controls.Add(Me.LblStatisticsB)
        Me.GrpStatistics.Controls.Add(Me.TxtStatisticsB)
        Me.GrpStatistics.Controls.Add(Me.LblStatisticsC)
        Me.GrpStatistics.Controls.Add(Me.TxtStatisticsC)
        Me.GrpStatistics.Controls.Add(Me.LblStatisticsF)
        Me.GrpStatistics.Controls.Add(Me.TxtStatisticsF)
        Me.GrpStatistics.Controls.Add(Me.LblStatisticsAv)
        Me.GrpStatistics.Controls.Add(Me.TxtStatisticsAv)
        Me.GrpStatistics.Controls.Add(Me.LblStatisticsSd)
        Me.GrpStatistics.Controls.Add(Me.TxtStatisticsSd)
        Me.GrpStatistics.Font = New System.Drawing.Font("微軟正黑體", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(136, Byte))
        Me.GrpStatistics.Location = New System.Drawing.Point(23, 271)
        Me.GrpStatistics.Name = "GrpStatistics"
        Me.GrpStatistics.Size = New System.Drawing.Size(500, 157)
        Me.GrpStatistics.TabIndex = 4
        Me.GrpStatistics.TabStop = False
        Me.GrpStatistics.Text = "Statistics"
        '
        'LblStatisticsMd
        '
        Me.LblStatisticsMd.AutoSize = True
        Me.LblStatisticsMd.Location = New System.Drawing.Point(220, 77)
        Me.LblStatisticsMd.Name = "LblStatisticsMd"
        Me.LblStatisticsMd.Size = New System.Drawing.Size(53, 19)
        Me.LblStatisticsMd.TabIndex = 0
        Me.LblStatisticsMd.Text = "Median"
        '
        'TxtStatisticsMd
        '
        Me.TxtStatisticsMd.Location = New System.Drawing.Point(346, 76)
        Me.TxtStatisticsMd.Name = "TxtStatisticsMd"
        Me.TxtStatisticsMd.ReadOnly = True
        Me.TxtStatisticsMd.Size = New System.Drawing.Size(100, 21)
        Me.TxtStatisticsMd.TabIndex = 8
        Me.TxtStatisticsMd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblStatisticsNo
        '
        Me.LblStatisticsNo.AutoSize = True
        Me.LblStatisticsNo.Location = New System.Drawing.Point(6, 23)
        Me.LblStatisticsNo.Name = "LblStatisticsNo"
        Me.LblStatisticsNo.Size = New System.Drawing.Size(127, 19)
        Me.LblStatisticsNo.TabIndex = 0
        Me.LblStatisticsNo.Text = "Number of Students"
        '
        'TxtStatisticsNo
        '
        Me.TxtStatisticsNo.Location = New System.Drawing.Point(139, 22)
        Me.TxtStatisticsNo.Name = "TxtStatisticsNo"
        Me.TxtStatisticsNo.ReadOnly = True
        Me.TxtStatisticsNo.Size = New System.Drawing.Size(75, 21)
        Me.TxtStatisticsNo.TabIndex = 1
        Me.TxtStatisticsNo.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblStatisticsA
        '
        Me.LblStatisticsA.AutoSize = True
        Me.LblStatisticsA.Location = New System.Drawing.Point(6, 50)
        Me.LblStatisticsA.Name = "LblStatisticsA"
        Me.LblStatisticsA.Size = New System.Drawing.Size(73, 19)
        Me.LblStatisticsA.TabIndex = 0
        Me.LblStatisticsA.Text = "Count of A"
        '
        'TxtStatisticsA
        '
        Me.TxtStatisticsA.Location = New System.Drawing.Point(139, 49)
        Me.TxtStatisticsA.Name = "TxtStatisticsA"
        Me.TxtStatisticsA.ReadOnly = True
        Me.TxtStatisticsA.Size = New System.Drawing.Size(75, 21)
        Me.TxtStatisticsA.TabIndex = 2
        Me.TxtStatisticsA.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblStatisticsB
        '
        Me.LblStatisticsB.AutoSize = True
        Me.LblStatisticsB.Location = New System.Drawing.Point(6, 77)
        Me.LblStatisticsB.Name = "LblStatisticsB"
        Me.LblStatisticsB.Size = New System.Drawing.Size(72, 19)
        Me.LblStatisticsB.TabIndex = 0
        Me.LblStatisticsB.Text = "Count of B"
        '
        'TxtStatisticsB
        '
        Me.TxtStatisticsB.Location = New System.Drawing.Point(139, 76)
        Me.TxtStatisticsB.Name = "TxtStatisticsB"
        Me.TxtStatisticsB.ReadOnly = True
        Me.TxtStatisticsB.Size = New System.Drawing.Size(75, 21)
        Me.TxtStatisticsB.TabIndex = 3
        Me.TxtStatisticsB.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblStatisticsC
        '
        Me.LblStatisticsC.AutoSize = True
        Me.LblStatisticsC.Location = New System.Drawing.Point(6, 104)
        Me.LblStatisticsC.Name = "LblStatisticsC"
        Me.LblStatisticsC.Size = New System.Drawing.Size(73, 19)
        Me.LblStatisticsC.TabIndex = 0
        Me.LblStatisticsC.Text = "Count of C"
        '
        'TxtStatisticsC
        '
        Me.TxtStatisticsC.Location = New System.Drawing.Point(139, 103)
        Me.TxtStatisticsC.Name = "TxtStatisticsC"
        Me.TxtStatisticsC.ReadOnly = True
        Me.TxtStatisticsC.Size = New System.Drawing.Size(75, 21)
        Me.TxtStatisticsC.TabIndex = 4
        Me.TxtStatisticsC.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblStatisticsF
        '
        Me.LblStatisticsF.AutoSize = True
        Me.LblStatisticsF.Location = New System.Drawing.Point(6, 131)
        Me.LblStatisticsF.Name = "LblStatisticsF"
        Me.LblStatisticsF.Size = New System.Drawing.Size(71, 19)
        Me.LblStatisticsF.TabIndex = 0
        Me.LblStatisticsF.Text = "Count of F"
        '
        'TxtStatisticsF
        '
        Me.TxtStatisticsF.Location = New System.Drawing.Point(139, 130)
        Me.TxtStatisticsF.Name = "TxtStatisticsF"
        Me.TxtStatisticsF.ReadOnly = True
        Me.TxtStatisticsF.Size = New System.Drawing.Size(75, 21)
        Me.TxtStatisticsF.TabIndex = 5
        Me.TxtStatisticsF.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblStatisticsAv
        '
        Me.LblStatisticsAv.AutoSize = True
        Me.LblStatisticsAv.Location = New System.Drawing.Point(220, 23)
        Me.LblStatisticsAv.Name = "LblStatisticsAv"
        Me.LblStatisticsAv.Size = New System.Drawing.Size(58, 19)
        Me.LblStatisticsAv.TabIndex = 0
        Me.LblStatisticsAv.Text = "Average"
        '
        'TxtStatisticsAv
        '
        Me.TxtStatisticsAv.Location = New System.Drawing.Point(346, 22)
        Me.TxtStatisticsAv.Name = "TxtStatisticsAv"
        Me.TxtStatisticsAv.ReadOnly = True
        Me.TxtStatisticsAv.Size = New System.Drawing.Size(100, 21)
        Me.TxtStatisticsAv.TabIndex = 6
        Me.TxtStatisticsAv.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'LblStatisticsSd
        '
        Me.LblStatisticsSd.AutoSize = True
        Me.LblStatisticsSd.Location = New System.Drawing.Point(220, 50)
        Me.LblStatisticsSd.Name = "LblStatisticsSd"
        Me.LblStatisticsSd.Size = New System.Drawing.Size(120, 19)
        Me.LblStatisticsSd.TabIndex = 0
        Me.LblStatisticsSd.Text = "Standard Deviation"
        '
        'TxtStatisticsSd
        '
        Me.TxtStatisticsSd.Location = New System.Drawing.Point(346, 49)
        Me.TxtStatisticsSd.Name = "TxtStatisticsSd"
        Me.TxtStatisticsSd.ReadOnly = True
        Me.TxtStatisticsSd.Size = New System.Drawing.Size(100, 21)
        Me.TxtStatisticsSd.TabIndex = 7
        Me.TxtStatisticsSd.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'GrpRecords
        '
        Me.GrpRecords.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GrpRecords.Controls.Add(Me.ChkRecords)
        Me.GrpRecords.Controls.Add(Me.BtnRecordsAdd)
        Me.GrpRecords.Controls.Add(Me.BtnRecordsRemove)
        Me.GrpRecords.Controls.Add(Me.LstRecords)
        Me.GrpRecords.Controls.Add(Me.TxtRecordsSearch)
        Me.GrpRecords.Controls.Add(Me.ChkRecordsSearch)
        Me.GrpRecords.Font = New System.Drawing.Font("微軟正黑體", 9.0!)
        Me.GrpRecords.Location = New System.Drawing.Point(530, 63)
        Me.GrpRecords.Name = "GrpRecords"
        Me.GrpRecords.Size = New System.Drawing.Size(225, 365)
        Me.GrpRecords.TabIndex = 5
        Me.GrpRecords.TabStop = False
        Me.GrpRecords.Text = "Records"
        '
        'ChkRecords
        '
        Me.ChkRecords.AutoSize = True
        Me.ChkRecords.Location = New System.Drawing.Point(6, 49)
        Me.ChkRecords.Name = "ChkRecords"
        Me.ChkRecords.Size = New System.Drawing.Size(140, 15)
        Me.ChkRecords.TabIndex = 3
        Me.ChkRecords.Text = "Allow Repeated Name"
        Me.ChkRecords.UseVisualStyleBackColor = True
        '
        'BtnRecordsAdd
        '
        Me.BtnRecordsAdd.Enabled = False
        Me.BtnRecordsAdd.Location = New System.Drawing.Point(6, 22)
        Me.BtnRecordsAdd.Name = "BtnRecordsAdd"
        Me.BtnRecordsAdd.Size = New System.Drawing.Size(75, 21)
        Me.BtnRecordsAdd.TabIndex = 1
        Me.BtnRecordsAdd.Text = "Add"
        '
        'BtnRecordsRemove
        '
        Me.BtnRecordsRemove.Enabled = False
        Me.BtnRecordsRemove.Location = New System.Drawing.Point(87, 22)
        Me.BtnRecordsRemove.Name = "BtnRecordsRemove"
        Me.BtnRecordsRemove.Size = New System.Drawing.Size(75, 21)
        Me.BtnRecordsRemove.TabIndex = 2
        Me.BtnRecordsRemove.Text = "Remove"
        '
        'LstRecords
        '
        Me.LstRecords.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.LstRecords.FormattingEnabled = True
        Me.LstRecords.ItemHeight = 16
        Me.LstRecords.Location = New System.Drawing.Point(6, 70)
        Me.LstRecords.Name = "LstRecords"
        Me.LstRecords.Size = New System.Drawing.Size(213, 244)
        Me.LstRecords.TabIndex = 4
        '
        'TxtRecordsSearch
        '
        Me.TxtRecordsSearch.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TxtRecordsSearch.Location = New System.Drawing.Point(6, 317)
        Me.TxtRecordsSearch.Name = "TxtRecordsSearch"
        Me.TxtRecordsSearch.Size = New System.Drawing.Size(213, 21)
        Me.TxtRecordsSearch.TabIndex = 5
        '
        'ChkRecordsSearch
        '
        Me.ChkRecordsSearch.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ChkRecordsSearch.AutoSize = True
        Me.ChkRecordsSearch.Location = New System.Drawing.Point(6, 344)
        Me.ChkRecordsSearch.Name = "ChkRecordsSearch"
        Me.ChkRecordsSearch.Size = New System.Drawing.Size(155, 15)
        Me.ChkRecordsSearch.TabIndex = 6
        Me.ChkRecordsSearch.Text = "Using Regular Expression"
        Me.ChkRecordsSearch.UseVisualStyleBackColor = True
        '
        'TmrMain
        '
        '
        'FrmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(778, 449)
        Me.Controls.Add(Me.PrbMain)
        Me.Controls.Add(Me.LblName)
        Me.Controls.Add(Me.TxtName)
        Me.Controls.Add(Me.GrpInput)
        Me.Controls.Add(Me.GrpResult)
        Me.Controls.Add(Me.GrpStatistics)
        Me.Controls.Add(Me.GrpRecords)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "FrmMain"
        Me.Padding = New System.Windows.Forms.Padding(20, 60, 20, 18)
        Me.Style = MetroFramework.MetroColorStyle.Green
        Me.Text = "Calculation of Module Grade"
        Me.Theme = MetroFramework.MetroThemeStyle.Light
        Me.GrpInput.ResumeLayout(False)
        Me.GrpInput.PerformLayout()
        Me.GrpResult.ResumeLayout(False)
        Me.GrpResult.PerformLayout()
        Me.GrpStatistics.ResumeLayout(False)
        Me.GrpStatistics.PerformLayout()
        Me.GrpRecords.ResumeLayout(False)
        Me.GrpRecords.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Private WithEvents PrbMain As MetroFramework.Controls.MetroProgressBar
    Private WithEvents LblName As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtName As MetroFramework.Controls.MetroTextBox
    Private WithEvents GrpInput As GroupBox
    Private WithEvents LblInputMain As MetroFramework.Controls.MetroLabel
    Private WithEvents LblInputTest As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtInputTest As MetroFramework.Controls.MetroTextBox
    Private WithEvents LblInputQuizzes As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtInputQuizzes As MetroFramework.Controls.MetroTextBox
    Private WithEvents LblInputProject As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtInputProject As MetroFramework.Controls.MetroTextBox
    Private WithEvents LblInputExam As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtInputExam As MetroFramework.Controls.MetroTextBox
    Private WithEvents GrpResult As GroupBox
    Private WithEvents LblResultCA As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtResultCA As MetroFramework.Controls.MetroTextBox
    Private WithEvents LblResultModule As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtResultModule As MetroFramework.Controls.MetroTextBox
    Private WithEvents LblResultGrade As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtReusltGrade As MetroFramework.Controls.MetroTextBox
    Private WithEvents LblReusltRemarks As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtReusltRemarks As MetroFramework.Controls.MetroTextBox
    Private WithEvents GrpStatistics As GroupBox
    Private WithEvents LblStatisticsNo As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtStatisticsNo As MetroFramework.Controls.MetroTextBox
    Private WithEvents LblStatisticsA As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtStatisticsA As MetroFramework.Controls.MetroTextBox
    Private WithEvents LblStatisticsB As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtStatisticsB As MetroFramework.Controls.MetroTextBox
    Private WithEvents LblStatisticsC As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtStatisticsC As MetroFramework.Controls.MetroTextBox
    Private WithEvents LblStatisticsF As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtStatisticsF As MetroFramework.Controls.MetroTextBox
    Private WithEvents LblStatisticsAv As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtStatisticsAv As MetroFramework.Controls.MetroTextBox
    Private WithEvents LblStatisticsSd As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtStatisticsSd As MetroFramework.Controls.MetroTextBox
    Private WithEvents LblStatisticsMd As MetroFramework.Controls.MetroLabel
    Private WithEvents TxtStatisticsMd As MetroFramework.Controls.MetroTextBox
    Private WithEvents GrpRecords As GroupBox
    Private WithEvents BtnRecordsAdd As MetroFramework.Controls.MetroButton
    Private WithEvents BtnRecordsRemove As MetroFramework.Controls.MetroButton
    Private WithEvents LstRecords As ListBox
    Private WithEvents TxtRecordsSearch As MetroFramework.Controls.MetroTextBox
    Private WithEvents ChkRecordsSearch As MetroFramework.Controls.MetroCheckBox
    Friend WithEvents TmrMain As Timer
    Friend WithEvents ChkRecords As MetroFramework.Controls.MetroCheckBox
End Class
