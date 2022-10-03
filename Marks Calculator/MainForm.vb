' Name:         Marks Calculator Project
' Purpose:      Used to calculate student grades.
' Programmer:   Chow Cheuk Hei, Tse Ka Yu on 2 / 10 / 2022

Imports System.Text.RegularExpressions
Imports MetroFramework.Controls

Public Class FrmMain

#Region "Fields"

    ''' <summary>
    ''' 用來存儲目前數據的容器
    ''' </summary>
    Private Temp As Record

    ''' <summary>
    ''' 用來存儲多筆數據的容器
    ''' </summary>
    Private ReadOnly Data As List(Of Record)

#End Region

#Region "Constructors"

    Public Sub New()
        ' 設計工具需要此呼叫。
        InitializeComponent()
        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。
        MinimumSize = Size
        Temp = New Record(True)
        Data = New List(Of Record)()
    End Sub

#End Region

#Region "Properties"

    ''' <summary>
    ''' 表示數據欄位相應的 Record
    ''' </summary>
    ''' <returns></returns>
    Private ReadOnly Property InputedRecord As Record
        Get
            Return New Record(TxtName.Text, Double.Parse(TxtInputTest.Text), Double.Parse(TxtInputQuizzes.Text), Double.Parse(TxtInputProject.Text), Double.Parse(TxtInputExam.Text))
        End Get
    End Property

    ''' <summary>
    ''' 表示 LstRecords 已選取項目相應的 Record
    ''' </summary>
    ''' <returns></returns>
    Private ReadOnly Property SelectedRecord As Record
        Get
            If LstRecords.SelectedIndex = 0 Then
                Return Temp
            End If
            Return Data.Where(
                Function(record As Record) As Boolean
                    Return record.StudentName = LstRecords.SelectedItem.ToString()
                End Function
            ).Single()
        End Get
    End Property

#End Region

#Region "Enumerations"

    ''' <summary>
    ''' 代表控制項正在輸入數據
    ''' </summary>
    Private Enum IsTyping
        Yes = 1
        No = 2
    End Enum

    ''' <summary>
    ''' 代表控制項正在插入數據
    ''' </summary>
    Private Enum IsAdding
        Yes = 1
        No = 2
    End Enum

#End Region

#Region "Methods"

    Private Sub GetInputs()
        TxtName.Text = SelectedRecord.StudentName
        TxtInputTest.Text = SelectedRecord.TestMarks.ToString()
        TxtInputQuizzes.Text = SelectedRecord.QuizzesMarks.ToString()
        TxtInputProject.Text = SelectedRecord.ProjectMarks.ToString()
        TxtInputExam.Text = SelectedRecord.ExamMarks.ToString()
    End Sub

    Private Sub ShowResult(Result As Record)
        If Result.IsReal Then
            TxtResultCA.Text = Result.CAMarks.ToString()
            TxtResultModule.Text = Result.ModuleMarks.ToString()
            TxtReusltGrade.Text = Result.ModuleGrade.ToString()
            TxtReusltRemarks.Text = Result.Remarks.ToString()
        Else
            TxtResultCA.Text = "[Error Input]"
            TxtResultModule.Text = "[Error Input]"
            TxtReusltGrade.Text = "[Error Input]"
            TxtReusltRemarks.Text = "[Error Input]"
        End If
    End Sub

    Private Sub ShowStatistics()
        TxtStatisticsNo.Text = Data.Count.ToString()
        TxtStatisticsA.Text = Data.LongCount(
            Function(Record As Record) As Boolean
                Return Record.ModuleGrade = "A"
            End Function
        ).ToString()
        TxtStatisticsB.Text = Data.LongCount(
            Function(Record As Record) As Boolean
                Return Record.ModuleGrade = "B"
            End Function
        ).ToString()
        TxtStatisticsC.Text = Data.LongCount(
            Function(Record As Record) As Boolean
                Return Record.ModuleGrade = "C"
            End Function
        ).ToString()
        TxtStatisticsF.Text = Data.LongCount(
            Function(Record As Record) As Boolean
                Return Record.ModuleGrade = "F"
            End Function
        ).ToString()
        Dim N As Double = Data.LongCount()
        If N = 0 Then
            TxtStatisticsAv.Text = "[NaN]"
            TxtStatisticsSd.Text = "[NaN]"
            Return
        End If
        Dim Av As Double = Data.Sum(
            Function(Record As Record) As Double
                Return Record.ModuleMarks
            End Function
        ) / N
        TxtStatisticsAv.Text = Av.ToString()
        Dim Sd As Double = Math.Sqrt(
            Data.Sum(
                Function(Record As Record) As Double
                    Return (Record.ModuleMarks - Av) ^ 2
                End Function
            ) / N
        )
        TxtStatisticsSd.Text = Sd.ToString()
    End Sub

    Private Sub RecordsSearch()
        LstRecords.Items.Clear()
        LstRecords.Items.Add("(Input)")
        For Each Record As Record In Data
            Dim IsMatched As Boolean = False
            Try
                If ChkRecordsSearch.Checked Then
                    IsMatched = Regex.IsMatch(Record.StudentName, TxtRecordsSearch.Text)
                Else
                    IsMatched = Record.StudentName.IndexOf(TxtRecordsSearch.Text) <> -1
                End If
            Catch ex As Exception
            End Try
            If IsMatched Then
                LstRecords.Items.Add(Record.StudentName)
            End If
        Next
        LstRecords.SelectedIndex = 0
    End Sub

#End Region

#Region "Handles"

    Private Sub MainForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LblInputMain.Text = "CA Components: Test - " + (Record.TestScale * 100).ToString()
        LblInputMain.Text += "%, Quiz - " + (Record.QuizzesScale * 100).ToString()
        LblInputMain.Text += "%, Project - " + (Record.ProjectScale * 100).ToString() + "%"
        GrpResult.Text += " [CA - " + (Record.CAScale * 100).ToString()
        GrpResult.Text += "%, Exam - " + (Record.ExamScale * 100).ToString() + "%]"
        LstRecords.Items.Add("(Input)")
        LstRecords.SelectedIndex = 0
        GetInputs()
    End Sub

    Private Sub TxtInput_Enter(sender As Object, e As EventArgs) Handles TxtInputTest.Enter, TxtInputQuizzes.Enter, TxtInputProject.Enter, TxtInputExam.Enter
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New NotImplementedException()
        End If
        If LstRecords.Tag = IsAdding.Yes Then
            TryCast(sender, MetroTextBox).Tag = IsTyping.Yes
        End If
    End Sub

    Private Sub TxtInput_TextChanged(sender As Object, e As EventArgs) Handles TxtInputTest.TextChanged, TxtInputQuizzes.TextChanged, TxtInputProject.TextChanged, TxtInputExam.TextChanged
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New NotImplementedException()
        End If
        If LstRecords.Tag = IsAdding.Yes AndAlso TryCast(sender, MetroTextBox).Tag = IsTyping.Yes Then
            Dim Number As Double = 0
            ShowResult(If(Double.TryParse(TryCast(sender, MetroTextBox).Text, Number), InputedRecord, New Record(False)))
        End If
    End Sub

    Private Sub TxtInput_Leave(sender As Object, e As EventArgs) Handles TxtInputTest.Leave, TxtInputQuizzes.Leave, TxtInputProject.Leave, TxtInputExam.Leave
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New NotImplementedException()
        End If
        If LstRecords.Tag = IsAdding.Yes Then
            Dim Number As Double = 0
            Double.TryParse(TryCast(sender, MetroTextBox).Text, Number)
            If Number > 100 Then
                Number = 100
            ElseIf Number < 0 Then
                Number = 0
            End If
            TryCast(sender, MetroTextBox).Text = Number.ToString()
            TryCast(sender, MetroTextBox).Tag = IsTyping.No
            Temp = InputedRecord
            ShowResult(InputedRecord)
        End If
    End Sub

    Private Sub BtnRecordsAdd_Click(sender As Object, e As EventArgs) Handles BtnRecordsAdd.Click
        If TxtName.Text = String.Empty Then
            MessageBox.Show(Me, "Student name cannot be empty!", Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        For Each record As Record In Data
            If record.StudentName = InputedRecord.StudentName Then
                MessageBox.Show(Me, "Student name is already exist!", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
        Next
        Data.Add(InputedRecord)
        Temp.Clear()
        ShowStatistics()
        RecordsSearch()
    End Sub

    Private Sub BtnRecordsRemove_Click(sender As Object, e As EventArgs) Handles BtnRecordsRemove.Click
        Dim Index As Integer = LstRecords.SelectedIndex
        Data.Remove(SelectedRecord)
        ShowStatistics()
        RecordsSearch()
        LstRecords.SelectedIndex = If(Index < LstRecords.Items.Count, Index, 0)
    End Sub

    Private Sub LstRecord_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LstRecords.SelectedIndexChanged
        If LstRecords.SelectedIndex = 0 Then
            BtnRecordsAdd.Enabled = True
            BtnRecordsRemove.Enabled = False
            TxtName.ReadOnly = False
            TxtInputTest.ReadOnly = False
            TxtInputQuizzes.ReadOnly = False
            TxtInputProject.ReadOnly = False
            TxtInputExam.ReadOnly = False
            LstRecords.Tag = IsAdding.Yes
        Else
            BtnRecordsAdd.Enabled = False
            BtnRecordsRemove.Enabled = True
            TxtName.ReadOnly = True
            TxtInputTest.ReadOnly = True
            TxtInputQuizzes.ReadOnly = True
            TxtInputProject.ReadOnly = True
            TxtInputExam.ReadOnly = True
            LstRecords.Tag = IsAdding.No
        End If
        GetInputs()
        ShowResult(InputedRecord)
    End Sub

    Private Sub AnyRecordsSearch_Event(sender As Object, e As EventArgs) Handles TxtRecordsSearch.TextChanged, ChkRecordsSearch.CheckedChanged
        Dim Index As Integer = LstRecords.SelectedIndex
        RecordsSearch()
        LstRecords.SelectedIndex = If(Index < LstRecords.Items.Count, Index, 0)
    End Sub

    Private Sub FrmMain_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        LstRecords.Height = TxtRecordsSearch.Location.Y - LstRecords.Location.Y - 6
    End Sub

#End Region

    ''' <summary>
    ''' 一個類別用來管理一筆數據
    ''' </summary>
    Private Class Record

#Region "Constants"

        Friend Const TestScale As Double = 0.5
        Friend Const QuizzesScale As Double = 0.2
        Friend Const ProjectScale As Double = 0.3
        Friend Const CAScale As Double = 0.4
        Friend Const ExamScale As Double = 0.6
        Private Const None As String = "[None]"

#End Region

#Region "Fields"

        Private Name As String
        Private Test As Double
        Private Quizzes As Double
        Private Project As Double
        Private Exam As Double

#End Region

#Region "Constructors"

        Public Sub New(Valid As Boolean)
            Name = None
            If Valid = True Then
                Test = 0
                Quizzes = 0
                Project = 0
                Exam = 0
            Else
                Test = Double.MaxValue
                Quizzes = Double.MaxValue
                Project = Double.MaxValue
                Exam = Double.MaxValue
            End If
        End Sub
        Public Sub New(Name As String)
            Me.Name = Name
            Test = 0
            Quizzes = 0
            Project = 0
            Exam = 0
        End Sub
        Public Sub New(Name As String, Test As Double, Quizzes As Double, Project As Double, Exam As Double)
            Me.Name = Name
            Me.Test = Test
            Me.Quizzes = Quizzes
            Me.Project = Project
            Me.Exam = Exam
        End Sub

#End Region

#Region "Properties"

        Public ReadOnly Property CAMarks As Double
            Get
                Return Test * TestScale + Quizzes * QuizzesScale + Project * ProjectScale
            End Get
        End Property

        Public ReadOnly Property ModuleMarks As Double
            Get
                Return CAMarks * CAScale + Exam * ExamScale
            End Get
        End Property

        Public ReadOnly Property IsReal As Boolean
            Get
                Return Test >= 0 AndAlso Test <= 100 AndAlso
                    Project >= 0 AndAlso Project <= 100 AndAlso
                    Quizzes >= 0 AndAlso Quizzes <= 100 AndAlso
                    Exam >= 0 AndAlso Exam <= 100
            End Get
        End Property

        Public ReadOnly Property ModuleGrade As String
            Get
                If CAMarks < 40 OrElse ExamMarks < 40 Then
                    Return "F"
                ElseIf ModuleMarks >= 75 AndAlso ModuleMarks <= 100 Then
                    Return "A"
                ElseIf ModuleMarks >= 65 AndAlso ModuleMarks < 75 Then
                    Return "B"
                ElseIf ModuleMarks >= 40 AndAlso ModuleMarks < 65 Then
                    Return "C"
                Else
                    Return "---"
                End If
            End Get
        End Property

        Public ReadOnly Property Remarks As String
            Get
                If ModuleGrade = "F" Then
                    If ModuleMarks >= 30 Then
                        Return "Resit Exam"
                    Else
                        Return "Restudy"
                    End If
                ElseIf ModuleGrade = "---" Then
                    Return "---"
                Else
                    Return "Pass"
                End If
            End Get
        End Property

        Public Property StudentName As String
            Get
                Return If(Name <> None, Name, String.Empty)
            End Get
            Set(value As String)
                Name = value
            End Set
        End Property

        Public Property TestMarks As Double
            Get
                Return Test
            End Get
            Set(value As Double)
                Test = value
            End Set
        End Property

        Public Property QuizzesMarks As Double
            Get
                Return Quizzes
            End Get
            Set(value As Double)
                Quizzes = value
            End Set
        End Property

        Public Property ProjectMarks As Double
            Get
                Return Project
            End Get
            Set(value As Double)
                Project = value
            End Set
        End Property

        Public Property ExamMarks As Double
            Get
                Return Exam
            End Get
            Set(value As Double)
                Exam = value
            End Set
        End Property

#End Region

#Region "Methods"

        Public Sub Clear()
            Name = None
            Test = 0
            Quizzes = 0
            Project = 0
            Exam = 0
        End Sub

#End Region

    End Class

End Class
