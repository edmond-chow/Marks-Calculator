' Name:         Marks Calculator Project
' Purpose:      Used to calculate student grades.
' Programmer:   Chow Cheuk Hei, Tse Ka Yu on 2 / 10 / 2022

Imports System.Text.RegularExpressions
Imports MetroFramework.Controls

Public Class FrmMain

#Region "Fields"

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
        Data = New List(Of Record)()
    End Sub

#End Region

#Region "Properties"

    ''' <summary>
    ''' 表示數據欄位相應的 Record（可以是輸入中的數據）
    ''' </summary>
    ''' <returns></returns>
    Private ReadOnly Property TempRecord As Record
        Get
            Return New Record(TxtName.Text, Double.Parse(TxtInputTest.Text), Double.Parse(TxtInputQuizzes.Text), Double.Parse(TxtInputProject.Text), Double.Parse(TxtInputExam.Text))
        End Get
    End Property

    ''' <summary>
    ''' 表示 LstRecords 已選取項目在 Data 中相應的 Record（可以是輸入中的數據）
    ''' </summary>
    ''' <returns></returns>
    Private ReadOnly Property SelectedRecord As Record
        Get
            If LstRecords.SelectedIndex = 0 Then
                Return InputRecord
            End If
            Return Data.Where(
                Function(record As Record) As Boolean
                    Return record.StudentName = LstRecords.SelectedItem.ToString()
                End Function
            ).Single()
        End Get
    End Property

    ''' <summary>
    ''' 表示輸入中的數據相應的 Record（即 Data 中首個元素）
    ''' </summary>
    ''' <returns></returns>
    Private Property InputRecord As Record
        Get
            Return Data.Item(0)
        End Get
        Set(value As Record)
            Data.Item(0) = value
        End Set
    End Property

#End Region

#Region "Enumerations"

    ''' <summary>
    ''' 一個枚舉代表控制項正在輸入
    ''' </summary>
    Private Enum IsTyping
        Yes = 1
        No = 2
    End Enum

#End Region

#Region "Methods"

    Private Sub GetResult()
        TxtName.Text = SelectedRecord.StudentName
        TxtInputTest.Text = SelectedRecord.TestMarks.ToString()
        TxtInputQuizzes.Text = SelectedRecord.QuizzesMarks.ToString()
        TxtInputProject.Text = SelectedRecord.ProjectMarks.ToString()
        TxtInputExam.Text = SelectedRecord.ExamMarks.ToString()
    End Sub

    Private Sub ShowResult()
        If InputRecord.IsReal Then
            TxtResultCA.Text = InputRecord.CAMarks.ToString()
            TxtResultModule.Text = InputRecord.ModuleMarks.ToString()
            TxtReusltGrade.Text = InputRecord.ModuleGrade.ToString()
            TxtReusltRemarks.Text = InputRecord.Remarks.ToString()
        Else
            TxtResultCA.Text = "[Error Input]"
            TxtResultModule.Text = "[Error Input]"
            TxtReusltGrade.Text = "[Error Input]"
            TxtReusltRemarks.Text = "[Error Input]"
        End If
    End Sub

    Private Sub ShowStatistics()
        TxtStatisticsNo.Text = (Data.Count - 1).ToString()
        TxtStatisticsA.Text = Data.LongCount(
            Function(record As Record) As Boolean
                Return record.ModuleGrade = "A"
            End Function
        ).ToString()
        TxtStatisticsB.Text = Data.LongCount(
            Function(record As Record) As Boolean
                Return record.ModuleGrade = "B"
            End Function
        ).ToString()
        TxtStatisticsC.Text = Data.LongCount(
            Function(record As Record) As Boolean
                Return record.ModuleGrade = "C"
            End Function
        ).ToString()
        TxtStatisticsD.Text = Data.LongCount(
            Function(record As Record) As Boolean
                Return record.ModuleGrade = "D"
            End Function
        ).ToString()
        Dim N As Double = Data.Skip(1).LongCount()
        If N = 0 Then
            TxtStatisticsAv.Text = "[NaN]"
            TxtStatisticsSd.Text = "[NaN]"
            Return
        End If
        Dim Av As Double = Data.Skip(1).Sum(
            Function(record As Record) As Double
                Return record.ModuleMarks
            End Function
        ) / N
        TxtStatisticsAv.Text = Av.ToString()
        Dim Sd As Double = Math.Sqrt(
            Data.Skip(1).Sum(
                Function(record As Record) As Double
                    Return (record.ModuleMarks - Av) ^ 2
                End Function
            ) / N
        )
        TxtStatisticsSd.Text = Sd.ToString()
    End Sub

    Private Sub RecordsSearch()
        LstRecords.Items.Clear()
        LstRecords.Items.Add("(Input)")
        For Each record As Record In Data.Skip(1)
            Dim IsMatched As Boolean = False
            Try
                If (ChkRecordsSearch.Checked = False AndAlso record.StudentName.IndexOf(TxtRecordsSearch.Text) <> -1) _
                OrElse Regex.IsMatch(record.StudentName, TxtRecordsSearch.Text) Then
                    IsMatched = True
                End If
            Catch ex As Exception
            End Try
            If IsMatched Then
                LstRecords.Items.Add(record.StudentName)
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
        Data.Add(New Record(True))
        LstRecords.Items.Add("(Input)")
        LstRecords.SelectedIndex = 0
        GetResult()
    End Sub

    Private Sub TxtInput_Enter(sender As Object, e As EventArgs) Handles TxtInputTest.Enter, TxtInputQuizzes.Enter, TxtInputProject.Enter, TxtInputExam.Enter
        If TypeOf sender IsNot MetroTextBox Then
            Throw New NotImplementedException()
        End If
        Dim txt As MetroTextBox = TryCast(sender, MetroTextBox)
        txt.Tag = IsTyping.Yes
    End Sub

    Private Sub TxtInput_Leave(sender As Object, e As EventArgs) Handles TxtInputTest.Leave, TxtInputQuizzes.Leave, TxtInputProject.Leave, TxtInputExam.Leave
        If TypeOf sender IsNot MetroTextBox Then
            Throw New NotImplementedException()
        End If
        Dim txt As MetroTextBox = TryCast(sender, MetroTextBox)
        Dim num As Double = 0
        Double.TryParse(txt.Text, num)
        If num > 100 Then
            num = 100
        ElseIf num < 0 Then
            num = 0
        End If
        txt.Text = num.ToString()
        txt.Tag = IsTyping.No
        ShowResult()
    End Sub

    Private Sub TxtInput_TextChanged(sender As Object, e As EventArgs) Handles TxtInputTest.TextChanged, TxtInputQuizzes.TextChanged, TxtInputProject.TextChanged, TxtInputExam.TextChanged
        If TypeOf sender IsNot MetroTextBox Then
            Throw New NotImplementedException()
        End If
        Dim txt As MetroTextBox = TryCast(sender, MetroTextBox)
        If txt.Tag = IsTyping.Yes Then
            Dim output As Double = 0
            InputRecord = If(Double.TryParse(txt.Text, output), TempRecord, New Record(False))
            ShowResult()
        End If
    End Sub

    Private Sub BtnRecordsAdd_Click(sender As Object, e As EventArgs) Handles BtnRecordsAdd.Click
        If TxtName.Text = String.Empty Then
            MessageBox.Show(Me, "Student name cannot be empty!", Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        For Each record As Record In Data.Skip(1)
            If record.StudentName = TempRecord.StudentName Then
                MessageBox.Show(Me, "Student name is already exist!", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If
        Next
        Data.Add(TempRecord)
        InputRecord.Clear()
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
        Else
            BtnRecordsAdd.Enabled = False
            BtnRecordsRemove.Enabled = True
            TxtName.ReadOnly = True
            TxtInputTest.ReadOnly = True
            TxtInputQuizzes.ReadOnly = True
            TxtInputProject.ReadOnly = True
            TxtInputExam.ReadOnly = True
        End If
        GetResult()
        ShowResult()
    End Sub

    Private Sub AnyRecordsSearch_Event(sender As Object, e As EventArgs) Handles TxtRecordsSearch.TextChanged, ChkRecordsSearch.CheckedChanged
        Dim Index As Integer = LstRecords.SelectedIndex
        RecordsSearch()
        LstRecords.SelectedIndex = If(Index < LstRecords.Items.Count, Index, 0)
    End Sub

#End Region

    ''' <summary>
    ''' 一個類別用來管理一筆數據
    ''' </summary>
    Private Class Record

#Region "Constants"

        Friend Const TestScale As Double = 0.5
        Friend Const ProjectScale As Double = 0.3
        Friend Const QuizzesScale As Double = 0.2
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
                Return Test * TestScale + Project * ProjectScale + Quizzes * QuizzesScale
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
                ElseIf ModuleGrade = "[Error Input]" Then
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
