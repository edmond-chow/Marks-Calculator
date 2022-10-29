' Name:         Marks Calculator Project
' Purpose:      Used to calculate student grades.
' Programmer:   Chow Cheuk Hei, Tse Ka Yu on 2 / 10 / 2022

Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions
Imports MetroFramework.Controls
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class FrmMain

#Region "Constants"

    Private Const FileName As String = "Records.json"

#End Region

#Region "Fields"

    ''' <summary>
    ''' 用來存儲目前數據的執行個體
    ''' </summary>
    Private Temp As Record

    ''' <summary>
    ''' 用來存儲多筆數據的執行個體
    ''' </summary>
    Private Data As List(Of Record)

    ''' <summary>
    ''' 用來代表數據的文件流
    ''' </summary>
    Private DataFile As FileStream

    ''' <summary>
    ''' 表示表單已經完成加載狀態
    ''' </summary>
    Private LoadHasFinish As Boolean

    ''' <summary>
    ''' 表示表單已經進入關閉狀態
    ''' </summary>
    Private CloseHasStarted As Boolean

    Private LastWindowState As FormWindowState
    Private ReadOnly RandomNumberGenerator As Random

#End Region

#Region "Constructors"

    Public Sub New()
        ' 設計工具需要此呼叫。
        InitializeComponent()
        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。
        MinimumSize = Size
        Temp = New Record(True)
        Data = New List(Of Record)()
        DataFile = Nothing
        LoadHasFinish = False
        CloseHasStarted = False
        LastWindowState = WindowState
        RandomNumberGenerator = New Random()
    End Sub

#End Region

#Region "Properties"

    ''' <summary>
    ''' 表示數據欄位相應的 Record
    ''' </summary>
    ''' <returns></returns>
    Private ReadOnly Property InputedRecord As Record
        Get
            Dim MyRecord As New Record(TxtName.Text, Double.Parse(TxtInputTest.Text), Double.Parse(TxtInputQuizzes.Text), Double.Parse(TxtInputProject.Text), Double.Parse(TxtInputExam.Text))
            Dim RandomNumber As Integer = 0
            While True
                RandomNumber = RandomNumberGenerator.Next(100000000, 999999999)
                If Data.LongCount(
                    Function(Record As Record) As Boolean
                        Return Record.ID = RandomNumber
                    End Function
                ) = 0 Then
                    Exit While
                End If
            End While
            MyRecord.ID = RandomNumber
            Return MyRecord
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
                Function(Record As Record) As Boolean
                    Return Record.ID = CType(TxtRecordsSearch.Tag, List(Of Integer)).ElementAt(LstRecords.SelectedIndex - 1)
                End Function
            ).Single()
        End Get
    End Property

    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Const WS_MINIMIZEBOX As Integer = &H20000
            Dim Params As CreateParams = MyBase.CreateParams
            Params.Style = Params.Style Or WS_MINIMIZEBOX
            Return Params
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

    Private Sub SuspendControls()
        BtnRecordsAdd.Enabled = False
        BtnRecordsRemove.Enabled = False
        ChkRecords.Enabled = False
        TxtName.Enabled = False
        TxtInputTest.Enabled = False
        TxtInputQuizzes.Enabled = False
        TxtInputProject.Enabled = False
        TxtInputExam.Enabled = False
        LstRecords.Enabled = False
        TxtRecordsSearch.Enabled = False
        ChkRecordsSearch.Enabled = False
        TxtResultCA.Enabled = False
        TxtResultModule.Enabled = False
        TxtReusltGrade.Enabled = False
        TxtReusltRemarks.Enabled = False
        TxtStatisticsNo.Enabled = False
        TxtStatisticsAv.Enabled = False
        TxtStatisticsSd.Enabled = False
        TxtStatisticsMd.Enabled = False
        TxtStatisticsA.Enabled = False
        TxtStatisticsB.Enabled = False
        TxtStatisticsC.Enabled = False
        TxtStatisticsF.Enabled = False
    End Sub

    Private Sub ResumeControls()
        BtnRecordsAdd.Enabled = True
        BtnRecordsRemove.Enabled = True
        ChkRecords.Enabled = True
        TxtName.Enabled = True
        TxtInputTest.Enabled = True
        TxtInputQuizzes.Enabled = True
        TxtInputProject.Enabled = True
        TxtInputExam.Enabled = True
        LstRecords.Enabled = True
        TxtRecordsSearch.Enabled = True
        ChkRecordsSearch.Enabled = True
        TxtResultCA.Enabled = True
        TxtResultModule.Enabled = True
        TxtReusltGrade.Enabled = True
        TxtReusltRemarks.Enabled = True
        TxtStatisticsNo.Enabled = True
        TxtStatisticsAv.Enabled = True
        TxtStatisticsSd.Enabled = True
        TxtStatisticsMd.Enabled = True
        TxtStatisticsA.Enabled = True
        TxtStatisticsB.Enabled = True
        TxtStatisticsC.Enabled = True
        TxtStatisticsF.Enabled = True
    End Sub

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
            Const ErrorInput As String = "[Error Input]"
            TxtResultCA.Text = ErrorInput
            TxtResultModule.Text = ErrorInput
            TxtReusltGrade.Text = ErrorInput
            TxtReusltRemarks.Text = ErrorInput
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
            Const NaN As String = "[NaN]"
            TxtStatisticsAv.Text = NaN
            TxtStatisticsSd.Text = NaN
            TxtStatisticsMd.Text = NaN
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
        Dim Sorted As Double() = Data.Select(
            Function(Record As Record) As Double
                Return Record.ModuleMarks
            End Function
        ).ToArray()
        Array.Sort(Sorted)
        Dim Md As Double = If(Sorted.LongLength Mod 2 = 0,
            (Sorted(Sorted.LongLength / 2) + Sorted(Sorted.LongLength / 2 - 1)) / 2,
            Sorted((Sorted.LongLength - 1) / 2)
        )
        TxtStatisticsMd.Text = Md.ToString()
    End Sub

    Private Sub RecordsSearch()
        LstRecords.Items.Clear()
        LstRecords.Items.Add("(Input)")
        TxtRecordsSearch.Tag = New List(Of Integer)
        For Each Record As Record In Data
            Dim IsMatched As Boolean = False
            If ChkRecordsSearch.Checked Then
                Try
                    IsMatched = Regex.IsMatch(Record.StudentName, TxtRecordsSearch.Text)
                Catch Exception As Exception
                End Try
            Else
                IsMatched = Record.StudentName.IndexOf(TxtRecordsSearch.Text) <> -1
            End If
            If IsMatched Then
                LstRecords.Items.Add(Record.StudentName)
                CType(TxtRecordsSearch.Tag, List(Of Integer)).Add(Record.ID)
            End If
        Next
        LstRecords.SelectedIndex = 0
    End Sub

    Private Shared Function IsNotTheSameID(Enumerable As IEnumerable(Of IReliability)) As Boolean
        For i As Integer = 0 To Enumerable.Count() - 2
            For j As Integer = 1 To Enumerable.Count() - 1
                If Enumerable.ElementAt(i).ID = Enumerable.ElementAt(i + j).ID Then
                    Return False
                End If
            Next
        Next
        Return True
    End Function

    Private Shared Function IsNotTheSame(Enumerable As IEnumerable(Of IRecord)) As Boolean
        For i As Integer = 0 To Enumerable.Count() - 2
            For j As Integer = 1 To Enumerable.Count() - 1
                If Enumerable.ElementAt(i).StudentName = Enumerable.ElementAt(i + j).StudentName Then
                    Return False
                End If
            Next
        Next
        Return True
    End Function

    Private Async Function ReadDataFile() As Task(Of List(Of Record))
        SuspendControls()
        PrbMain.Show()
        Dim Records As New List(Of Record)
        Try
            If File.Exists(FileName) Then
                DataFile = File.Open(FileName, FileMode.Open)
                Dim Json(DataFile.Length - 1) As Byte
                Await DataFile.ReadAsync(Json, 0, DataFile.Length)
                For Each RecordToken As JToken In JsonConvert.DeserializeObject(Of JArray)(Encoding.UTF8.GetString(Json)).Children()
                    Dim TempRecord As New Record(True)
                    For Each Field As JProperty In RecordToken
                        Dim PropertyInfo As PropertyInfo = GetType(Record).GetProperty(Field.Name)
                        If PropertyInfo.PropertyType = GetType(String) Then
                            PropertyInfo.SetValue(TempRecord, Field.Value.ToString())
                        ElseIf PropertyInfo.PropertyType = GetType(Double) Then
                            PropertyInfo.SetValue(TempRecord, Double.Parse(Field.Value.ToString()))
                        ElseIf PropertyInfo.PropertyType = GetType(Integer) Then
                            PropertyInfo.SetValue(TempRecord, Integer.Parse(Field.Value.ToString()))
                        Else
                            Throw New NotImplementedException()
                        End If
                    Next
                    Records.Add(TempRecord)
                Next
            Else
                DataFile = File.Create(FileName)
            End If
        Catch Exception As Exception
            Records = New List(Of Record)()
        End Try
        If Not IsNotTheSameID(Records) Then
            Records = New List(Of Record)()
        End If
        ResumeControls()
        PrbMain.Hide()
        Return Records
    End Function

    Private Async Function WriteDataFile() As Task
        SuspendControls()
        PrbMain.Show()
        Try
            Dim Json As Byte() = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(Data, Formatting.Indented))
            DataFile.SetLength(0)
            Await DataFile.WriteAsync(Json, 0, Json.Length)
            DataFile.Close()
        Catch Exception As Exception
        End Try
        ResumeControls()
        PrbMain.Hide()
    End Function

#End Region

#Region "Handles"

    Private Async Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LblInputMain.Text = "CA Components: Test - " + (Record.TestScale * 100).ToString()
        LblInputMain.Text += "%, Quiz - " + (Record.QuizzesScale * 100).ToString()
        LblInputMain.Text += "%, Project - " + (Record.ProjectScale * 100).ToString() + "%"
        GrpResult.Text += " [CA - " + (Record.CAScale * 100).ToString()
        GrpResult.Text += "%, Exam - " + (Record.ExamScale * 100).ToString() + "%]"
        PrbMain.ProgressBarStyle = ProgressBarStyle.Marquee
        Data = Await ReadDataFile()
        ShowStatistics()
        RecordsSearch()
        If Not IsNotTheSame(Data) Then
            ChkRecords.Checked = True
        End If
        LoadHasFinish = True
    End Sub

    Private Async Sub FrmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If CloseHasStarted = False Then
            CloseHasStarted = True
            If LoadHasFinish = True Then
                Await WriteDataFile()
            End If
        End If
    End Sub

    Private Sub TxtInput_Enter(sender As Object, e As EventArgs) Handles TxtInputTest.Enter, TxtInputQuizzes.Enter, TxtInputProject.Enter, TxtInputExam.Enter
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New NotImplementedException()
        End If
        If LstRecords.Tag = IsAdding.Yes Then
            CType(sender, MetroTextBox).Tag = IsTyping.Yes
        End If
    End Sub

    Private Sub TxtInput_TextChanged(sender As Object, e As EventArgs) Handles TxtInputTest.TextChanged, TxtInputQuizzes.TextChanged, TxtInputProject.TextChanged, TxtInputExam.TextChanged
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New NotImplementedException()
        End If
        If LstRecords.Tag = IsAdding.Yes AndAlso CType(sender, MetroTextBox).Tag = IsTyping.Yes Then
            Dim Number As Double = 0
            ShowResult(If(Double.TryParse(CType(sender, MetroTextBox).Text, Number), InputedRecord, New Record(False)))
        End If
    End Sub

    Private Sub TxtInput_Leave(sender As Object, e As EventArgs) Handles TxtInputTest.Leave, TxtInputQuizzes.Leave, TxtInputProject.Leave, TxtInputExam.Leave
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New NotImplementedException()
        End If
        If LstRecords.Tag = IsAdding.Yes Then
            Dim Number As Double = 0
            Double.TryParse(CType(sender, MetroTextBox).Text, Number)
            If Number > 100 Then
                Number = 100
            ElseIf Number < 0 Then
                Number = 0
            End If
            CType(sender, MetroTextBox).Text = Number.ToString()
            CType(sender, MetroTextBox).Tag = IsTyping.No
            Temp = InputedRecord
            ShowResult(InputedRecord)
        End If
    End Sub

    Private Sub BtnRecordsAdd_Click(sender As Object, e As EventArgs) Handles BtnRecordsAdd.Click
        If TxtName.Text = String.Empty Then
            MessageBox.Show(Me, "Student name cannot be empty!", Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        If Not ChkRecords.Checked Then
            For Each Record As Record In Data
                If Record.StudentName = InputedRecord.StudentName Then
                    MessageBox.Show(Me, "Student name is already exist!", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Return
                End If
            Next
        End If
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

    Private Sub ChkRecords_CheckedChanged(sender As Object, e As EventArgs) Handles ChkRecords.CheckedChanged
        If Not ChkRecords.Checked AndAlso Not IsNotTheSame(Data) Then
            MessageBox.Show(Me, "Some of the student names have the same! But it still avoids the same input though.", Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub LstRecord_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LstRecords.SelectedIndexChanged
        If LstRecords.SelectedIndex = 0 Then
            BtnRecordsAdd.Enabled = True
            BtnRecordsRemove.Enabled = False
            ChkRecords.Enabled = True
            TxtName.ReadOnly = False
            TxtInputTest.ReadOnly = False
            TxtInputQuizzes.ReadOnly = False
            TxtInputProject.ReadOnly = False
            TxtInputExam.ReadOnly = False
            LstRecords.Tag = IsAdding.Yes
        Else
            BtnRecordsAdd.Enabled = False
            BtnRecordsRemove.Enabled = True
            ChkRecords.Enabled = False
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

    Private Sub ChkEnterKeys_Event(sender As Object, e As KeyEventArgs) Handles ChkRecords.KeyDown, ChkRecordsSearch.KeyDown
        If sender Is Nothing OrElse TypeOf sender IsNot MetroCheckBox Then
            Throw New NotImplementedException()
        End If
        If e.KeyCode = Keys.Enter Then
            CType(sender, MetroCheckBox).Checked = Not CType(sender, MetroCheckBox).Checked
        End If
    End Sub

    Private Sub FrmMain_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If LastWindowState <> WindowState AndAlso WindowState = FormWindowState.Normal Then
            TmrMain.Enabled = True
        End If
        LastWindowState = WindowState
    End Sub

    Private Sub TmrMain_Tick(sender As Object, e As EventArgs) Handles TmrMain.Tick
        If FocusMe() Then
            TmrMain.Enabled = False
        End If
    End Sub

#End Region

    Private Interface IReliability

#Region "Properties"

        Property ID As Integer

#End Region

    End Interface

    Private Interface IRecord

#Region "Properties"

        Property StudentName As String
        Property TestMarks As Double
        Property QuizzesMarks As Double
        Property ProjectMarks As Double
        Property ExamMarks As Double

#End Region

    End Interface

    Private Class Record
        Implements IRecord, IReliability

#Region "Constants"

        Friend Const TestScale As Double = 0.5
        Friend Const QuizzesScale As Double = 0.2
        Friend Const ProjectScale As Double = 0.3
        Friend Const CAScale As Double = 0.4
        Friend Const ExamScale As Double = 0.6
        Private Const None As String = "[None]"
        Private Const Invalid As String = "---"

#End Region

#Region "Fields"

        Private Name As String
        Private Test As Double
        Private Quizzes As Double
        Private Project As Double
        Private Exam As Double
        Private Code As Integer

#End Region

#Region "Constructors"

        Public Sub New(Valid As Boolean)
            Name = None
            If Valid = True Then
                Test = 0
                Quizzes = 0
                Project = 0
                Exam = 0
                Code = 0
            Else
                Test = Double.MaxValue
                Quizzes = Double.MaxValue
                Project = Double.MaxValue
                Exam = Double.MaxValue
                Code = 0
            End If
        End Sub

        Public Sub New(Name As String)
            Me.Name = Name
            Test = 0
            Quizzes = 0
            Project = 0
            Exam = 0
            Code = 0
        End Sub

        Public Sub New(Name As String, Test As Double, Quizzes As Double, Project As Double, Exam As Double)
            Me.Name = Name
            Me.Test = Test
            Me.Quizzes = Quizzes
            Me.Project = Project
            Me.Exam = Exam
            Code = 0
        End Sub

#End Region

#Region "Properties"

        <JsonIgnore>
        Public ReadOnly Property CAMarks As Double
            Get
                Return Test * TestScale + Quizzes * QuizzesScale + Project * ProjectScale
            End Get
        End Property

        <JsonIgnore>
        Public ReadOnly Property ModuleMarks As Double
            Get
                Return CAMarks * CAScale + Exam * ExamScale
            End Get
        End Property

        <JsonIgnore>
        Public ReadOnly Property IsReal As Boolean
            Get
                Return Test >= 0 AndAlso Test <= 100 AndAlso
                    Project >= 0 AndAlso Project <= 100 AndAlso
                    Quizzes >= 0 AndAlso Quizzes <= 100 AndAlso
                    Exam >= 0 AndAlso Exam <= 100
            End Get
        End Property

        <JsonIgnore>
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
                    Return Invalid
                End If
            End Get
        End Property

        <JsonIgnore>
        Public ReadOnly Property Remarks As String
            Get
                If ModuleGrade = "F" Then
                    If ModuleMarks >= 30 Then
                        Return "Resit Exam"
                    Else
                        Return "Restudy"
                    End If
                ElseIf ModuleGrade = Invalid Then
                    Return Invalid
                Else
                    Return "Pass"
                End If
            End Get
        End Property

        <JsonProperty>
        Public Property StudentName As String Implements IRecord.StudentName
            Get
                Return If(Name <> None, Name, String.Empty)
            End Get
            Set(Value As String)
                Name = Value
            End Set
        End Property

        <JsonProperty>
        Public Property TestMarks As Double Implements IRecord.TestMarks
            Get
                Return Test
            End Get
            Set(Value As Double)
                Test = Value
            End Set
        End Property

        <JsonProperty>
        Public Property QuizzesMarks As Double Implements IRecord.QuizzesMarks
            Get
                Return Quizzes
            End Get
            Set(Value As Double)
                Quizzes = Value
            End Set
        End Property

        <JsonProperty>
        Public Property ProjectMarks As Double Implements IRecord.ProjectMarks
            Get
                Return Project
            End Get
            Set(Value As Double)
                Project = Value
            End Set
        End Property

        <JsonProperty>
        Public Property ExamMarks As Double Implements IRecord.ExamMarks
            Get
                Return Exam
            End Get
            Set(Value As Double)
                Exam = Value
            End Set
        End Property

        <JsonProperty>
        Public Property ID As Integer Implements IReliability.ID
            Get
                Return Code
            End Get
            Set(Value As Integer)
                Code = Value
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
            Code = 0
        End Sub

#End Region

    End Class

End Class
