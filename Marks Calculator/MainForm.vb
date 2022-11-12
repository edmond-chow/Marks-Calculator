﻿' Name:         Marks Calculator Project
' Purpose:      Used to calculate student grades.
' Programmer:   Chow Cheuk Hei, Tse Ka Yu on 2 / 10 / 2022

Imports System.IO
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Runtime.Serialization
Imports System.Security
Imports System.Text
Imports System.Text.RegularExpressions
Imports System.Threading
Imports MetroFramework.Controls
Imports MetroFramework.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq

Public Class FrmMain

#Region "Constants"

    ''' <summary>
    ''' 數據緩存的檔案名稱
    ''' </summary>
    Private Const FileName As String = "Records.json"

    ''' <summary>
    ''' 除錯時間
    ''' </summary>
    Private Const DebugTimeout As Integer = 1000

#End Region

#Region "Fields"

    ''' <summary>
    ''' 用來存儲多筆數據的執行個體
    ''' </summary>
    Private Data As List(Of Record)

    ''' <summary>
    ''' 用來存儲暫存數據的執行個體
    ''' </summary>
    Private Temp As Record

    ''' <summary>
    ''' 用來代表數據的文件流
    ''' </summary>
    Private DataFile As FileStream

    ''' <summary>
    ''' 表示目前的表單狀態
    ''' </summary>
    Private State As FormState

    ''' <summary>
    ''' 表示表單上一個視窗狀態
    ''' </summary>
    Private LastWindowState As FormWindowState

    ''' <summary>
    ''' 表示表單上一個視窗大小
    ''' </summary>
    Private LastSize As Size

    ''' <summary>
    ''' 隨機數生成的執行個體
    ''' </summary>
    Private ReadOnly RandomNumberGenerator As Random

    ''' <summary>
    ''' 用來表示變更焦點請求，這個標誌為會短暫變為 True
    ''' </summary>
    Private FocusMeRequest As Boolean

    ''' <summary>
    ''' 用來表示變更視窗按鈕設定請求，這個標誌為會短暫變為 True
    ''' </summary>
    Private WindowButtonsRequest As Boolean

    ''' <summary>
    ''' 鎖定控制項的保留項
    ''' </summary>
    Private Reserved As List(Of (FieldInfo, Object))

#End Region

#Region "Constructors"

    Public Sub New()
        ' 設計工具需要此呼叫。
        InitializeComponent()
        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。
        MinimumSize = Size
        Data = New List(Of Record)()
        Temp = True
        DataFile = Nothing
        State = FormState.Initializing
        LastWindowState = WindowState
        LastSize = Size
        RandomNumberGenerator = New Random()
        FocusMeRequest = False
        WindowButtonsRequest = False
        Reserved = Nothing
    End Sub

#End Region

#Region "Properties"

    ''' <summary>
    ''' 表示數據欄位相應的 Record
    ''' </summary>
    ''' <returns></returns>
    Private ReadOnly Property InputedRecord As Record
        Get
            Dim MyRecord As Record = (TxtName.Text, Double.Parse(TxtInputTest.Text), Double.Parse(TxtInputQuizzes.Text), Double.Parse(TxtInputProject.Text), Double.Parse(TxtInputExam.Text))
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
    ''' 表示 LstRecords 已選取項目相應的上一個 Record
    ''' </summary>
    ''' <returns></returns>
    Private Property SelectedPrevRecord As Record
        Get
            If LstRecords.SelectedIndex <= 1 Then
                Throw New BranchesShouldNotBeInstantiatedException()
            End If
            Return Data.Item(CType(TxtRecordsSearch.Tag, List(Of Integer)).Item(LstRecords.SelectedIndex - 2))
        End Get
        Set(Value As Record)
            If LstRecords.SelectedIndex <= 1 Then
                Throw New BranchesShouldNotBeInstantiatedException()
            End If
            Data.Item(CType(TxtRecordsSearch.Tag, List(Of Integer)).Item(LstRecords.SelectedIndex - 2)) = Value
        End Set
    End Property

    ''' <summary>
    ''' 表示 LstRecords 已選取項目相應的 Record
    ''' </summary>
    ''' <returns></returns>
    Private Property SelectedRecord As Record
        Get
            If LstRecords.SelectedIndex = -1 Then
                Throw New BranchesShouldNotBeInstantiatedException()
            End If
            If LstRecords.SelectedIndex = 0 Then
                Return Temp
            End If
            Return Data.Item(CType(TxtRecordsSearch.Tag, List(Of Integer)).Item(LstRecords.SelectedIndex - 1))
        End Get
        Set(Value As Record)
            If LstRecords.SelectedIndex = -1 Then
                Throw New BranchesShouldNotBeInstantiatedException()
            End If
            If LstRecords.SelectedIndex = 0 Then
                Temp = Value
            End If
            Data.Item(CType(TxtRecordsSearch.Tag, List(Of Integer)).Item(LstRecords.SelectedIndex - 1)) = Value
        End Set
    End Property

    ''' <summary>
    ''' 表示 LstRecords 已選取項目相應的下一個 Record
    ''' </summary>
    ''' <returns></returns>
    Private Property SelectedNextRecord As Record
        Get
            If LstRecords.SelectedIndex >= LstRecords.Items.Count - 1 OrElse LstRecords.SelectedIndex <= 0 Then
                Throw New BranchesShouldNotBeInstantiatedException()
            End If
            Return Data.Item(CType(TxtRecordsSearch.Tag, List(Of Integer)).Item(LstRecords.SelectedIndex))
        End Get
        Set(Value As Record)
            If LstRecords.SelectedIndex >= LstRecords.Items.Count - 1 OrElse LstRecords.SelectedIndex <= 0 Then
                Throw New BranchesShouldNotBeInstantiatedException()
            End If
            Data.Item(CType(TxtRecordsSearch.Tag, List(Of Integer)).Item(LstRecords.SelectedIndex)) = Value
        End Set
    End Property

    ''' <summary>
    ''' 鎖定控制項的篩選器
    ''' </summary>
    ''' <returns></returns>
    Private Property Selector As IEnumerable(Of (FieldInfo, Object))
        Get
            Return GetType(FrmMain).GetRuntimeFields().Where(
                Function(Field As FieldInfo) As Boolean
                    Return Field.FieldType = GetType(MetroButton) OrElse
                        Field.FieldType = GetType(MetroTextBox) OrElse
                        Field.FieldType = GetType(MetroCheckBox) OrElse
                        Field.FieldType = GetType(ListBox)
                End Function
            ).Select(
                Function(Field As FieldInfo) As (FieldInfo, Object)
                    Return (Field, Field.GetValue(Me))
                End Function
            ).Select(
                Function(Tuple As (Field As FieldInfo, Value As Object)) As (FieldInfo, Object)
                    Return (Tuple.Field, Tuple.Value.GetType().GetProperty("Enabled").GetValue(Tuple.Value))
                End Function
            )
        End Get
        Set(Tuples As IEnumerable(Of (FieldInfo, Object)))
            For Each Tuple As (Field As FieldInfo, Value As Object) In Tuples
                Tuple.Field.FieldType.GetProperty("Enabled").SetValue(Tuple.Field.GetValue(Me), Tuple.Value)
            Next
        End Set
    End Property

    ''' <summary>
    ''' 實現 Windows 視窗的最細化功能（對於 Form.FormBorderStyle 為 FormBorderStyle.None 的視窗）
    ''' </summary>
    ''' <returns></returns>
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
    ''' 表示控制項正在輸入數據
    ''' </summary>
    Private Enum IsTyping
        Yes = 1
        No = 2
    End Enum

    ''' <summary>
    ''' 表示控制項正在插入數據
    ''' </summary>
    Private Enum IsAdding
        Yes = 1
        No = 2
    End Enum

    ''' <summary>
    ''' 表示目前的表單狀態
    ''' </summary>
    Private Enum FormState
        Initializing = 1
        LoadHasFinish = 2
        CloseHasStarted = 3
        Finalizing = 4
    End Enum

#End Region

#Region "Methods"

    Private Sub SuspendControls()
        PrbMain.Show()
        Reserved = Selector.ToList()
        Selector = Selector.Select(
            Function(Tuple As (Field As FieldInfo, Object)) As (FieldInfo, Object)
                Return (Tuple.Field, False)
            End Function
        )
    End Sub

    Private Sub ResumeControls()
        PrbMain.Hide()
        Selector = Reserved
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
        Dim Sorted() As Double = Data.Select(
            Function(Record As Record) As Double
                Return Record.ModuleMarks
            End Function
        ).ToArray()
        Array.Sort(Sorted)
        Dim Md As Double =
        If(
            Sorted.LongLength Mod 2 = 0,
            (Sorted(Sorted.LongLength / 2) + Sorted(Sorted.LongLength / 2 - 1)) / 2,
            Sorted((Sorted.LongLength - 1) / 2)
        )
        TxtStatisticsMd.Text = Md.ToString()
    End Sub

    Private Sub RecordsSearch(IndexFunc As IndexFunc)
        LstRecords.Items.Clear()
        LstRecords.Items.Add("(Input)")
        TxtRecordsSearch.Tag = New List(Of Integer)()
        For i As Integer = 0 To Data.Count - 1
            Dim IsMatched As Boolean = False
            If ChkRecordsSearch.Checked Then
                Try
                    IsMatched = Regex.IsMatch(Data.Item(i).StudentName, TxtRecordsSearch.Text)
                Catch Exception As Exception
                End Try
            Else
                IsMatched = Data(i).StudentName.IndexOf(TxtRecordsSearch.Text) <> -1
            End If
            If IsMatched Then
                LstRecords.Items.Add(Data.Item(i).StudentName)
                CType(TxtRecordsSearch.Tag, List(Of Integer)).Add(i)
            End If
        Next
        LstRecords.SelectedIndex = IndexFunc.Invoke()
    End Sub

    Private Function RecordsIsSorted() As Boolean
        For i As Integer = 0 To Data.Count - 2
            If Not Data.ElementAt(i).CompareTo(Data.ElementAt(i + 1)) <= 0 Then
                Return False
            End If
        Next
        Return True
    End Function

    Private Shared Function IsNotTheSameID(Enumerable As IEnumerable(Of IReliability)) As Boolean
        For i As Integer = 0 To Enumerable.Count() - 2
            For j As Integer = 1 To Enumerable.Count() - 1 - i
                If Enumerable.ElementAt(i).ID = Enumerable.ElementAt(i + j).ID Then
                    Return False
                End If
            Next
        Next
        Return True
    End Function

    Private Shared Function IsNotTheSame(Enumerable As IEnumerable(Of IRecord)) As Boolean
        For i As Integer = 0 To Enumerable.Count() - 2
            For j As Integer = 1 To Enumerable.Count() - 1 - i
                If Enumerable.ElementAt(i).StudentName = Enumerable.ElementAt(i + j).StudentName Then
                    Return False
                End If
            Next
        Next
        Return True
    End Function

    Private Shared Async Function DebugTest() As Task
        If Environment.CommandLine.Contains("debug") Then
            Await Task.Run(
                Sub()
                    Thread.Sleep(DebugTimeout)
                End Sub
            ).ConfigureAwait(False)
        End If
    End Function

    Private Async Function ReadDataFile() As Task
        SuspendControls()
        Try
            If File.Exists(FileName) Then
                DataFile = File.Open(FileName, FileMode.Open)
                If DataFile.Length = 0 Then
                    Throw New BranchesShouldNotBeInstantiatedException()
                End If
                Dim Json(DataFile.Length - 1) As Byte
                Await DataFile.ReadAsync(Json, 0, DataFile.Length).ConfigureAwait(True)
                For Each RecordToken As JToken In JsonConvert.DeserializeObject(Of JArray)(Encoding.UTF8.GetString(Json)).Children()
                    Dim TempRecord As Record = True
                    For Each Field As JProperty In RecordToken
                        Dim PropertyInfo As PropertyInfo = GetType(Record).GetProperty(Field.Name)
                        If PropertyInfo.PropertyType = GetType(String) Then
                            PropertyInfo.SetValue(TempRecord, Field.Value.ToString())
                        ElseIf PropertyInfo.PropertyType = GetType(Double) Then
                            PropertyInfo.SetValue(TempRecord, Double.Parse(Field.Value.ToString()))
                        ElseIf PropertyInfo.PropertyType = GetType(Integer) Then
                            PropertyInfo.SetValue(TempRecord, Integer.Parse(Field.Value.ToString()))
                        Else
                            Throw New BranchesShouldNotBeInstantiatedException()
                        End If
                    Next
                    Data.Add(TempRecord)
                Next
            Else
                DataFile = File.Create(FileName)
            End If
        Catch Exception As Exception
            Data = New List(Of Record)()
        End Try
        If Not IsNotTheSameID(Data) Then
            Data = New List(Of Record)()
        End If
        Await DebugTest().ConfigureAwait(True)
        ResumeControls()
    End Function

    Private Async Function WriteDataFile() As Task
        SuspendControls()
        Await DebugTest().ConfigureAwait(True)
        Try
            If DataFile Is Nothing Then
                Throw New BranchesShouldNotBeInstantiatedException()
            End If
            Dim Json() As Byte = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(Data, Formatting.Indented))
            DataFile.SetLength(0)
            Await DataFile.WriteAsync(Json, 0, Json.Length).ConfigureAwait(True)
            DataFile.Close()
        Catch Exception As Exception
        End Try
        ResumeControls()
    End Function

#End Region

#Region "Handles"

    Private Async Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Selector = Selector.Select(
            Function(Tuple As (Field As FieldInfo, Object)) As (FieldInfo, Object)
                Return (Tuple.Field, True)
            End Function
        )
        FormBorderStyle = FormBorderStyle.Sizable
        LblInputMain.Text = "CA Components: " + Record.CAComponents
        GrpResult.Text += " [" + Record.ModuleResult + "]"
        PrbMain.ProgressBarStyle = ProgressBarStyle.Marquee
        Await ReadDataFile().ConfigureAwait(True)
        ShowStatistics()
        RecordsSearch(
            Function() As Integer
                Return 0
            End Function
        )
        If Not IsNotTheSame(Data) Then
            ChkRecords.Checked = True
        End If
        WindowButtonsRequest = True
        State = FormState.LoadHasFinish
    End Sub

    Private Async Sub FrmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If State = FormState.LoadHasFinish Then
            State = FormState.CloseHasStarted
            e.Cancel = True
            Await WriteDataFile().ConfigureAwait(True)
            State = FormState.Finalizing
            Close()
        ElseIf State = FormState.CloseHasStarted Then
            e.Cancel = True
        End If
    End Sub

    Private Sub TxtInput_Enter(sender As Object, e As EventArgs) Handles TxtInputTest.Enter, TxtInputQuizzes.Enter, TxtInputProject.Enter, TxtInputExam.Enter
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New BranchesShouldNotBeInstantiatedException()
        End If
        If LstRecords.Tag = IsAdding.Yes Then
            CType(sender, MetroTextBox).Tag = IsTyping.Yes
        End If
    End Sub

    Private Sub TxtInput_TextChanged(sender As Object, e As EventArgs) Handles TxtInputTest.TextChanged, TxtInputQuizzes.TextChanged, TxtInputProject.TextChanged, TxtInputExam.TextChanged
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New BranchesShouldNotBeInstantiatedException()
        End If
        If LstRecords.Tag = IsAdding.Yes AndAlso CType(sender, MetroTextBox).Tag = IsTyping.Yes Then
            Dim Number As Double = 0
            ShowResult(If(Double.TryParse(CType(sender, MetroTextBox).Text, Number), InputedRecord, False))
        End If
    End Sub

    Private Sub TxtInput_Leave(sender As Object, e As EventArgs) Handles TxtInputTest.Leave, TxtInputQuizzes.Leave, TxtInputProject.Leave, TxtInputExam.Leave
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New BranchesShouldNotBeInstantiatedException()
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

    Private Sub TxtNameWithInput_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtName.KeyDown, TxtInputTest.KeyDown, TxtInputQuizzes.KeyDown, TxtInputProject.KeyDown, TxtInputExam.KeyDown
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New BranchesShouldNotBeInstantiatedException()
        End If
        If LstRecords.Tag = IsAdding.Yes AndAlso e.KeyCode = Keys.Enter Then '（實現批量資料輸入的快速 Enter 按鍵）
            If Not CType(sender, MetroTextBox).Equals(TxtInputExam) Then
                SelectNextControl(ActiveControl, True, True, True, True)
            Else
                BtnRecordsAdd.PerformClick()
                SelectNextControl(TxtName, True, True, True, True)
            End If
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TxtName_Leave(sender As Object, e As EventArgs) Handles TxtName.Leave
        If LstRecords.Tag = IsAdding.Yes Then
            Temp = InputedRecord
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
        RecordsSearch(
            Function() As Integer
                Return 0
            End Function
        )
    End Sub

    Private Sub BtnRecordsRemove_Click(sender As Object, e As EventArgs) Handles BtnRecordsRemove.Click
        Data.Remove(SelectedRecord)
        ShowStatistics()
        Dim CaptureIndex As Integer = LstRecords.SelectedIndex
        RecordsSearch(
            Function() As Integer
                Return If(CaptureIndex < LstRecords.Items.Count, CaptureIndex, LstRecords.Items.Count - 1)
            End Function
        )
    End Sub

    Private Sub BtnRecordsUp_Click(sender As Object, e As EventArgs) Handles BtnRecordsUp.Click
        Dim Temp As Record = SelectedPrevRecord
        SelectedPrevRecord = SelectedRecord
        SelectedRecord = Temp
        Dim CaptureIndex As Integer = LstRecords.SelectedIndex - 1
        RecordsSearch(
            Function() As Integer
                Return CaptureIndex
            End Function
        )
    End Sub

    Private Sub BtnRecordsSquare_Click(sender As Object, e As EventArgs) Handles BtnRecordsSquare.Click
        Data.Sort()
        Dim CaptureIndex As Integer = LstRecords.SelectedIndex
        RecordsSearch(
            Function() As Integer
                Return CaptureIndex
            End Function
        )
        BtnRecordsSquare.Enabled = False
    End Sub

    Private Sub BtnRecordsDown_Click(sender As Object, e As EventArgs) Handles BtnRecordsDown.Click
        Dim Temp As Record = SelectedNextRecord
        SelectedNextRecord = SelectedRecord
        SelectedRecord = Temp
        Dim CaptureIndex As Integer = LstRecords.SelectedIndex + 1
        RecordsSearch(
            Function() As Integer
                Return CaptureIndex
            End Function
        )
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
            BtnRecordsUp.Enabled = False
            BtnRecordsSquare.Enabled = False
            BtnRecordsDown.Enabled = False
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
            BtnRecordsUp.Enabled = LstRecords.SelectedIndex > 1
            BtnRecordsSquare.Enabled = Not RecordsIsSorted()
            BtnRecordsDown.Enabled = LstRecords.SelectedIndex < LstRecords.Items.Count - 1
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
        Dim CaptureIndex As Integer = LstRecords.SelectedIndex
        RecordsSearch(
            Function() As Integer
                Return If(CaptureIndex < LstRecords.Items.Count, CaptureIndex, LstRecords.Items.Count - 1)
            End Function
        )
    End Sub

    Private Sub ChkEnterKeys_Event(sender As Object, e As KeyEventArgs) Handles ChkRecords.KeyDown, ChkRecordsSearch.KeyDown
        If sender Is Nothing OrElse TypeOf sender IsNot MetroCheckBox Then
            Throw New BranchesShouldNotBeInstantiatedException()
        End If
        If e.KeyCode = Keys.Enter Then '（實現透過鍵盤 Enter 按鍵改變核對方塊的狀態）
            CType(sender, MetroCheckBox).Checked = Not CType(sender, MetroCheckBox).Checked
        End If
    End Sub

    Private Sub FrmMain_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If LastWindowState <> WindowState Then
            Dim MetroFormButtonType As Type = GetType(MetroForm).GetNestedType("MetroFormButton", BindingFlags.NonPublic)
            Dim MetroFormButtonTag As Type = GetType(MetroForm).GetNestedType("WindowButtons", BindingFlags.NonPublic)
            For Each Control As Control In Controls '（修復對於在視窗空白位置雙擊從而改變視窗狀態時，最大化或一般按鈕樣式無法改變樣式的問題）
                If Control.GetType() = MetroFormButtonType AndAlso Control.Tag.GetType() = MetroFormButtonTag AndAlso Control.Tag = MetroFormButtonTag.GetEnumValues()(1) Then
                    If WindowState = FormWindowState.Normal Then
                        Control.Text = "1"
                    ElseIf WindowState = FormWindowState.Maximized Then
                        Control.Text = "2"
                    End If
                End If
            Next
            If WindowState = FormWindowState.Normal Then
                FocusMeRequest = True
            Else
                Size = LastSize '（大小容易受到多次觸發的改變，基於這種易失性故額外恢復原有大小）
            End If
        End If
        LastWindowState = WindowState
    End Sub

    Private Sub TmrMain_Tick(sender As Object, e As EventArgs) Handles TmrMain.Tick
        If FocusMeRequest = True AndAlso FocusMe() Then '（對於視窗由最大化即 Form.WindowState 為 FormWindowState.Maximized 變為一般即 Form.WindowState 為 FormWindowState.Normal 會失去焦點的修復）
            FocusMeRequest = False
        End If
        If WindowButtonsRequest = True Then '（修改標題列的按鈕即 MetroForm.MetroFormButton 的屬性 Tabstop 為 False，實現對當按下按鍵 Tab 時，略過改變視窗狀態的按鈕）
            Dim MetroFormButtonType As Type = GetType(MetroForm).GetNestedType("MetroFormButton", BindingFlags.NonPublic)
            Dim MetroFormButtonTag As Type = GetType(MetroForm).GetNestedType("WindowButtons", BindingFlags.NonPublic)
            For Each Control As Control In Controls
                If Control.GetType() = MetroFormButtonType AndAlso Control.Tag.GetType() = MetroFormButtonTag Then
                    Control.TabStop = False
                    WindowButtonsRequest = False
                End If
            Next
        End If
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_NCCALCSIZE As Integer = &H83
        Const WM_NCHITTEST As Integer = &H84
        Select Case m.Msg
            Case WM_NCCALCSIZE '（透過對訊息 WM_NCCALCSIZE 的捕獲，保留視窗狀態變更的動畫，其中屬性 FormBorderStyle 需要被設置為 FormBorderStyle.Sizable）
                If WindowState = FormWindowState.Maximized Then
                    Dim X As Integer = Native.GetSystemMetrics(Native.SM_CXSIZEFRAME)
                    Dim Y As Integer = Native.GetSystemMetrics(Native.SM_CYSIZEFRAME)
                    Dim Border As Integer = Native.GetSystemMetrics(Native.SM_CXPADDEDBORDER)
                    Dim Params As Native.NCCALCSIZE_PARAMS = Marshal.PtrToStructure(Of Native.NCCALCSIZE_PARAMS)(m.LParam)
                    Params.rgrc(0).left += X + Border
                    Params.rgrc(0).top += Y + Border
                    Params.rgrc(0).right -= X + Border
                    Params.rgrc(0).bottom -= Y + Border
                    Marshal.StructureToPtr(Params, m.LParam, True)
                ElseIf WindowState = FormWindowState.Normal Then
                    LastSize = Size '（大小容易受到多次觸發的改變，基於這種易失性故額外儲存原有大小）
                End If
            Case WM_NCHITTEST '（透過對訊息 WM_NCHITTEST 的捕獲，實現視窗拖放有效範圍的限制）
                Dim X As Integer = (m.LParam.ToInt32() And &HFFFF) - Location.X '（Message.LParam，對於 64 位元硬件平台取低 32 位的地址，低 16 位元代表滑鼠遊標的 x 座標）
                Dim Y As Integer = (m.LParam.ToInt32() >> 16) - Location.Y '（Message.LParam，對於 64 位元硬件平台取低 32 位的地址，高 16 位元代表滑鼠遊標的 y 座標）
                If X >= 23 AndAlso X < Size.Width - 23 AndAlso Y >= 63 AndAlso Y < Size.Height - 23 Then
                ElseIf X < 5 OrElse X >= Size.Width - 5 OrElse Y < 5 OrElse Y >= Size.Height - 5 Then
                    If X >= Size.Width - 5 AndAlso Y >= Size.Height - 5 Then
                        MyBase.WndProc(m)
                    End If
                Else
                    MyBase.WndProc(m)
                End If
            Case Else
                MyBase.WndProc(m)
        End Select
    End Sub

#End Region

#Region "Delegates"

    Private Delegate Function IndexFunc() As Integer

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
        Implements IRecord, IReliability, IEquatable(Of Object), IComparable(Of Record), IComparable

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
                Test = Double.NaN
                Quizzes = Double.NaN
                Project = Double.NaN
                Exam = Double.NaN
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

        Public Shared ReadOnly Property Scale(Number As Double) As String
            Get
                Return (Number * 100).ToString() + "%"
            End Get
        End Property

        Public Shared ReadOnly Property CAComponents As String
            Get
                Return Enumerate(", ",
                    Enumerate(" - ", "Test", Scale(TestScale)),
                    Enumerate(" - ", "Quiz", Scale(QuizzesScale)),
                    Enumerate(" - ", "Project", Scale(ProjectScale))
                )
            End Get
        End Property

        Public Shared ReadOnly Property ModuleResult As String
            Get
                Return Enumerate(", ",
                    Enumerate(" - ", "CA", Scale(CAScale)),
                    Enumerate(" - ", "Exam", Scale(ExamScale))
                )
            End Get
        End Property

        <JsonIgnore>
        Public ReadOnly Property CAMarks As Double
            Get
                Return If(IsReal, Test * TestScale + Quizzes * QuizzesScale + Project * ProjectScale, Double.NaN)
            End Get
        End Property

        <JsonIgnore>
        Public ReadOnly Property ModuleMarks As Double
            Get
                Return If(IsReal, CAMarks * CAScale + Exam * ExamScale, Double.NaN)
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
                If Not IsReal Then
                    Return Invalid
                ElseIf CAMarks < 40 OrElse Exam < 40 Then
                    Return "F"
                ElseIf ModuleMarks >= 75 Then
                    Return "A"
                ElseIf ModuleMarks >= 65 Then
                    Return "B"
                Else
                    Return "C"
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

        Private Shared Function Enumerate(Separator As String, ParamArray Appends() As String) As String
            If Appends Is Nothing OrElse Appends.Length = 0 Then
                Return String.Empty
            End If
            Dim [Return] As String = Appends(0)
            For Each Append As String In Appends.Skip(1)
                [Return] += Separator + Append
            Next
            Return [Return]
        End Function

        Public Sub Clear()
            Name = None
            Test = 0
            Quizzes = 0
            Project = 0
            Exam = 0
            Code = 0
        End Sub

        Public Overrides Function Equals(Other As Object) As Boolean Implements IEquatable(Of Object).Equals
            Return TypeOf Other Is Record AndAlso Me = CType(Other, Record)
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return Code
        End Function

        Public Overrides Function ToString() As String
            Return MyBase.ToString() + " (" + Name.ToString() + ", " + Code.ToString() + ")"
        End Function

        Public Function CompareTo(Other As Record) As Integer Implements IComparable(Of Record).CompareTo
            Return If(Name.CompareTo(Other.Name) <> 0, Name.CompareTo(Other.Name), Code - Other.Code)
        End Function

        Public Function CompareTo(Other As Object) As Integer Implements IComparable.CompareTo
            Return If(TypeOf Other Is Record, CompareTo(CType(Other, Record)), 0)
        End Function

#End Region

#Region "Operators"

        Public Shared Operator =(Left As Record, Right As Record) As Boolean
            Return Left.Name = Right.Name AndAlso
                Left.Test = Right.Test AndAlso
                Left.Quizzes = Right.Quizzes AndAlso
                Left.Project = Right.Project AndAlso
                Left.Exam = Right.Exam AndAlso
                Left.Code = Right.Code
        End Operator

        Public Shared Operator <>(Left As Record, Right As Record) As Boolean
            Return Not Left = Right
        End Operator

        Public Shared Operator <(Left As Record, Right As Record) As Boolean
            Return Left.CompareTo(Right) < 0
        End Operator

        Public Shared Operator >(Left As Record, Right As Record) As Boolean
            Return Left.CompareTo(Right) > 0
        End Operator

        Public Shared Operator <=(Left As Record, Right As Record) As Boolean
            Return Left.CompareTo(Right) <= 0
        End Operator

        Public Shared Operator >=(Left As Record, Right As Record) As Boolean
            Return Left.CompareTo(Right) >= 0
        End Operator

        Public Shared Widening Operator CType(Valid As Boolean) As Record
            Return New Record(Valid)
        End Operator

        Public Shared Widening Operator CType(Name As String) As Record
            Return New Record(Name)
        End Operator

        Public Shared Widening Operator CType(Value As (Name As String, Test As Double, Quizzes As Double, Project As Double, Exam As Double)) As Record
            Return New Record(Value.Name, Value.Test, Value.Quizzes, Value.Project, Value.Exam)
        End Operator

        Public Shared Narrowing Operator CType(Code As Integer) As Record
            Return New Record(True) With {.ID = Code}
        End Operator

#End Region

    End Class

    <Serializable>
    Private Class BranchesShouldNotBeInstantiatedException
        Inherits NotImplementedException

#Region "Constructors"

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(message As String)
            MyBase.New(message)
        End Sub

        Public Sub New(message As String, inner As Exception)
            MyBase.New(message, inner)
        End Sub

        <SecuritySafeCritical>
        Protected Sub New(info As SerializationInfo, context As StreamingContext)
            MyBase.New(info, context)
        End Sub

#End Region

    End Class

    Private Class Native

#Region "Constants"

        Public Const SM_CXSIZEFRAME As Integer = 32
        Public Const SM_CYSIZEFRAME As Integer = 33
        Public Const SM_CXPADDEDBORDER As Integer = 92

#End Region

#Region "Decorations"

        <DllImport("user32.dll")>
        Public Shared Function GetSystemMetrics(smIndex As Integer) As Integer
        End Function

#End Region

        <StructLayout(LayoutKind.Sequential)>
        Public Structure RECT

#Region "Fields"

            Public left As Integer
            Public top As Integer
            Public right As Integer
            Public bottom As Integer

#End Region

        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Public Structure WINDOWPOS

#Region "Fields"

            Public hWnd As IntPtr
            Public hWndInsertAfter As IntPtr
            Public x As Integer
            Public y As Integer
            Public cx As Integer
            Public cy As Integer
            Public flags As UInteger

#End Region

        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Public Structure NCCALCSIZE_PARAMS

#Region "Fields"

            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=3)>
            Public rgrc As RECT()
            Public lppos As WINDOWPOS

#End Region

        End Structure

    End Class

End Class
