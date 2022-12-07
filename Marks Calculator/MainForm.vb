' Name:         Marks Calculator Project
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
Imports System.Data.Common
Imports MetroFramework.Controls
Imports MetroFramework.Forms
Imports MetroFramework.Drawing
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports MySql.Data.MySqlClient
Imports System.Buffers

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

    ''' <summary>
    ''' 表示表單的邊框的粗幼度
    ''' </summary>
    Private Const Border As Integer = 23

    ''' <summary>
    ''' 表示表單的邊框以及標題列的粗幼度
    ''' </summary>
    Private Const BorderWithTitle As Integer = Border + 40

    ''' <summary>
    ''' 表示進度條的闊度
    ''' </summary>
    Private Const ProgressBarWidth As Integer = 250

    ''' <summary>
    ''' 表示進度條的高度
    ''' </summary>
    Private Const ProgressBarHeight As Integer = 5

    ''' <summary>
    ''' 表示進度條水平位移的變化量
    ''' </summary>
    Private Const ProgressBarDeltaX As Integer = 5

    ''' <summary>
    ''' 表示隨機數的最大值
    ''' </summary>
    Private Const RandomNumberMaximum As Integer = 999999999

    ''' <summary>
    ''' 表示隨機數的最小值
    ''' </summary>
    Private Const RandomNumberMinimum As Integer = 100000000

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
    ''' 鎖定控制項的保留項
    ''' </summary>
    Private Reserved As List(Of (Field As FieldInfo, Value As Object))

    ''' <summary>
    ''' 用來存儲資料來源的訊息
    ''' </summary>
    Private DataSourceInfo As (Host As String, Username As String, Password As String)

    ''' <summary>
    ''' 用來存儲資料來源的執行個體
    ''' </summary>
    Private DataSourceConnection As MySqlConnection

    ''' <summary>
    ''' 用來表示數據正在鎖定的執行個體
    ''' </summary>
    Private IsDataControlsLocking As Boolean

    ''' <summary>
    ''' 表示進度條的指數
    ''' </summary>
    Private ProgressIndex As Integer

    ''' <summary>
    ''' 表示連線訊息的表單
    ''' </summary>
    Private ReadOnly FrmConnection As FrmConnect

#End Region

#Region "Constructors"

    Public Sub New()
        ' 設計工具需要此呼叫。
        InitializeComponent()
        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。
        MinimumSize = Size
        RandomNumberGenerator = New Random()
        FrmConnection = New FrmConnect(
            Function()
                Return Source
            End Function,
            Sub(Tuple As (Host As String, Username As String, Password As String))
                Source = Tuple
            End Sub
        )
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
                RandomNumber = RandomNumberGenerator.Next(RandomNumberMinimum, RandomNumberMaximum)
                Dim Count As Long = Data.LongCount(
                    Function(Record As Record) As Boolean
                        Return Record.ID = RandomNumber
                    End Function
                )
                If Count = 0 Then
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
    Private Property Selector As IEnumerable(Of (Field As FieldInfo, Value As Object))
        Get
            Return GetType(FrmMain).GetRuntimeFields().Where(
                Function(Field As FieldInfo) As Boolean
                    Return Field.FieldType = GetType(MetroButton) OrElse
                        Field.FieldType = GetType(MetroTextBox) OrElse
                        Field.FieldType = GetType(MetroCheckBox) OrElse
                        Field.FieldType = GetType(ListBox)
                End Function
            ).Select(
                Function(Field As FieldInfo) As (Field As FieldInfo, Value As Object)
                    Return (Field, Field.GetValue(Me))
                End Function
            ).Select(
                Function(Tuple As (Field As FieldInfo, Value As Object)) As (Field As FieldInfo, Value As Object)
                    Return (Tuple.Field, Tuple.Value.GetType().GetProperty("Enabled").GetValue(Tuple.Value))
                End Function
            )
        End Get
        Set(Tuples As IEnumerable(Of (Field As FieldInfo, Value As Object)))
            For Each Tuple As (Field As FieldInfo, Value As Object) In Tuples
                Tuple.Field.FieldType.GetProperty("Enabled").SetValue(Tuple.Field.GetValue(Me), Tuple.Value)
            Next
        End Set
    End Property

    ''' <summary>
    ''' 用來存儲資料來源的訊息
    ''' </summary>
    ''' <returns></returns>
    Private Property Source As (Host As String, Username As String, Password As String)
        Get
            Return DataSourceInfo
        End Get
        Set(Tuple As (Host As String, Username As String, Password As String))
            DataSourceInfo = Tuple
        End Set
    End Property

    ''' <summary>
    ''' 表示連接按鈕具有兩種狀態
    ''' </summary>
    Private Property Connection As ConnectState
        Get
            If TypeOf BtnDataSourceConnect.Tag IsNot ConnectState Then
                Throw New BranchesShouldNotBeInstantiatedException()
            End If
            Return BtnDataSourceConnect.Tag
        End Get
        Set(State As ConnectState)
            BtnDataSourceConnect.Tag = State
            If State = ConnectState.Connecting Then
                BtnDataSourceConnect.Text = "Connecting"
                BtnDataSourceConnect.Enabled = False
                Progress = True
            ElseIf State = ConnectState.Connected Then
                BtnDataSourceConnect.Text = "Disconnect"
                BtnDataSourceConnect.Enabled = True
                BtnDataSourceUpload.Enabled = True
                BtnDataSourceDownload.Enabled = True
                TxtDataSourceDatabase.Enabled = True
                TxtDataSourceTable.Enabled = True
                Progress = False
            ElseIf State = ConnectState.Disconnecting Then
                BtnDataSourceConnect.Text = "Disconnecting"
                BtnDataSourceConnect.Enabled = False
                Progress = True
            ElseIf State = ConnectState.Disconnected Then
                BtnDataSourceConnect.Text = "Connect"
                BtnDataSourceConnect.Enabled = True
                BtnDataSourceUpload.Enabled = False
                BtnDataSourceDownload.Enabled = False
                TxtDataSourceDatabase.Enabled = False
                TxtDataSourceTable.Enabled = False
                Progress = False
            Else
                Throw New BranchesShouldNotBeInstantiatedException()
            End If
        End Set
    End Property

    ''' <summary>
    ''' 表示連接按鈕正在連線的封鎖狀態
    ''' </summary>
    Private WriteOnly Property ConnectLock As Boolean
        Set(State As Boolean)
            Progress = State
            BtnDataSourceConnect.Enabled = Not State
            BtnDataSourceUpload.Enabled = Not State
            BtnDataSourceDownload.Enabled = Not State
            TxtDataSourceDatabase.ReadOnly = State
            TxtDataSourceTable.ReadOnly = State
            DataControlsLock = State
        End Set
    End Property

    ''' <summary>
    ''' 表示數據鎖定時的狀態變化
    ''' </summary>
    Private Property DataControlsLock As Boolean
        Get
            Return IsDataControlsLocking
        End Get
        Set(State As Boolean)
            IsDataControlsLocking = State
            If State Then
                BtnRecordsAdd.Enabled = False
                BtnRecordsRemove.Enabled = False
                BtnRecordsUp.Enabled = False
                BtnRecordsSquare.Enabled = False
                BtnRecordsDown.Enabled = False
                ChkRecords.Enabled = False
                TxtName.ReadOnly = True
                TxtInputTest.ReadOnly = True
                TxtInputQuizzes.ReadOnly = True
                TxtInputProject.ReadOnly = True
                TxtInputExam.ReadOnly = True
                LstRecords.Tag = IsTyping.No
            Else
                Dim Index As Integer = LstRecords.SelectedIndex
                LstRecords.SelectedIndex = -1
                LstRecords.SelectedIndex = Index
            End If
        End Set
    End Property

    ''' <summary>
    ''' 表示進度條是否顯示
    ''' </summary>
    Private Property Progress As Boolean
        Get
            Return TmrMain.Enabled
        End Get
        Set(Value As Boolean)
            TmrMain.Enabled = Value
            If Value Then
                ProgressIndex = 0
            Else
                Dim Graphics As Graphics = CreateGraphics()
                DrawProgressTrack(Graphics)
                Graphics.Dispose()
            End If
        End Set
    End Property

    ''' <summary>
    ''' 實現 Windows 視窗的最細化功能（對於 Form.FormBorderStyle 為 FormBorderStyle.None 的視窗）
    ''' </summary>
    ''' <returns></returns>
    Protected Overrides ReadOnly Property CreateParams As CreateParams
        Get
            Dim Params As CreateParams = MyBase.CreateParams
            Params.Style = Params.Style Or Native.WS_MINIMIZEBOX
            Return Params
        End Get
    End Property

#End Region

#Region "Enumerations"

    ''' <summary>
    ''' 表示表單或控制項正在輸入數據
    ''' </summary>
    Private Enum IsTyping
        Yes = 1
        No = 2
    End Enum

    ''' <summary>
    ''' 表示表單或控制項正在插入數據
    ''' </summary>
    Private Enum IsAdding
        Yes = 1
        No = 2
    End Enum

    ''' <summary>
    ''' 表示表單或控制項正在重置大小
    ''' </summary>
    Private Enum IsResizing
        Yes = 1
        No = 2
    End Enum

    ''' <summary>
    ''' 表示表單或控制項已進行初始化
    ''' </summary>
    Private Enum IsInitialized
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

    ''' <summary>
    ''' 表示目前的連線狀態
    ''' </summary>
    Public Enum ConnectState
        Connecting = 0
        Connected = 1
        Disconnecting = 2
        Disconnected = 3
    End Enum

#End Region

#Region "Methods"

    Private Sub SuspendControls()
        Progress = True
        Reserved = Selector.ToList()
        Selector = Selector.Select(
            Function(Tuple As (Field As FieldInfo, Value As Object)) As (Field As FieldInfo, Value As Object)
                Return (Tuple.Field, False)
            End Function
        )
    End Sub

    Private Sub ResumeControls()
        Progress = False
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
            Sorted.Length Mod 2 = 0,
            (Sorted(Sorted.Length / 2) + Sorted(Sorted.Length / 2 - 1)) / 2,
            Sorted((Sorted.Length - 1) / 2)
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
                IsMatched = Data(i).StudentName.Contains(TxtRecordsSearch.Text)
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

    Private Sub ShowException(Exception As Exception)
        ShowException(Exception,
            Sub()
            End Sub
        )
    End Sub

    Private Async Sub ShowException(Exception As Exception, Action As Action)
        While Tag = IsResizing.Yes
            Await Task.Run(
                Sub()
                    Thread.Sleep(1)
                End Sub
            ).ConfigureAwait(True)
        End While
        While Exception IsNot Nothing
            MessageBox.Show(Me,
                Exception.Message + Environment.NewLine +
                Environment.NewLine +
                Exception.TargetSite.ToString() + Environment.NewLine +
                Environment.NewLine +
                Exception.StackTrace.ToString() _
            , Exception.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            Exception = Exception.InnerException
        End While
        Action.Invoke()
    End Sub

    Private Async Function ShowMessage(owner As IWin32Window, text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon) As Task(Of DialogResult)
        While Tag = IsResizing.Yes
            Await Task.Run(
                Sub()
                    Thread.Sleep(1)
                End Sub
            ).ConfigureAwait(True)
        End While
        Return MessageBox.Show(owner, text, caption, buttons, icon)
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
        Data = New List(Of Record)()
        Temp = True
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
            ShowException(Exception)
            Data.Clear()
            DataFile = Nothing
        End Try
        If Not IsNotTheSameID(Data) Then
            Data.Clear()
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
            ShowException(Exception)
        End Try
        Data = Nothing
        Temp = Nothing
        ResumeControls()
    End Function

    Private Async Sub FocusMeRequest()
        If FocusMe() = False Then
            Await Task.Run(
                Sub()
                    Thread.Sleep(1)
                End Sub
            ).ConfigureAwait(True)
            FocusMeRequest()
        End If
    End Sub

    Private Sub WindowButtonsRequest(Action As Action(Of Control))
        Dim MetroFormButtonTag As Type = GetType(MetroForm).GetNestedType("WindowButtons", BindingFlags.NonPublic)
        WindowButtonsRequest(Action, MetroFormButtonTag.GetEnumValues())
    End Sub

    Private Async Sub WindowButtonsRequest(Action As Action(Of Control), Validation As Array)
        Dim MetroFormButtonType As Type = GetType(MetroForm).GetNestedType("MetroFormButton", BindingFlags.NonPublic)
        Dim IntegrityCheck(Validation.Length - 1) As (Tag As Object, Validated As Boolean)
        For i As Integer = 0 To IntegrityCheck.Length - 1
            IntegrityCheck(i) = (Validation(i), False)
        Next
        For Each Control As Control In Controls
            Dim TagIsValidated As Boolean = False
            For Each ValidationTag As Object In Validation
                If Control.Tag.GetType() = ValidationTag.GetType() AndAlso Control.Tag = ValidationTag Then
                    TagIsValidated = True
                    Exit For
                End If
            Next
            If Control.GetType() = MetroFormButtonType AndAlso TagIsValidated = True Then
                Action.Invoke(Control)
                For i As Integer = 0 To IntegrityCheck.Length - 1
                    If IntegrityCheck(i).Tag = Control.Tag Then
                        IntegrityCheck(i) = (IntegrityCheck(i).Tag, True)
                        Exit For
                    End If
                Next
            End If
        Next
        For Each Check As (Tag As Object, Validated As Boolean) In IntegrityCheck
            If Check.Validated = False Then
                Await Task.Run(
                    Sub()
                        Thread.Sleep(1)
                    End Sub
                ).ConfigureAwait(True)
                WindowButtonsRequest(Action)
                Exit For
            End If
        Next
    End Sub

    ''' <summary>
    ''' 把本地的數據上傳到資料來源中的某一個表格將會保證以下事項：<br></br>
    ''' <br></br>
    ''' 1. 如果相應的資料庫不存在則會建立一個<br></br>
    ''' 2. 如果相應的表格不存在則會建立一個<br></br>
    ''' 3. 在相應的表格當中查找表格所包含的欄位，並確保相應的欄位是存在的，以及驗證相應的欄位是具有正確的資料類型，如果不能滿足上述條件則會終止流程<br></br>
    ''' 4. 如果表格不存在相應的欄位，則會重新加入<br></br>
    ''' 5. 檢查本地的數據與資料庫的數據中具有相同 ID 的數據，其中不會篩選到資料庫的數據中任何一個欄位當中為 Null 的數據，然後把其中的 ID 記錄下來，然後檢查一筆或多筆在資料來源中的數據與本地具有相同 ID 的唯一數據，加上重複性標誌。<br></br>
    ''' 6. 如果重覆性標誌不只是數據的重複，則會詢問是否需要撇除本地數據的取代或終止流程<br></br>
    ''' 7. 把重複的數據從數據庫中移除<br></br>
    ''' 8. 把經過篩選的數據上傳到數據庫
    ''' </summary>
    ''' <returns></returns>
    Private Async Function Upload() As Task
        If Data.Count = 0 Then
            Return
        End If
        Dim SqlCommand As New StringBuilder()
        Try
            SqlCommand.Append("CREATE DATABASE IF NOT EXISTS `")
            SqlCommand.Append(TxtDataSourceDatabase.Text)
            SqlCommand.Append("`;")
            Await New MySqlCommand(SqlCommand.ToString(), DataSourceConnection).ExecuteScalarAsync().ConfigureAwait(True)
            SqlCommand.Clear()
            SqlCommand.Append("CREATE TABLE IF NOT EXISTS `")
            SqlCommand.Append(TxtDataSourceDatabase.Text)
            SqlCommand.Append("`.`")
            SqlCommand.Append(TxtDataSourceTable.Text)
            SqlCommand.Append("`( `ID` INT NOT NULL, `StudentName` TEXT NOT NULL, `Test` DOUBLE NOT NULL, `Quizzes` DOUBLE NOT NULL, `Project` DOUBLE NOT NULL, `Exam` DOUBLE NOT NULL );")
            Await New MySqlCommand(SqlCommand.ToString(), DataSourceConnection).ExecuteScalarAsync().ConfigureAwait(True)
            SqlCommand.Clear()
            SqlCommand.Append("SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '")
            SqlCommand.Append(TxtDataSourceDatabase.Text)
            SqlCommand.Append("' AND TABLE_NAME = '")
            SqlCommand.Append(TxtDataSourceTable.Text)
            SqlCommand.Append("';")
            Dim DataReader As DbDataReader = Await New MySqlCommand(SqlCommand.ToString(), DataSourceConnection).ExecuteReaderAsync().ConfigureAwait(True)
            Dim IntegrityCheck() As (Field As Object, Type As Object, Validated As Boolean) = {
                ("ID", "int", False), ("StudentName", "text", False), ("Test", "double", False), ("Quizzes", "double", False), ("Project", "double", False), ("Exam", "double", False)
            }
            While Await DataReader.ReadAsync().ConfigureAwait(True)
                For i As Integer = 0 To IntegrityCheck.Length - 1
                    If DataReader("COLUMN_NAME") = IntegrityCheck(i).Field Then
                        If DataReader("DATA_TYPE") <> IntegrityCheck(i).Type Then
                            Await ShowMessage(Me, "Some of the pairs of a field and data type are not matching!", Text, MessageBoxButtons.OK, MessageBoxIcon.Error).ConfigureAwait(True)
                            DataReader.Close()
                            Return
                        End If
                        IntegrityCheck(i) = (IntegrityCheck(i).Field, IntegrityCheck(i).Type, True)
                    End If
                Next
            End While
            DataReader.Close()
            For i As Integer = 0 To IntegrityCheck.Length - 1
                If IntegrityCheck(i).Validated = False Then
                    SqlCommand.Clear()
                    SqlCommand.Append("ALTER TABLE `")
                    SqlCommand.Append(TxtDataSourceDatabase.Text)
                    SqlCommand.Append("`.`")
                    SqlCommand.Append(TxtDataSourceTable.Text)
                    SqlCommand.Append("` ADD `")
                    SqlCommand.Append(IntegrityCheck(i).Field.ToString())
                    SqlCommand.Append("` ")
                    SqlCommand.Append(IntegrityCheck(i).Type.ToString().ToUpper())
                    SqlCommand.Append(" NOT NULL;")
                    Await New MySqlCommand(SqlCommand.ToString(), DataSourceConnection).ExecuteScalarAsync().ConfigureAwait(True)
                End If
            Next
            Dim DuplicationCheck As New List(Of (ID As Integer, IsDuplicated As Integer))()
            For Each Record As Record In Data
                SqlCommand.Clear()
                SqlCommand.Append("SELECT * FROM `")
                SqlCommand.Append(TxtDataSourceDatabase.Text)
                SqlCommand.Append("`.`")
                SqlCommand.Append(TxtDataSourceTable.Text)
                SqlCommand.Append("` WHERE `ID` = '")
                SqlCommand.Append(Record.ID.ToString())
                SqlCommand.Append("' AND `ID` IS NOT NULL AND `StudentName` IS NOT NULL AND `Test` IS NOT NULL AND `Quizzes` IS NOT NULL AND `Project` IS NOT NULL AND `Exam` IS NOT NULL;")
                DataReader = Await New MySqlCommand(SqlCommand.ToString(), DataSourceConnection).ExecuteReaderAsync().ConfigureAwait(True)
                Dim IsDuplicated As Integer = False
                While Await DataReader.ReadAsync().ConfigureAwait(True)
                    Dim Temp As Record = True
                    GetType(Record).GetProperty("StudentName").SetValue(Temp, DataReader("StudentName"))
                    GetType(Record).GetProperty("TestMarks").SetValue(Temp, DataReader("Test"))
                    GetType(Record).GetProperty("QuizzesMarks").SetValue(Temp, DataReader("Quizzes"))
                    GetType(Record).GetProperty("ProjectMarks").SetValue(Temp, DataReader("Project"))
                    GetType(Record).GetProperty("ExamMarks").SetValue(Temp, DataReader("Exam"))
                    GetType(Record).GetProperty("ID").SetValue(Temp, DataReader("ID"))
                    If Record <> Temp Then
                        DuplicationCheck.Add((Record.ID, False))
                        IsDuplicated = False
                        Exit While
                    End If
                    IsDuplicated = True
                End While
                If IsDuplicated Then
                    DuplicationCheck.Add((Record.ID, True))
                End If
                DataReader.Close()
            Next
            Dim NotSameCount As Long = DuplicationCheck.LongCount(
                Function(Check As (ID As Integer, IsDuplicated As Integer)) As Boolean
                    Return Not Check.IsDuplicated
                End Function
            )
            Dim Iterator As IEnumerable(Of Record) = Data
            If NotSameCount > 0 Then
                Dim Result As DialogResult = Await ShowMessage(Me, "Some of the local records have the same ""ID"" as one or more than one of the records of the data source, which is not matching the same fields. Would you like to replace it with the local one?", Text, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning).ConfigureAwait(True)
                If Result = DialogResult.No Then
                    Iterator = Iterator.Where(
                        Function(Record As Record) As Boolean
                            Return DuplicationCheck.LongCount(
                                Function(Duplication As (ID As Integer, IsDuplicated As Integer)) As Boolean
                                    Return Duplication.ID = Record.ID AndAlso Duplication.IsDuplicated = False
                                End Function
                            ) = 0
                        End Function
                    ).ToList()
                    DuplicationCheck = DuplicationCheck.Where(
                        Function(Duplication As (ID As Integer, IsDuplicated As Integer)) As Boolean
                            Return Duplication.IsDuplicated
                        End Function
                    ).ToList()
                ElseIf Result = DialogResult.Cancel Then
                    Return
                End If
            End If
            For Each Duplication As (ID As Integer, IsDuplicated As Integer) In DuplicationCheck
                SqlCommand.Clear()
                SqlCommand.Append("DELETE FROM `")
                SqlCommand.Append(TxtDataSourceDatabase.Text)
                SqlCommand.Append("`.`")
                SqlCommand.Append(TxtDataSourceTable.Text)
                SqlCommand.Append("` WHERE `ID` = '")
                SqlCommand.Append(Duplication.ID.ToString())
                SqlCommand.Append("' AND `ID` IS NOT NULL AND `StudentName` IS NOT NULL AND `Test` IS NOT NULL AND `Quizzes` IS NOT NULL AND `Project` IS NOT NULL AND `Exam` IS NOT NULL;")
                Await New MySqlCommand(SqlCommand.ToString(), DataSourceConnection).ExecuteScalarAsync().ConfigureAwait(True)
            Next
            For Each Record As Record In Iterator
                SqlCommand.Clear()
                SqlCommand.Append("INSERT INTO `")
                SqlCommand.Append(TxtDataSourceDatabase.Text)
                SqlCommand.Append("`.`")
                SqlCommand.Append(TxtDataSourceTable.Text)
                SqlCommand.Append("`( `ID`, `StudentName`, `Test`, `Quizzes`, `Project`, `Exam` ) VALUES ( '")
                SqlCommand.Append(Record.ID.ToString())
                SqlCommand.Append("', '")
                SqlCommand.Append(Record.StudentName)
                SqlCommand.Append("', '")
                SqlCommand.Append(Record.TestMarks.ToString())
                SqlCommand.Append("', '")
                SqlCommand.Append(Record.QuizzesMarks.ToString())
                SqlCommand.Append("', '")
                SqlCommand.Append(Record.ProjectMarks.ToString())
                SqlCommand.Append("', '")
                SqlCommand.Append(Record.ExamMarks.ToString())
                SqlCommand.Append("' );")
                Await New MySqlCommand(SqlCommand.ToString(), DataSourceConnection).ExecuteScalarAsync().ConfigureAwait(True)
            Next
        Catch Exception As Exception
            ShowException(Exception)
        End Try
    End Function

    ''' <summary>
    ''' 把資料來源中的某一個表格的數據下載到本地將會保證以下事項：<br></br>
    ''' <br></br>
    ''' 1. 在相應的表格當中查找表格所包含的欄位，並確保相應的欄位是存在的，以及驗證相應的欄位是具有正確的資料類型，如果不能滿足上述條件則會終止流程<br></br>
    ''' 2. 如果表格不存在，則會終止流程<br></br>
    ''' 3. 如果表格不存在相應的欄位，則會終止流程<br></br>
    ''' 4. 從資料來源當中提取數據，其中不會篩選到資料庫的數據中任何一個欄位當中為 Null 的數據<br></br>
    ''' 5. 把已提取的數據進行篩選，如果部分數據具有相同的 ID 並且其他的欄位完全相同，則會保留在已提取的數據中，否則把這些已提取數據的從中移除，並把這個共同的 ID 存放在撇除的列表當中<br></br>
    ''' 6. 檢查本地的數據與已提取的數據中具有相同 ID 的數據，然後把其中的 ID 記錄下來，然後檢查已提取的數據與本地具有相同 ID 的數據，加上重複性標誌。<br></br>
    ''' 7. 如果重覆性標誌不只是數據的重複，則會詢問是否需要撇除已提取數據的取代或終止流程<br></br>
    ''' 8. 把經過篩選的已提取數據寫入到本地的數據
    ''' </summary>
    ''' <returns></returns>
    Private Async Function Download() As Task
        Try
            Dim SqlCommand As New StringBuilder()
            SqlCommand.Append("SELECT * FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '")
            SqlCommand.Append(TxtDataSourceDatabase.Text)
            SqlCommand.Append("' AND TABLE_NAME = '")
            SqlCommand.Append(TxtDataSourceTable.Text)
            SqlCommand.Append("';")
            Dim DataReader As DbDataReader = Await New MySqlCommand(SqlCommand.ToString(), DataSourceConnection).ExecuteReaderAsync().ConfigureAwait(True)
            Dim IntegrityCheck() As (Field As Object, Type As Object, Validated As Boolean) = {
                ("ID", "int", False), ("StudentName", "text", False), ("Test", "double", False), ("Quizzes", "double", False), ("Project", "double", False), ("Exam", "double", False)
            }
            Dim TableValidated As Boolean = False
            While Await DataReader.ReadAsync().ConfigureAwait(True)
                TableValidated = True
                For i As Integer = 0 To IntegrityCheck.Length - 1
                    If DataReader("COLUMN_NAME") = IntegrityCheck(i).Field Then
                        If DataReader("DATA_TYPE") <> IntegrityCheck(i).Type Then
                            Await ShowMessage(Me, "Some of the pairs of a field and data type are not matching!", Text, MessageBoxButtons.OK, MessageBoxIcon.Error).ConfigureAwait(True)
                            DataReader.Close()
                            Return
                        End If
                        IntegrityCheck(i) = (IntegrityCheck(i).Field, IntegrityCheck(i).Type, True)
                    End If
                Next
            End While
            DataReader.Close()
            For i As Integer = 0 To IntegrityCheck.Length - 1
                If TableValidated = False Then
                    Await ShowMessage(Me, "The table of the data source are missing!", Text, MessageBoxButtons.OK, MessageBoxIcon.Error).ConfigureAwait(True)
                    Return
                ElseIf IntegrityCheck(i).Validated = False Then
                    Await ShowMessage(Me, "Some of the fields of the table of the data source are missing!", Text, MessageBoxButtons.OK, MessageBoxIcon.Error).ConfigureAwait(True)
                    Return
                End If
            Next
            SqlCommand.Clear()
            SqlCommand.Append("SELECT * FROM `")
            SqlCommand.Append(TxtDataSourceDatabase.Text)
            SqlCommand.Append("`.`")
            SqlCommand.Append(TxtDataSourceTable.Text)
            SqlCommand.Append("` WHERE `ID` IS NOT NULL AND `StudentName` IS NOT NULL AND `Test` IS NOT NULL AND `Quizzes` IS NOT NULL AND `Project` IS NOT NULL AND `Exam` IS NOT NULL;")
            DataReader = Await New MySqlCommand(SqlCommand.ToString(), DataSourceConnection).ExecuteReaderAsync().ConfigureAwait(True)
            Dim DownloadExcepts As New List(Of Integer)
            Dim Downloads As New List(Of Record)()
            While Await DataReader.ReadAsync().ConfigureAwait(True)
                Dim Temp As Record = True
                GetType(Record).GetProperty("StudentName").SetValue(Temp, DataReader("StudentName"))
                GetType(Record).GetProperty("TestMarks").SetValue(Temp, DataReader("Test"))
                GetType(Record).GetProperty("QuizzesMarks").SetValue(Temp, DataReader("Quizzes"))
                GetType(Record).GetProperty("ProjectMarks").SetValue(Temp, DataReader("Project"))
                GetType(Record).GetProperty("ExamMarks").SetValue(Temp, DataReader("Exam"))
                GetType(Record).GetProperty("ID").SetValue(Temp, DataReader("ID"))
                Dim DownloadExceptsCount As Long = DownloadExcepts.LongCount(
                    Function(Record As Record) As Boolean
                        Return Record.ID = Temp.ID
                    End Function
                )
                If DownloadExceptsCount = 1 Then
                    Continue While
                Else
                    Dim ExceptEnumerable As IEnumerable(Of Record) = Downloads.Where(
                        Function(Record As Record) As Boolean
                            Return Record.ID = Temp.ID
                        End Function
                    )
                    If ExceptEnumerable.LongCount() = 1 Then
                        Dim Except As Record = ExceptEnumerable.Single()
                        If Except <> Temp Then
                            Downloads.Remove(Except)
                            DownloadExcepts.Add(Except.ID)
                        End If
                        Continue While
                    Else
                        Downloads.Add(Temp)
                    End If
                End If
            End While
            DataReader.Close()
            Dim DuplicationCheck As New List(Of (ID As Integer, IsDuplicated As Integer))()
            For Each DataRecord As Record In Data
                Dim Enumerable As IEnumerable(Of Record) = Downloads.Where(
                    Function(Record As Record) As Boolean
                        Return Record.ID = DataRecord.ID
                    End Function
                )
                If Enumerable.LongCount() = 1 Then
                    DuplicationCheck.Add((Enumerable.Single().ID, Enumerable.Single() = DataRecord))
                End If
            Next
            Dim NotSameCount As Long = DuplicationCheck.LongCount(
                Function(Check As (ID As Integer, IsDuplicated As Integer)) As Boolean
                    Return Not Check.IsDuplicated
                End Function
            )
            Dim Iterator As IEnumerable(Of Record) = Downloads
            If NotSameCount > 0 Then
                Dim Result As DialogResult = Await ShowMessage(Me, "Some of the records of the data source have the same ""ID"" as one of the local records, which is not matching the same fields. Would you like to replace it with the new one?", Text, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning).ConfigureAwait(True)
                If Result = DialogResult.No Then
                    Iterator = Iterator.Where(
                        Function(Record As Record) As Boolean
                            Return DuplicationCheck.LongCount(
                                Function(Duplication As (ID As Integer, IsDuplicated As Integer)) As Boolean
                                    Return Duplication.ID = Record.ID AndAlso Duplication.IsDuplicated = False
                                End Function
                            ) = 0
                        End Function
                    ).ToList()
                    DuplicationCheck = DuplicationCheck.Where(
                        Function(Duplication As (ID As Integer, IsDuplicated As Integer)) As Boolean
                            Return Duplication.IsDuplicated
                        End Function
                    ).ToList()
                ElseIf Result = DialogResult.Cancel Then
                    Return
                End If
            End If
            For Each IteratorRecord As Record In Iterator
                Dim DuplicationEnumerable As IEnumerable(Of (ID As Integer, IsDuplicated As Integer)) = DuplicationCheck.Where(
                    Function(Tuple As (ID As Integer, IsDuplicated As Integer)) As Boolean
                        Return Tuple.ID = IteratorRecord.ID
                    End Function
                )
                If DuplicationEnumerable.LongCount() > 0 Then
                    Dim DuplicationEnumerableCount As Long = DuplicationEnumerable.LongCount(
                        Function(Tuple As (ID As Integer, IsDuplicated As Integer)) As Boolean
                            Return Tuple.IsDuplicated = False
                        End Function
                    )
                    If DuplicationEnumerableCount > 0 Then
                        Data.Remove(Data.Where(
                            Function(Record As Record) As Boolean
                                Return Record.ID = IteratorRecord.ID
                            End Function
                        ).Single())
                    Else
                        Continue For
                    End If
                End If
                Data.Add(IteratorRecord)
            Next
        Catch Exception As Exception
            ShowException(Exception)
        End Try
    End Function

    Private Sub DrawProgressTrack(Paint As Graphics)
        Dim Boundary As New Rectangle(New Point(0, 0), New Size(ClientSize.Width, ProgressBarHeight))
        Dim Brusher As New SolidBrush(MetroPaint.GetStyleColor(Style))
        Paint.FillRectangle(Brusher, Boundary)
    End Sub

    Private Sub DrawProgressBar(Paint As Graphics)
        ProgressIndex = ProgressIndex Mod (ClientSize.Width + ProgressBarWidth)
        Dim Boundary As New Rectangle(New Point(ProgressIndex - ProgressBarWidth, 0), New Size(ProgressBarWidth, ProgressBarHeight))
        Dim Brusher As New SolidBrush(Color.GreenYellow)
        Paint.FillRectangle(Brusher, Boundary)
    End Sub

#End Region

#Region "Handles"

    Private Async Sub FrmMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        State = FormState.Initializing
        LastWindowState = WindowState
        LastSize = Size
        DataSourceInfo = ("localhost", "root", "")
        Selector = Selector.Select(
            Function(Tuple As (Field As FieldInfo, Value As Object)) As (Field As FieldInfo, Value As Object)
                Return (Tuple.Field, True)
            End Function
        )
        Tag = IsResizing.No
        PanMain.Tag = IsInitialized.No
        FormBorderStyle = FormBorderStyle.Sizable
        LblInputMain.Text = "CA Components: " + Record.CAComponents
        GrpResult.Text += " [" + Record.ModuleResult + "]"
        WindowButtonsRequest(
            Sub(Control As Control)
                Control.TabStop = False
            End Sub
        )
        '（修改標題列的按鈕即 MetroForm.MetroFormButton 的屬性 Tabstop 為 False，實現對當按下按鍵 Tab 時，略過改變視窗狀態的按鈕）
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
        TxtDataSourceDatabase.Text = "marks"
        TxtDataSourceTable.Text = Date.Now.Year.ToString()
        Connection = ConnectState.Disconnected
        DataControlsLock = False
        State = FormState.LoadHasFinish
    End Sub

    Private Sub FrmMain_Paint(sender As Object, e As PaintEventArgs) Handles Me.Paint
        If PanMain.Tag = IsInitialized.No Then
            PanMain.Tag = IsInitialized.Yes
            Dim ExStyle As Integer = 0
            For Each PropertyCreateParams As PropertyInfo In GetType(Panel).GetRuntimeProperties()
                If PropertyCreateParams.Name = "CreateParams" Then
                    Dim Params As Object = PropertyCreateParams.GetValue(PanMain)
                    For Each PropertyExStyle As PropertyInfo In PropertyCreateParams.PropertyType.GetRuntimeProperties()
                        If PropertyExStyle.Name = "ExStyle" Then
                            ExStyle = CType(PropertyExStyle.GetValue(Params), Integer)
                            Exit For
                        End If
                    Next
                    Exit For
                End If
            Next
            If Native.GetWindowLong(PanMain.Handle, Native.GWL_EXSTYLE) <> ExStyle Then '（驗證 PanMain.CreateParams.ExStyle 為初始狀態）
                Throw New BranchesShouldNotBeInstantiatedException()
            End If
            Native.SetWindowLong(PanMain.Handle, Native.GWL_EXSTYLE, ExStyle Or Native.WS_EX_COMPOSITED) '（把控制項 PanMain 動態地設置其視窗風格，實現雙緩衝允許在不閃爍的情況下繪製窗口及其後代）
        End If
    End Sub

    Private Async Sub FrmMain_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If State = FormState.LoadHasFinish Then
            State = FormState.CloseHasStarted
            e.Cancel = True
            If Connection = ConnectState.Connected Then
                While DataControlsLock
                    Await Task.Run(
                        Sub()
                            Thread.Sleep(1)
                        End Sub
                    ).ConfigureAwait(True)
                End While
                BtnDataSourceConnect.PerformClick()
                While Connection <> ConnectState.Disconnected
                    Await Task.Run(
                        Sub()
                            Thread.Sleep(1)
                        End Sub
                    ).ConfigureAwait(True)
                End While
            End If
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
                SelectNextControl(TxtName, True, True, True, True)
                BtnRecordsAdd.PerformClick()
            End If
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TxtName_Leave(sender As Object, e As EventArgs) Handles TxtName.Leave
        If LstRecords.Tag = IsAdding.Yes Then
            Temp = InputedRecord
        End If
    End Sub

    Private Async Sub BtnDataSourceConnect_Click(sender As Object, e As EventArgs) Handles BtnDataSourceConnect.Click
        If Connection = ConnectState.Disconnected Then
            If FrmConnection.ShowDialog(Me) = DialogResult.Cancel Then
                Return
            End If
            Try
                DataSourceConnection = New MySqlConnection("DATASOURCE=" + DataSourceInfo.Host + ";USERNAME=" + DataSourceInfo.Username + ";PASSWORD=" + DataSourceInfo.Password + ";")
                Connection = ConnectState.Connecting
                Await Task.Run(
                    Async Function() As Task
                        Await DataSourceConnection.OpenAsync().ConfigureAwait(False)
                    End Function
                ).ConfigureAwait(True)
                Connection = ConnectState.Connected
            Catch Exception As Exception
                ShowException(Exception,
                    Sub()
                        Connection = ConnectState.Disconnected
                    End Sub
                )
            End Try
        ElseIf Connection = ConnectState.Connected Then
            Connection = ConnectState.Disconnecting
            Await Task.Run(
                Async Function() As Task
                    Await DataSourceConnection.CloseAsync().ConfigureAwait(False)
                End Function
            ).ConfigureAwait(True)
            Connection = ConnectState.Disconnected
        End If
    End Sub

    Private Async Sub BtnDataSourceUpload_Click(sender As Object, e As EventArgs) Handles BtnDataSourceUpload.Click
        ConnectLock = True
        Await Upload().ConfigureAwait(True)
        ConnectLock = False
    End Sub

    Private Async Sub BtnDataSourceDownload_Click(sender As Object, e As EventArgs) Handles BtnDataSourceDownload.Click
        ConnectLock = True
        Await Download().ConfigureAwait(True)
        Dim CaptureIndex As Integer = LstRecords.SelectedIndex
        ShowStatistics()
        RecordsSearch(
            Function() As Integer
                Return CaptureIndex
            End Function
        )
        ConnectLock = False
    End Sub

    Private Sub BtnRecordsAdd_Click(sender As Object, e As EventArgs) Handles BtnRecordsAdd.Click
        If TxtName.Text = String.Empty Then
            MessageBox.Show(Me, "The record to be inserted into the local records should be a record that has a non-empty ""StudentName"".", Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        If Not ChkRecords.Checked Then
            For Each Record As Record In Data
                If Record.StudentName = InputedRecord.StudentName Then
                    MessageBox.Show(Me, "The record to be inserted into the local records should not match the same ""StudentName"". If you would like to suppress the restriction, you have to tick out the ""Allow Duplicated Name"" checkbox.", Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
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
            MessageBox.Show(Me, "Some of the local records have the same ""StudentName""! But whenever a record is inserted into the local records, it is still avoided all these records match the same ""StudentName"".", Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
        End If
    End Sub

    Private Sub LstRecord_SelectedIndexChanged(sender As Object, e As EventArgs) Handles LstRecords.SelectedIndexChanged
        If LstRecords.SelectedIndex = -1 Then
            Return
        End If
        If DataControlsLock = False Then
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

    Private Sub TmrMain_Tick(sender As Object, e As EventArgs) Handles TmrMain.Tick
        If Tag = IsResizing.No Then
            Dim Graphics As Graphics = CreateGraphics() '（透過繪製進度條，保持視窗的標題列位置能夠捕獲相認的訊息）
            DrawProgressTrack(Graphics)
            DrawProgressBar(Graphics)
            Graphics.Dispose()
            ProgressIndex += ProgressBarDeltaX
        End If
    End Sub

    Protected Overrides Sub OnPaintBackground(e As PaintEventArgs)
        MyBase.OnPaintBackground(e)
        If Progress = True Then
            DrawProgressTrack(e.Graphics)
            DrawProgressBar(e.Graphics)
        End If
    End Sub

    Protected Overrides Sub OnPaint(e As PaintEventArgs)
        If Tag = IsResizing.No Then
            MyBase.OnPaint(e)
        End If
    End Sub

    Private Sub FrmMain_ResizeBegin(sender As Object, e As EventArgs) Handles Me.ResizeBegin
        Tag = IsResizing.Yes
        WindowButtonsRequest(
            Sub(Control As Control)
                Control.Hide()
            End Sub
        )
    End Sub

    Private Sub FrmMain_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        Tag = IsResizing.No
        WindowButtonsRequest(
            Sub(Control As Control)
                Control.Show()
            End Sub
        )
    End Sub

    Private Sub FrmMain_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        If LastWindowState <> WindowState Then
            Dim MetroFormButtonTag As Type = GetType(MetroForm).GetNestedType("WindowButtons", BindingFlags.NonPublic)
            WindowButtonsRequest(
                Sub(Control As Control)
                    If WindowState = FormWindowState.Normal Then
                        Control.Text = "1"
                    ElseIf WindowState = FormWindowState.Maximized Then
                        Control.Text = "2"
                    End If
                End Sub,
                New Object() {MetroFormButtonTag.GetEnumValues()(1)}
            )
            '（修復對於在視窗空白位置雙擊從而改變視窗狀態時，最大化或一般按鈕樣式無法改變樣式的問題）
            If WindowState = FormWindowState.Normal Then
                Owner.Show() '（對於視窗由最大化即 Form.WindowState 為 FormWindowState.Maximized 變為一般即 Form.WindowState 為 FormWindowState.Normal 會失去分層視窗之底層陰影的修復）
                FocusMeRequest() '（對於視窗由最大化即 Form.WindowState 為 FormWindowState.Maximized 變為一般即 Form.WindowState 為 FormWindowState.Normal 會失去焦點的修復）
            Else
                Size = LastSize '（大小容易受到多次觸發的改變，基於這種易失性故額外恢復原有大小）
            End If
            Resizable = WindowState <> FormWindowState.Maximized
        End If
        LastWindowState = WindowState
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        Select Case m.Msg
            Case Native.WM_NCCALCSIZE '（透過對訊息 WM_NCCALCSIZE 的捕獲，保留視窗狀態變更的動畫，其中屬性 FormBorderStyle 需要被設置為 FormBorderStyle.Sizable）
                If WindowState = FormWindowState.Maximized Then '（在最大化模式下補足表單邊界）
                    Dim XFrame As Integer = Native.GetSystemMetrics(Native.SM_CXSIZEFRAME)
                    Dim YFrame As Integer = Native.GetSystemMetrics(Native.SM_CYSIZEFRAME)
                    Dim Border As Integer = Native.GetSystemMetrics(Native.SM_CXPADDEDBORDER)
                    Dim Params As Native.NCCALCSIZE_PARAMS = Marshal.PtrToStructure(Of Native.NCCALCSIZE_PARAMS)(m.LParam)
                    Params.rgrc(0).left += XFrame + Border
                    Params.rgrc(0).top += YFrame + Border
                    Params.rgrc(0).right -= XFrame + Border
                    Params.rgrc(0).bottom -= YFrame + Border
                    Marshal.StructureToPtr(Params, m.LParam, True)
                ElseIf WindowState = FormWindowState.Normal Then
                    LastSize = Size '（大小容易受到多次觸發的改變，基於這種易失性故額外儲存原有大小）
                End If
            Case Native.WM_NCHITTEST '（透過對訊息 WM_NCHITTEST 的捕獲，實現視窗非客戶端區域拖放有效範圍的限制）
                Dim X As Integer = (m.LParam.ToInt32() And &HFFFF) - Location.X '（Message.LParam，對於 64 位元硬件平台取低 32 位的地址，低 16 位元代表滑鼠遊標的 x 座標）
                Dim Y As Integer = (m.LParam.ToInt32() >> 16) - Location.Y '（Message.LParam，對於 64 位元硬件平台取低 32 位的地址，高 16 位元代表滑鼠遊標的 y 座標）
                Dim BorderX As Integer = 0
                Dim BorderY As Integer = 0
                If WindowState = FormWindowState.Maximized Then '（在最大化模式下補足表單邊界）
                    Dim XFrame As Integer = Native.GetSystemMetrics(Native.SM_CXSIZEFRAME)
                    Dim YFrame As Integer = Native.GetSystemMetrics(Native.SM_CYSIZEFRAME)
                    Dim Border As Integer = Native.GetSystemMetrics(Native.SM_CXPADDEDBORDER)
                    BorderX = XFrame + Border
                    BorderY = YFrame + Border
                End If
                If X - BorderX >= Border AndAlso X < ClientSize.Width - Border AndAlso Y - BorderY >= BorderWithTitle AndAlso Y < ClientSize.Height - Border Then
                    m.Result = New IntPtr(Native.HTNOWHERE) '（在視窗中間的內容部分）
                    Return
                Else
                    If X >= ClientSize.Width - Border AndAlso Y >= ClientSize.Height - Border Then
                        MyBase.WndProc(m) '（在視窗右下角的大小調整部分）
                    ElseIf WindowState <> FormWindowState.Maximized OrElse Y - BorderY <= BorderWithTitle Then '（限制在最大化時只能夠在標題列拖放）
                        m.Result = New IntPtr(Native.HTCAPTION) '（在視窗標題列的部分）
                        Return
                    End If
                End If
            Case Native.WM_NCLBUTTONDBLCLICK '（透過對訊息 WM_NCLBUTTONDBLCLICK 的捕獲，實現視窗非客戶端區域雙擊有效範圍的限制）
                Dim Y As Integer = (m.LParam.ToInt32() >> 16) - Location.Y '（Message.LParam，對於 64 位元硬件平台取低 32 位的地址，高 16 位元代表滑鼠遊標的 y 座標）
                Dim BorderY As Integer = 0
                If WindowState = FormWindowState.Maximized Then '（在最大化模式下補足表單邊界）
                    Dim YFrame As Integer = Native.GetSystemMetrics(Native.SM_CYSIZEFRAME)
                    Dim Border As Integer = Native.GetSystemMetrics(Native.SM_CXPADDEDBORDER)
                    BorderY = YFrame + Border
                End If
                If Y - BorderY <= BorderWithTitle Then
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

        Friend Shared ReadOnly Property Scale(Number As Double) As String
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
        Public Const GWL_EXSTYLE As Integer = -20
        Public Const WM_NCCALCSIZE As Integer = &H83
        Public Const WM_NCHITTEST As Integer = &H84
        Public Const WM_NCLBUTTONDBLCLICK As Integer = &HA3
        Public Const HTNOWHERE As Integer = 0
        Public Const HTCAPTION As Integer = 2
        Public Const WS_MINIMIZEBOX As Integer = &H20000
        Public Const WS_EX_COMPOSITED As Integer = &H2000000

#End Region

#Region "Decorations"

        <DllImport("user32.dll")>
        Public Shared Function GetSystemMetrics(nIndex As Integer) As Integer
        End Function

        <DllImport("user32.dll")>
        Public Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
        End Function

        <DllImport("user32.dll")>
        Public Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer
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
            Public rgrc() As RECT
            Public lppos As WINDOWPOS

#End Region

        End Structure

    End Class

End Class
