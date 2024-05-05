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
Imports System.Data.Common
Imports System.Net.Sockets
Imports MetroFramework.Controls
Imports MetroFramework.Forms
Imports MetroFramework.Drawing
Imports Newtonsoft.Json
Imports MySql.Data.MySqlClient

Public Class FrmMain

#Region "Constants"

    ''' <summary>
    ''' 本地數據緩存的檔案名稱
    ''' </summary>
    Private Const FileName As String = "Records.json"

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

    ''' <summary>
    ''' 表示 StringBuilder 初始容量的值
    ''' </summary>
    Private Const InitialStringCapacity As Integer = 65536

    ''' <summary>
    ''' 表示中斷一下 SuspendNow 的延遲時間
    ''' </summary>
    Private Const DelayDuration As Integer = 50

    ''' <summary>
    ''' 表示不是數字的值
    ''' </summary>
    Private Const NaN As String = "[NaN]"

#End Region

#Region "Fields"

    ''' <summary>
    ''' 用來存儲多筆已經輸入的數據
    ''' </summary>
    Private Data As List(Of Record)

    ''' <summary>
    ''' 用來暫存一筆正在輸入的數據
    ''' </summary>
    Private Temp As Record

    ''' <summary>
    ''' 用來代表本地數據緩存的文件流
    ''' </summary>
    Private DataFile As FileStream

    ''' <summary>
    ''' 表示進度條的偏移量
    ''' </summary>
    Private ProgressOffset As Integer

    ''' <summary>
    ''' 表示正在繪畫進度條
    ''' </summary>
    Private DrawingProgress As Boolean

    ''' <summary>
    ''' 表示目前的表單狀態
    ''' </summary>
    Private State As FormState

    ''' <summary>
    ''' 表示表單上一個視窗狀態（對於視窗狀態而言：修改標題列的按鈕）
    ''' </summary>
    Private LastWindowState As FormWindowState

    ''' <summary>
    ''' 表示表單上一個視窗大小（對於訊息迴圈而言：大小容易受到多次觸發的改變，基於這種易失性故額外恢復原有大小）
    ''' </summary>
    Private LastSize As Size

    ''' <summary>
    ''' 隨機數生成的執行個體
    ''' </summary>
    Private ReadOnly RandomNumberGenerator As Random

    ''' <summary>
    ''' 用來存儲資料來源的連接訊息
    ''' </summary>
    Private DataSourceInfo As (Host As String, Username As String, Password As String)

    ''' <summary>
    ''' 用來連接資料來源的執行個體
    ''' </summary>
    Private DataSourceConnection As MySqlConnection

    ''' <summary>
    ''' 用來讀取資料來源的執行個體
    ''' </summary>
    Private DataReader As DbDataReader

    ''' <summary>
    ''' 用來表示數據欄位正在鎖定的執行個體（用於防止使用者透過控制項，加入數據）
    ''' </summary>
    Private IsDataControlsLocking As Boolean

    ''' <summary>
    ''' 中斷與恢復全局鎖定控制項的保留項
    ''' </summary>
    Private Reserved As List(Of (Field As FieldInfo, Value As Object))

    ''' <summary>
    ''' 表示連線訊息的表單
    ''' </summary>
    Private ReadOnly FrmConnection As FrmConnect

    ''' <summary>
    ''' 用來暫存搜尋索引的列表
    ''' </summary>
    Private SearchIndexList As List(Of Integer)

    ''' <summary>
    ''' 用來暫存上傳與下載的錯誤代碼
    ''' </summary>
    Private ErrorCodes As Dictionary(Of ErrorKeys, Integer)

#End Region

#Region "Constructors"

    Public Sub New()
        ' 設計工具需要此呼叫。
        InitializeComponent()
        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。
        MinimumSize = Size
        Data = New List(Of Record)()
        Temp = New Record()
        RandomNumberGenerator = New Random()
        FrmConnection = New FrmConnect(
            Function() As (Host As String, Username As String, Password As String)
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
    ''' 表示滿足標準範圍的數據
    ''' </summary>
    Private ReadOnly Property RealData As IEnumerable(Of Record)
        Get
            Return Data.Where(
                Function(Record As Record) As Boolean
                    Return Record.IsReal
                End Function
            )
        End Get
    End Property

    ''' <summary>
    ''' 表示數據輸入欄位的 Record，其中會生成隨機的 ID
    ''' </summary>
    Private ReadOnly Property InputedRecord As Record
        Get
            Dim Result As Record = (TxtName.Text, Double.Parse(TxtInputTest.Text), Double.Parse(TxtInputQuizzes.Text), Double.Parse(TxtInputProject.Text), Double.Parse(TxtInputExam.Text))
            While True
                Dim RandomNumber As Integer = RandomNumberGenerator.Next(RandomNumberMinimum, RandomNumberMaximum)
                Dim Count As Long = Data.LongCount(
                    Function(Record As Record) As Boolean
                        Return Record.ID = RandomNumber
                    End Function
                )
                If Count = 0 Then
                    Result.ID = RandomNumber
                    Exit While
                End If
            End While
            Return Result
        End Get
    End Property

    ''' <summary>
    ''' 表示 LstRecords 已選取項目的上一個 Record，透過 SearchIndexList 定位 Data 中滿足搜尋結果的元素
    ''' </summary>
    Private Property SelectedPrevRecord As Record
        Get
            If LstRecords.SelectedIndex <= 1 Then
                Throw New BranchesShouldNotBeInstantiatedException("Index out of range!")
            End If
            Return Data(SearchIndexList(LstRecords.SelectedIndex - 2))
        End Get
        Set(Value As Record)
            If LstRecords.SelectedIndex <= 1 Then
                Throw New BranchesShouldNotBeInstantiatedException("Index out of range!")
            End If
            Data(SearchIndexList(LstRecords.SelectedIndex - 2)) = Value
        End Set
    End Property

    ''' <summary>
    ''' 表示 LstRecords 已選取項目的 Record，透過 SearchIndexList 定位 Data 中滿足搜尋結果的元素（當 LstRecords 選取首筆項目時會回傳 Temp）
    ''' </summary>
    Private Property SelectedRecord As Record
        Get
            If LstRecords.SelectedIndex = -1 Then
                Throw New BranchesShouldNotBeInstantiatedException("Index out of range!")
            End If
            If LstRecords.SelectedIndex = 0 Then
                Return Temp
            End If
            Return Data(SearchIndexList(LstRecords.SelectedIndex - 1))
        End Get
        Set(Value As Record)
            If LstRecords.SelectedIndex = -1 Then
                Throw New BranchesShouldNotBeInstantiatedException("Index out of range!")
            End If
            If LstRecords.SelectedIndex = 0 Then
                Temp = Value
            End If
            Data(SearchIndexList(LstRecords.SelectedIndex - 1)) = Value
        End Set
    End Property

    ''' <summary>
    ''' 表示 LstRecords 已選取項目的下一個 Record，透過 SearchIndexList 定位 Data 中滿足搜尋結果的元素
    ''' </summary>
    Private Property SelectedNextRecord As Record
        Get
            If LstRecords.SelectedIndex >= LstRecords.Items.Count - 1 OrElse LstRecords.SelectedIndex <= 0 Then
                Throw New BranchesShouldNotBeInstantiatedException("Index out of range!")
            End If
            Return Data(SearchIndexList(LstRecords.SelectedIndex))
        End Get
        Set(Value As Record)
            If LstRecords.SelectedIndex >= LstRecords.Items.Count - 1 OrElse LstRecords.SelectedIndex <= 0 Then
                Throw New BranchesShouldNotBeInstantiatedException("Index out of range!")
            End If
            Data(SearchIndexList(LstRecords.SelectedIndex)) = Value
        End Set
    End Property

    ''' <summary>
    ''' 中斷與恢復全局鎖定控制項的篩選器
    ''' </summary>
    Private Property Selector As IEnumerable(Of (Field As FieldInfo, Value As Object))
        Get
            Return GetType(FrmMain).GetRuntimeFields().Where(
                Function(Field As FieldInfo) As Boolean
                    Return Field.FieldType.IsSubclassOf(GetType(Control))
                End Function
            ).Select(
                Function(Field As FieldInfo) As (Field As FieldInfo, Value As Object)
                    Return (Field, Field.GetValue(Me).GetType().GetProperty("Enabled").GetValue(Field.GetValue(Me)))
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
    ''' 用來存儲資料來源的連接訊息
    ''' </summary>
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
                Throw New BranchesShouldNotBeInstantiatedException("Type not matching!")
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
                Throw New BranchesShouldNotBeInstantiatedException("Switch case out of bound!")
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
    ''' 表示數據欄位鎖定時的狀態變化（用於防止使用者透過控制項，加入數據）
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
    ''' 連線至資料庫的 Sql 指令
    ''' </summary>
    Private ReadOnly Property ConnectionCmd As String
        Get
            Return "DATASOURCE = " + DataSourceInfo.Host + "; USERNAME = " + DataSourceInfo.Username + "; PASSWORD = " + DataSourceInfo.Password + "; ALLOW USER VARIABLES = True; "
        End Get
    End Property

    ''' <summary>
    ''' 從資料庫上傳數據的 Sql 指令
    ''' </summary>
    Private ReadOnly Property UploadCmd As String
        Get
            Dim OpFT As String = ErrorCodes(ErrorKeys.NonExistFieldWithType).ToString()
            Dim OpST As String = ErrorCodes(ErrorKeys.InvalidSourceAndTable).ToString()
            Dim NoOp As String = ErrorCodes(ErrorKeys.NoOperationState).ToString()
            Dim Db As String = TxtDataSourceDatabase.Text
            Dim Tb As String = TxtDataSourceTable.Text
            Dim Nl As String = Environment.NewLine
            Dim Result As New StringBuilder(InitialStringCapacity)
            Result.Append("SET @Count = ").Append(Data.Count.ToString()).Append("; ").Append(Nl)
            Result.Append("IF @Count > 0 THEN ").Append(Nl)
            Result.Append("    IF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.SCHEMATA WHERE SCHEMA_NAME = '").Append(Db).Append("' ) THEN ").Append(Nl)
            Result.Append("        CREATE DATABASE `").Append(Db).Append("`; ").Append(Nl)
            Result.Append("    END IF; ").Append(Nl)
            Result.Append("    IF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' ) THEN ").Append(Nl)
            Result.Append("        CREATE TABLE `").Append(Db).Append("`.`").Append(Tb).Append("` ( `ID` INT NOT NULL, `StudentName` TEXT NOT NULL, `Test` DOUBLE NOT NULL, `Quizzes` DOUBLE NOT NULL, `Project` DOUBLE NOT NULL, `Exam` DOUBLE NOT NULL ); ").Append(Nl)
            Result.Append("    END IF; ").Append(Nl)
            Result.Append("    IF EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'ID' AND NOT UPPER (DATA_TYPE) = 'INT' ) THEN ").Append(Nl)
            Result.Append("        SELECT ").Append(OpFT).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("    ELSEIF EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'StudentName' AND NOT UPPER (DATA_TYPE) = 'TEXT' ) THEN ").Append(Nl)
            Result.Append("        SELECT ").Append(OpFT).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("    ELSEIF EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'Test' AND NOT UPPER (DATA_TYPE) = 'DOUBLE' ) THEN ").Append(Nl)
            Result.Append("        SELECT ").Append(OpFT).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("    ELSEIF EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'Quizzes' AND NOT UPPER (DATA_TYPE) = 'DOUBLE' ) THEN ").Append(Nl)
            Result.Append("        SELECT ").Append(OpFT).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("    ELSEIF EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'Project' AND NOT UPPER (DATA_TYPE) = 'DOUBLE' ) THEN ").Append(Nl)
            Result.Append("        SELECT ").Append(OpFT).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("    ELSEIF EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'Exam' AND NOT UPPER (DATA_TYPE) = 'DOUBLE' ) THEN ").Append(Nl)
            Result.Append("        SELECT ").Append(OpFT).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("    ELSEIF EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND UPPER (COLUMN_KEY) = 'PRI' AND NOT COLUMN_NAME = 'ID' ) THEN ").Append(Nl)
            Result.Append("        SELECT ").Append(OpST).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("    ELSE ").Append(Nl)
            Result.Append("        IF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'ID' ) THEN ").Append(Nl)
            Result.Append("            ALTER TABLE `").Append(Db).Append("`.`").Append(Tb).Append("` ADD `ID` INT NOT NULL; ").Append(Nl)
            Result.Append("        END IF; ").Append(Nl)
            Result.Append("        IF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'StudentName' ) THEN ").Append(Nl)
            Result.Append("            ALTER TABLE `").Append(Db).Append("`.`").Append(Tb).Append("` ADD `StudentName` TEXT NOT NULL; ").Append(Nl)
            Result.Append("        END IF; ").Append(Nl)
            Result.Append("        IF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'Test' ) THEN ").Append(Nl)
            Result.Append("            ALTER TABLE `").Append(Db).Append("`.`").Append(Tb).Append("` ADD `Test` DOUBLE NOT NULL; ").Append(Nl)
            Result.Append("        END IF; ").Append(Nl)
            Result.Append("        IF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'Quizzes' ) THEN ").Append(Nl)
            Result.Append("            ALTER TABLE `").Append(Db).Append("`.`").Append(Tb).Append("` ADD `Quizzes` DOUBLE NOT NULL; ").Append(Nl)
            Result.Append("        END IF; ").Append(Nl)
            Result.Append("        IF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'Project' ) THEN ").Append(Nl)
            Result.Append("            ALTER TABLE `").Append(Db).Append("`.`").Append(Tb).Append("` ADD `Project` DOUBLE NOT NULL; ").Append(Nl)
            Result.Append("        END IF; ").Append(Nl)
            Result.Append("        IF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'Exam' ) THEN ").Append(Nl)
            Result.Append("            ALTER TABLE `").Append(Db).Append("`.`").Append(Tb).Append("` ADD `Exam` DOUBLE NOT NULL; ").Append(Nl)
            Result.Append("        END IF; ").Append(Nl)
            Result.Append("        IF EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'ID' AND UPPER (COLUMN_KEY) = 'PRI' ) THEN ").Append(Nl)
            For Each Record In Data
                Result.Append("            IF NOT EXISTS ( SELECT NULL FROM `").Append(Db).Append("`.`").Append(Tb).Append("` WHERE `ID` = '").Append(Record.ID.ToString()).Append("' ) THEN ").Append(Nl)
                Result.Append("                INSERT INTO `").Append(Db).Append("`.`").Append(Tb).Append("` ( `ID`, `StudentName`, `Test`, `Quizzes`, `Project`, `Exam` ) VALUES ( '").Append(Record.ID.ToString()).Append("', '").Append(Record.StudentName).Append("', '").Append(Record.TestMarks.ToString()).Append("', '").Append(Record.QuizzesMarks.ToString()).Append("', '").Append(Record.ProjectMarks.ToString()).Append("', '").Append(Record.ExamMarks.ToString()).Append("' ); ").Append(Nl)
                Result.Append("            ELSE ").Append(Nl)
                Result.Append("                UPDATE `").Append(Db).Append("`.`").Append(Tb).Append("` SET `StudentName` = '").Append(Record.StudentName).Append("', `Test` = '").Append(Record.TestMarks.ToString()).Append("', `Quizzes` = '").Append(Record.QuizzesMarks.ToString()).Append("', `Project` = '").Append(Record.ProjectMarks.ToString()).Append("', `Exam` = '").Append(Record.ExamMarks.ToString()).Append("' WHERE `ID` = '").Append(Record.ID.ToString()).Append("'; ").Append(Nl)
                Result.Append("            END IF; ").Append(Nl)
            Next
            Result.Append("            SELECT NULL; ").Append(Nl)
            Result.Append("        ELSE ").Append(Nl)
            For Each Record In Data
                Result.Append("            IF NOT EXISTS ( SELECT NULL FROM `").Append(Db).Append("`.`").Append(Tb).Append("` WHERE `ID` = '").Append(Record.ID.ToString()).Append("' AND `StudentName` = '").Append(Record.StudentName).Append("' AND `Test` = '").Append(Record.TestMarks.ToString()).Append("' AND `Quizzes` = '").Append(Record.QuizzesMarks.ToString()).Append("' AND `Project` = '").Append(Record.ProjectMarks.ToString()).Append("' AND `Exam` = '").Append(Record.ExamMarks.ToString()).Append("' ) THEN ").Append(Nl)
                Result.Append("                INSERT INTO `").Append(Db).Append("`.`").Append(Tb).Append("` ( `ID`, `StudentName`, `Test`, `Quizzes`, `Project`, `Exam` ) VALUES ( '").Append(Record.ID.ToString()).Append("', '").Append(Record.StudentName).Append("', '").Append(Record.TestMarks.ToString()).Append("', '").Append(Record.QuizzesMarks.ToString()).Append("', '").Append(Record.ProjectMarks.ToString()).Append("', '").Append(Record.ExamMarks.ToString()).Append("' ); ").Append(Nl)
                Result.Append("            END IF; ").Append(Nl)
            Next
            Result.Append("            SELECT NULL; ").Append(Nl)
            Result.Append("        END IF; ").Append(Nl)
            Result.Append("    END IF; ").Append(Nl)
            Result.Append("ELSE ").Append(Nl)
            Result.Append("    SELECT ").Append(NoOp).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("END IF; ").Append(Nl)
            Return Result.ToString()
        End Get
    End Property

    ''' <summary>
    ''' 從資料庫下載數據的 Sql 指令
    ''' </summary>
    Private ReadOnly Property DownloadCmd As String
        Get
            Dim OpFT As String = ErrorCodes(ErrorKeys.NonExistFieldWithType).ToString()
            Dim OpST As String = ErrorCodes(ErrorKeys.InvalidSourceAndTable).ToString()
            Dim NoOp As String = ErrorCodes(ErrorKeys.NoOperationState).ToString()
            Dim Db As String = TxtDataSourceDatabase.Text
            Dim Tb As String = TxtDataSourceTable.Text
            Dim Nl As String = Environment.NewLine
            Dim Result As New StringBuilder(InitialStringCapacity)
            Result.Append("IF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND TABLE_TYPE = 'BASE TABLE' ) THEN ").Append(Nl)
            Result.Append("    SELECT ").Append(OpST).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("ELSEIF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'ID' AND UPPER (DATA_TYPE) = 'INT' ) THEN ").Append(Nl)
            Result.Append("    SELECT ").Append(OpFT).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("ELSEIF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'StudentName' AND UPPER (DATA_TYPE) = 'TEXT' ) THEN ").Append(Nl)
            Result.Append("    SELECT ").Append(OpFT).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("ELSEIF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'Test' AND UPPER (DATA_TYPE) = 'DOUBLE' ) THEN ").Append(Nl)
            Result.Append("    SELECT ").Append(OpFT).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("ELSEIF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'Quizzes' AND UPPER (DATA_TYPE) = 'DOUBLE' ) THEN ").Append(Nl)
            Result.Append("    SELECT ").Append(OpFT).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("ELSEIF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'Project' AND UPPER (DATA_TYPE) = 'DOUBLE' ) THEN ").Append(Nl)
            Result.Append("    SELECT ").Append(OpFT).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("ELSEIF NOT EXISTS ( SELECT NULL FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_SCHEMA = '").Append(Db).Append("' AND TABLE_NAME = '").Append(Tb).Append("' AND COLUMN_NAME = 'Exam' AND UPPER (DATA_TYPE) = 'DOUBLE' ) THEN ").Append(Nl)
            Result.Append("    SELECT ").Append(OpFT).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("ELSEIF NOT EXISTS ( SELECT NULL FROM `").Append(Db).Append("`.`").Append(Tb).Append("` WHERE `ID` IS NOT NULL AND `StudentName` IS NOT NULL AND `Test` IS NOT NULL AND `Quizzes` IS NOT NULL AND `Project` IS NOT NULL AND `Exam` IS NOT NULL ) THEN ").Append(Nl)
            Result.Append("    SELECT ").Append(NoOp).Append(" AS ERROR_CODE; ").Append(Nl)
            Result.Append("ELSE ").Append(Nl)
            Result.Append("    SELECT DISTINCT * FROM `").Append(Db).Append("`.`").Append(Tb).Append("` WHERE `ID` IS NOT NULL AND `StudentName` IS NOT NULL AND `Test` IS NOT NULL AND `Quizzes` IS NOT NULL AND `Project` IS NOT NULL AND `Exam` IS NOT NULL; ").Append(Nl)
            Result.Append("END IF; ").Append(Nl)
            Return Result.ToString()
        End Get
    End Property

    ''' <summary>
    ''' 表示進度條是否顯示
    ''' </summary>
    Private Property Progress As Boolean
        Get
            Return DrawingProgress
        End Get
        Set(Value As Boolean)
            DrawingProgress = Value
            If Value Then
                ProgressOffset = 0
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
    Private Enum ConnectState
        Connecting = 0
        Connected = 1
        Disconnecting = 2
        Disconnected = 3
    End Enum

    ''' <summary>
    ''' 表示上傳與下載的錯誤鍵值
    ''' </summary>
    Private Enum ErrorKeys
        NonExistFieldWithType = 1
        InvalidSourceAndTable = 2
        NoOperationState = 3
    End Enum

#End Region

#Region "Methods"

    ''' <summary>
    ''' 中斷全局鎖定控制項
    ''' </summary>
    Private Sub SuspendControls()
        Progress = True
        Reserved = Selector.ToList()
        Selector = Selector.Select(
            Function(Tuple As (Field As FieldInfo, Value As Object)) As (Field As FieldInfo, Value As Object)
                Return (Tuple.Field, False)
            End Function
        )
    End Sub

    ''' <summary>
    ''' 恢復全局鎖定控制項
    ''' </summary>
    Private Sub ResumeControls()
        Progress = False
        Selector = Reserved
    End Sub

    ''' <summary>
    ''' 取得用戶輸入的資料
    ''' </summary>
    Private Sub FulfilledInputs(Result As Record)
        TxtName.Text = Result.StudentName
        TxtInputTest.Text = Result.TestMarks.ToString()
        TxtInputQuizzes.Text = Result.QuizzesMarks.ToString()
        TxtInputProject.Text = Result.ProjectMarks.ToString()
        TxtInputExam.Text = Result.ExamMarks.ToString()
    End Sub

    ''' <summary>
    ''' 顯示分數計算結果
    ''' </summary>
    ''' <param name="Result">表示分數的記錄</param>
    Private Sub ShowResult(Result As Record)
        If Result.IsReal Then
            TxtResultCA.Text = Result.CAMarks.ToString()
            TxtResultModule.Text = Result.ModuleMarks.ToString()
        Else
            TxtResultCA.Text = NaN
            TxtResultModule.Text = NaN
        End If
        TxtReusltGrade.Text = Result.ModuleGrade
        TxtReusltRemarks.Text = Result.Remarks
    End Sub

    ''' <summary>
    ''' 顯示統計數據
    ''' </summary>
    Private Sub ShowStatistics()
        Dim Counter As (A As Integer, B As Integer, C As Integer, F As Integer)
        Dim Scores As New List(Of Double)
        Dim Su As Double = 0
        For Each Record In RealData
            Select Case Record.ModuleGrade
                Case "A"
                    Counter.A += 1
                Case "B"
                    Counter.B += 1
                Case "C"
                    Counter.C += 1
                Case "F"
                    Counter.F += 1
            End Select
            Scores.Add(Record.ModuleMarks)
            Su += Record.ModuleMarks
        Next
        TxtStatisticsNo.Text = Scores.Count.ToString()
        TxtStatisticsA.Text = Counter.A.ToString()
        TxtStatisticsB.Text = Counter.B.ToString()
        TxtStatisticsC.Text = Counter.C.ToString()
        TxtStatisticsF.Text = Counter.F.ToString()
        If Scores.Count > 0 Then
            Scores.Sort()
            Dim Av As Double = Su / Scores.Count
            TxtStatisticsAv.Text = Av.ToString()
            Dim Sd As Double = 0
            For Each Da As Double In Scores
                Dim Di As Double = Da - Av
                Sd += Di * Di
            Next
            Sd = Math.Sqrt(Sd / Scores.Count)
            TxtStatisticsSd.Text = Sd.ToString()
            Dim Ha As Integer = Scores.Count >> 1
            Dim Md As Double = Scores(Ha)
            If Scores.Count = Ha << 1 Then
                Md += Scores(Ha - 1)
                Md /= 2
            End If
            TxtStatisticsMd.Text = Md.ToString()
        Else
            TxtStatisticsAv.Text = NaN
            TxtStatisticsSd.Text = NaN
            TxtStatisticsMd.Text = NaN
        End If
    End Sub

    ''' <summary>
    ''' 在數據列表當中搜尋符合條件的 Record
    ''' </summary>
    ''' <param name="GetSelectedIndex">指定當搜尋程序結束的時候， LstRecords 會選取哪一個項目</param>
    Private Sub RecordsSearch(GetSelectedIndex As GetSelectedIndex)
        LstRecords.Items.Clear()
        LstRecords.Items.Add(" -> [Input]")
        SearchIndexList = New List(Of Integer)()
        Dim Thrown As Exception = Nothing
        For i As Integer = 0 To Data.Count - 1
            Dim IsMatched As Boolean = False
            If ChkRecordsSearch.Checked Then
                Try
                    IsMatched = Regex.IsMatch(Data(i).StudentName, TxtRecordsSearch.Text)
                Catch Exception As Exception
                    Thrown = Exception
                End Try
            Else
                IsMatched = Data(i).StudentName.Contains(TxtRecordsSearch.Text)
            End If
            If IsMatched Then
                LstRecords.Items.Add(Data(i).StudentName + If(Not Data(i).IsReal, " (Not in the criteria)", String.Empty))
                SearchIndexList.Add(i)
            End If
        Next
        LstRecords.SelectedIndex = GetSelectedIndex.Invoke(Thrown)
    End Sub

    ''' <summary>
    ''' 檢查數據列表是否已經排序
    ''' </summary>
    Private Function RecordsSorted() As Boolean
        For i As Integer = 0 To Data.Count - 2
            If Data(i).CompareTo(Data(i + 1)) > 0 Then
                Return False
            End If
        Next
        Return True
    End Function

    ''' <summary>
    ''' 檢查多筆紀錄是否具有獨特的 ID
    ''' </summary>
    ''' <param name="Records">實現了 IReliability 的多筆紀錄列舉器</param>
    Private Shared Function HaveUniqueIDs(Records As IEnumerable(Of IReliability)) As Boolean
        Dim Names As New HashSet(Of Integer)()
        For Each Record As Record In Records
            Names.Add(Record.ID)
        Next
        Return Names.Count = Records.Count()
    End Function

    ''' <summary>
    ''' 檢查多筆紀錄是否具有獨特的 StudentName
    ''' </summary>
    ''' <param name="Records">實現了 IRecord 的多筆紀錄列舉器</param>
    Private Shared Function HaveUniqueNames(Records As IEnumerable(Of IRecord)) As Boolean
        Dim Names As New HashSet(Of String)()
        For Each Record As Record In Records
            Names.Add(Record.StudentName)
        Next
        Return Names.Count = Records.Count()
    End Function

    ''' <summary>
    ''' 透過訊息視窗顯示異常對象，其中包括整個內部異常的鏈條
    ''' </summary>
    ''' <param name="Exception">未處理的異常對象</param>
    Private Async Sub ShowException(Exception As Exception)
        While Tag = IsResizing.Yes
            Await SuspendNow()
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
    End Sub

    ''' <summary>
    ''' 透過訊息視窗顯示執行信息
    ''' </summary>
    ''' <param name="owner"></param>
    ''' <param name="text"></param>
    ''' <param name="caption"></param>
    ''' <param name="buttons"></param>
    ''' <param name="icon"></param>
    Private Async Function ShowMessage(owner As IWin32Window, text As String, caption As String, buttons As MessageBoxButtons, icon As MessageBoxIcon) As Task(Of DialogResult)
        While Tag = IsResizing.Yes
            Await SuspendNow()
        End While
        Return MessageBox.Show(owner, text, caption, buttons, icon)
    End Function

    ''' <summary>
    ''' 中斷一下
    ''' </summary>
    Private Function SuspendNow() As Task
        Return Task.Delay(DelayDuration)
    End Function

    ''' <summary>
    ''' 檢查連線狀態是否被強制中斷
    ''' </summary>
    ''' <param name="Capture"></param>
    Private Async Sub SocketState(Capture As Exception)
        While Capture IsNot Nothing
            If TypeOf Capture Is SocketException Then
                Const WSAECONNRESET As Integer = 10054
                If CType(Capture, SocketException).ErrorCode = WSAECONNRESET Then
                    While DataControlsLock
                        Await SuspendNow()
                    End While
                    BtnDataSourceConnect.PerformClick()
                End If
                Exit While
            End If
            Capture = Capture.InnerException
        End While
    End Sub

    ''' <summary>
    ''' 讀取本地數據緩存
    ''' </summary>
    Private Async Function ReadDataFile() As Task
        SuspendControls()
        Try
            DataFile = File.Open(FileName, FileMode.OpenOrCreate, FileAccess.ReadWrite, FileShare.Read)
            If DataFile.Length > 0 Then
                Dim JsonBuffer(DataFile.Length - 1) As Byte
                Await DataFile.ReadAsync(JsonBuffer, 0, DataFile.Length)
                Data.AddRange(JsonConvert.DeserializeObject(Of HashSet(Of Record))(Encoding.UTF8.GetString(JsonBuffer)))
                If Not HaveUniqueIDs(Data) Then
                    Dim Result As DialogResult = Await ShowMessage(Me, "Some of the records from local storage have the same ""ID"". Such cannot be inserted!", Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Dim Reserved As New HashSet(Of Integer)
                    Dim Duplicated As New HashSet(Of Integer)
                    For Each Record As Record In Data
                        If Reserved.Contains(Record.ID) Then
                            Duplicated.Add(Record.ID)
                        End If
                        Reserved.Add(Record.ID)
                    Next
                    For Each ID As Integer In Duplicated
                        Reserved.Remove(ID)
                    Next
                    Data = Data.Where(
                        Function(Record As Record) As Boolean
                            Return Reserved.Contains(Record.ID)
                        End Function
                    ).ToList()
                End If
            End If
        Catch Exception As Exception
            ShowException(Exception)
            If DataFile IsNot Nothing Then
                DataFile.Close()
                DataFile = Nothing
            End If
        End Try
        ResumeControls()
    End Function

    ''' <summary>
    ''' 寫入本地數據緩存
    ''' </summary>
    Private Async Function WriteDataFile() As Task
        SuspendControls()
        Try
            If DataFile IsNot Nothing Then
                Dim JsonBuffer As Byte() = Encoding.UTF8.GetBytes(JsonConvert.SerializeObject(Data, Formatting.Indented))
                DataFile.SetLength(0)
                Await DataFile.WriteAsync(JsonBuffer, 0, JsonBuffer.Length)
                DataFile.Close()
                If Data.Count = 0 Then
                    File.Delete(FileName)
                End If
            End If
        Catch Exception As Exception
            ShowException(Exception)
        End Try
        ResumeControls()
    End Function

    ''' <summary>
    ''' 對於所有視窗按鈕中嘗試執行控制項的操作，如果在 Controls 找不到相應的控制項則會繼續重試，直到找到相應的 Tag 來執行操作
    ''' </summary>
    ''' <param name="Action">對於按鈕的控制項本身所執行的操作</param>
    Private Sub WindowButtonsRequest(Action As Action(Of Control))
        Dim MetroFormButtonTag As Type = GetType(MetroForm).GetNestedType("WindowButtons", BindingFlags.NonPublic)
        WindowButtonsRequest(Action, MetroFormButtonTag.GetEnumValues())
    End Sub

    ''' <summary>
    ''' 在被選取的多個視窗按鈕中嘗試執行對於控制項的操作，如果在 Controls 找不到相應的控制項則會繼續重試，直到找到相應的 Tag 來執行操作
    ''' </summary>
    ''' <param name="Action">對於按鈕的控制項本身所執行的操作</param>
    ''' <param name="Tags">要對於某些給定的視窗按鈕執行操作（在關閉、最大化或者還原、最小化當中選取）</param>
    Private Sub WindowButtonsRequest(Action As Action(Of Control), ParamArray Tags As Object())
        WindowButtonsRequest(Action, CType(Tags, Array))
    End Sub

    ''' <summary>
    ''' 在被選取的多個視窗按鈕中嘗試執行對於控制項的操作，如果在 Controls 找不到相應的控制項則會繼續重試，直到找到相應的 Tag 來執行操作
    ''' </summary>
    ''' <param name="Action">對於按鈕的控制項本身所執行的操作</param>
    ''' <param name="Tags">要對於某些給定的視窗按鈕執行操作（在關閉、最大化或者還原、最小化當中選取）</param>
    Private Async Sub WindowButtonsRequest(Action As Action(Of Control), Tags As Array)
        Dim MetroFormButtonType As Type = GetType(MetroForm).GetNestedType("MetroFormButton", BindingFlags.NonPublic)
        Dim Exists(Tags.Length - 1) As Boolean
        Dim Takens As Integer = 0
        While Takens < Exists.Length
            For Each Control As Control In Controls
                For i As Integer = 0 To Tags.Length - 1
                    If Exists(i) Then
                    ElseIf Control.GetType() <> MetroFormButtonType Then
                    ElseIf Control.Tag.GetType() <> Tags(i).GetType() Then
                    ElseIf Control.Tag = Tags(i) Then
                        Action.Invoke(Control)
                        Exists(i) = True
                        Takens += 1
                        Exit For
                    End If
                Next
            Next
            Await SuspendNow()
        End While
    End Sub

    ''' <summary>
    ''' 產生隨機的錯誤代碼
    ''' </summary>
    Private Sub GenerateErrorCodes()
        ErrorCodes = New Dictionary(Of ErrorKeys, Integer)()
        Dim Temp As New HashSet(Of Integer)
        While Temp.Count < GetType(ErrorKeys).GetEnumValues().Length
            Temp.Add(RandomNumberGenerator.Next(RandomNumberMinimum, RandomNumberMaximum))
        End While
        Dim Index As Integer = 0
        For Each Key As ErrorKeys In GetType(ErrorKeys).GetEnumValues()
            ErrorCodes.Add(Key, Temp(Index))
            Index += 1
        Next
    End Sub

    ''' <summary>
    ''' 把 Sql 指令片段轉換成能夠被調用的臨時儲存進程
    ''' </summary>
    ''' <param name="Cmd">Sql 指令片段</param>
    Private Shared Function MakeSubroutine(Cmd As String) As (BeginCmd As String, EndCmd As String)
        Dim Co As String = Guid.NewGuid().ToString()
        Dim Nl As String = Environment.NewLine
        Dim BeginBuilder As New StringBuilder(InitialStringCapacity)
        BeginBuilder.Append("CREATE DATABASE `MARKS_CALCULATOR_SCHEMA{").Append(Co).Append("}`; ").Append(Nl)
        BeginBuilder.Append("DELIMITER $ ").Append(Nl)
        BeginBuilder.Append("CREATE PROCEDURE `MARKS_CALCULATOR_SCHEMA{").Append(Co).Append("}`.`MODULE_GRADE_SUBROUTINE` () BEGIN ").Append(Nl)
        BeginBuilder.Append(Cmd)
        BeginBuilder.Append("END$ ").Append(Nl)
        BeginBuilder.Append("DELIMITER ; ").Append(Nl)
        Dim EndBuilder As New StringBuilder(InitialStringCapacity)
        EndBuilder.Append("CALL `MARKS_CALCULATOR_SCHEMA{").Append(Co).Append("}`.`MODULE_GRADE_SUBROUTINE` (); ").Append(Nl)
        EndBuilder.Append("DROP DATABASE `MARKS_CALCULATOR_SCHEMA{").Append(Co).Append("}`; ").Append(Nl)
        Return (BeginBuilder.ToString(), EndBuilder.ToString())
    End Function

    ''' <summary>
    ''' Sql 指令片段的執行器
    ''' </summary>
    ''' <param name="Cmd">Sql 指令片段</param>
    Private Async Function ExecuteQuery(Cmd As String) As Task
        If DataSourceConnection.ServerVersion.Contains("MariaDB") Then
            DataReader = Await New MySqlCommand(Cmd, DataSourceConnection).ExecuteReaderAsync()
        Else
            Dim Subroutine As (BeginCmd As String, EndCmd As String) = MakeSubroutine(Cmd)
            Await New MySqlScript(DataSourceConnection, Subroutine.BeginCmd).ExecuteAsync()
            DataReader = Await New MySqlCommand(Subroutine.EndCmd, DataSourceConnection).ExecuteReaderAsync()
        End If
    End Function

    ''' <summary>
    ''' 上傳 LstRecords 中的紀錄到 MySql 資料庫
    ''' </summary>
    Private Async Function Upload() As Task
        Try
            Await ExecuteQuery(UploadCmd)
            If Not Await DataReader.ReadAsync() Then
                DataReader.Close()
                Return
            ElseIf DataReader.VisibleFieldCount = 1 AndAlso DataReader.GetName(0) = "ERROR_CODE" AndAlso DataReader.GetDataTypeName(0) = "INT" Then
                Select Case DataReader("ERROR_CODE")
                    Case ErrorCodes(ErrorKeys.InvalidSourceAndTable)
                        Await ShowMessage(Me, "The primary key should be specified as ""ID"" whenever exists.", Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Case ErrorCodes(ErrorKeys.NonExistFieldWithType)
                        Await ShowMessage(Me, "Some of the fields and data type thereof are not matching!", Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Case ErrorCodes(ErrorKeys.NoOperationState)
                        Await ShowMessage(Me, "There are nothing in local records!", Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Select
                DataReader.Close()
                Return
            End If
            DataReader.Close()
        Catch Exception As Exception
            ShowException(Exception)
            If DataReader IsNot Nothing AndAlso DataReader.IsClosed = False Then
                DataReader.Close()
            End If
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 下載 LstRecords 中的紀錄到 MySql 資料庫
    ''' </summary>
    Private Async Function Download() As Task
        Try
            Await ExecuteQuery(DownloadCmd)
            If Not Await DataReader.ReadAsync() Then
                DataReader.Close()
                Return
            ElseIf DataReader.VisibleFieldCount = 1 AndAlso DataReader.GetName(0) = "ERROR_CODE" AndAlso DataReader.GetDataTypeName(0) = "INT" Then
                Select Case DataReader("ERROR_CODE")
                    Case ErrorCodes(ErrorKeys.InvalidSourceAndTable)
                        Await ShowMessage(Me, "The data source or table thereof are missing!", Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Case ErrorCodes(ErrorKeys.NonExistFieldWithType)
                        Await ShowMessage(Me, "Some of the fields may not exist and data type thereof may not matching!", Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Case ErrorCodes(ErrorKeys.NoOperationState)
                        Await ShowMessage(Me, "There are nothing in source records!", Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
                End Select
                DataReader.Close()
                Return
            End If
            Dim SourceDownload As New List(Of Record)()
            Dim Keys As New HashSet(Of Integer)()
            Do
                Dim Temp As New Record()
                GetType(Record).GetProperty("StudentName").SetValue(Temp, DataReader("StudentName"))
                GetType(Record).GetProperty("TestMarks").SetValue(Temp, DataReader("Test"))
                GetType(Record).GetProperty("QuizzesMarks").SetValue(Temp, DataReader("Quizzes"))
                GetType(Record).GetProperty("ProjectMarks").SetValue(Temp, DataReader("Project"))
                GetType(Record).GetProperty("ExamMarks").SetValue(Temp, DataReader("Exam"))
                GetType(Record).GetProperty("ID").SetValue(Temp, DataReader("ID"))
                SourceDownload.Add(Temp)
                Keys.Add(Temp.ID)
            Loop While Await DataReader.ReadAsync()
            DataReader.Close()
            If Keys.Count < SourceDownload.Count Then
                Dim Result As DialogResult = Await ShowMessage(Me, "Some of the records from data source have the same ""ID"". Would you like to continue with inserting the rest?", Text, MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)
                If Result = DialogResult.Cancel Then
                    Return
                End If
                For Each ID As Integer In Keys
                    Dim Ambiguous As List(Of Record) = SourceDownload.Where(
                        Function(Source As Record) As Boolean
                            Return Source.ID = ID
                        End Function
                    ).ToList()
                    If Ambiguous.Count > 1 Then
                        For Each Source As Record In Ambiguous
                            SourceDownload.Remove(Source)
                        Next
                    End If
                Next
            End If
            Dim LocalReplacement As New List(Of (Local As Record, Source As Record))()
            For Each Source As Record In SourceDownload.ToList()
                For Each Local As Record In Data
                    If Source = Local Then
                        SourceDownload.Remove(Source)
                        Exit For
                    ElseIf Source.ID = Local.ID Then
                        LocalReplacement.Add((Local, Source))
                        SourceDownload.Remove(Source)
                        Exit For
                    End If
                Next
            Next
            If LocalReplacement.Count = 0 Then
                Data.AddRange(SourceDownload)
            Else
                Dim Result As DialogResult = Await ShowMessage(Me, "Some of the records from data source have the same ""ID"" to a local record but not matching the same fields. Would you like to replace it with the new one?", Text, MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning)
                If Result = DialogResult.Yes Then
                    For Each Replacement In LocalReplacement
                        Data(Data.IndexOf(Replacement.Local)) = Replacement.Source
                    Next
                    Data.AddRange(SourceDownload)
                ElseIf Result = DialogResult.No Then
                    Data.AddRange(SourceDownload)
                End If
            End If
        Catch Exception As Exception
            ShowException(Exception)
            If DataReader IsNot Nothing AndAlso DataReader.IsClosed = False Then
                DataReader.Close()
            End If
            Throw
        End Try
    End Function

    ''' <summary>
    ''' 繪畫進度軌道
    ''' </summary>
    Private Sub DrawProgressTrack(Paint As Graphics)
        Dim Boundary As New Rectangle(New Point(0, 0), New Size(ClientSize.Width, ProgressBarHeight))
        Dim Brusher As New SolidBrush(MetroPaint.GetStyleColor(Style))
        Paint.FillRectangle(Brusher, Boundary)
    End Sub

    ''' <summary>
    ''' 繪畫進度條
    ''' </summary>
    Private Sub DrawProgressBar(Paint As Graphics)
        ProgressOffset = ProgressOffset Mod (ClientSize.Width + ProgressBarWidth)
        Dim Boundary As New Rectangle(New Point(ProgressOffset - ProgressBarWidth, 0), New Size(ProgressBarWidth, ProgressBarHeight))
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
        Await ReadDataFile()
        ShowStatistics()
        RecordsSearch(
            Function(Exception As Exception) As Integer
                Return 0
            End Function
        )
        If Not HaveUniqueNames(Data) Then
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
                Throw New BranchesShouldNotBeInstantiatedException("Status has been changed!")
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
                    Await SuspendNow()
                End While
                BtnDataSourceConnect.PerformClick()
                While Connection <> ConnectState.Disconnected
                    Await SuspendNow()
                End While
            End If
            Await WriteDataFile()
            State = FormState.Finalizing
            Close()
        ElseIf State = FormState.CloseHasStarted Then
            e.Cancel = True
        End If
    End Sub

    Private Sub TxtInput_Enter(sender As Object, e As EventArgs) Handles TxtInputTest.Enter, TxtInputQuizzes.Enter, TxtInputProject.Enter, TxtInputExam.Enter
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New BranchesShouldNotBeInstantiatedException("Type not matching!")
        End If
        If LstRecords.Tag = IsAdding.Yes Then
            CType(sender, MetroTextBox).Tag = IsTyping.Yes
        End If
    End Sub

    Private Sub TxtInput_TextChanged(sender As Object, e As EventArgs) Handles TxtInputTest.TextChanged, TxtInputQuizzes.TextChanged, TxtInputProject.TextChanged, TxtInputExam.TextChanged
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New BranchesShouldNotBeInstantiatedException("Type not matching!")
        End If
        If LstRecords.Tag = IsAdding.Yes AndAlso CType(sender, MetroTextBox).Tag = IsTyping.Yes Then
            Dim Number As Double = 0
            ShowResult(If(Double.TryParse(CType(sender, MetroTextBox).Text, Number), InputedRecord, Record.Null))
        End If
    End Sub

    Private Sub TxtInput_Leave(sender As Object, e As EventArgs) Handles TxtInputTest.Leave, TxtInputQuizzes.Leave, TxtInputProject.Leave, TxtInputExam.Leave
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New BranchesShouldNotBeInstantiatedException("Type not matching!")
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
            ShowResult(Temp)
        End If
    End Sub

    Private Sub TxtNameWithInput_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtName.KeyDown, TxtInputTest.KeyDown, TxtInputQuizzes.KeyDown, TxtInputProject.KeyDown, TxtInputExam.KeyDown
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New BranchesShouldNotBeInstantiatedException("Type not matching!")
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

    Private Sub TxtSource_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtDataSourceDatabase.KeyDown, TxtDataSourceTable.KeyDown
        If sender Is Nothing OrElse TypeOf sender IsNot MetroTextBox Then
            Throw New BranchesShouldNotBeInstantiatedException("Type not matching!")
        End If
        If e.KeyCode = Keys.Enter Then '（實現輸入資料的快速跳轉 Enter 按鍵）
            If Not CType(sender, MetroTextBox).Equals(TxtDataSourceTable) Then
                SelectNextControl(ActiveControl, True, True, True, True)
            Else
                SelectNextControl(TxtName, True, True, True, True)
            End If
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TxtRecordsSearch_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtRecordsSearch.KeyDown
        If e.KeyCode = Keys.Enter Then '（實現匹配正則表達式的快速切換 Enter 按鍵）
            ChkRecordsSearch.Checked = Not ChkRecordsSearch.Checked
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Async Sub BtnDataSourceConnect_Click(sender As Object, e As EventArgs) Handles BtnDataSourceConnect.Click
        If Connection = ConnectState.Disconnected Then
            If FrmConnection.ShowDialog(Me) = DialogResult.Cancel Then
                Return
            End If
            Try
                DataSourceConnection = New MySqlConnection(ConnectionCmd)
                Connection = ConnectState.Connecting
                Await DataSourceConnection.OpenAsync()
                Connection = ConnectState.Connected
            Catch Exception As Exception
                ShowException(Exception)
                Connection = ConnectState.Disconnected
            End Try
        ElseIf Connection = ConnectState.Connected Then
            Connection = ConnectState.Disconnecting
            Await DataSourceConnection.CloseAsync()
            Connection = ConnectState.Disconnected
        End If
    End Sub

    Private Async Sub BtnDataSourceUpload_Click(sender As Object, e As EventArgs) Handles BtnDataSourceUpload.Click
        ConnectLock = True
        Try
            GenerateErrorCodes()
            Await Upload()
        Catch Exception As Exception
            SocketState(Exception)
        End Try
        ConnectLock = False
    End Sub

    Private Async Sub BtnDataSourceDownload_Click(sender As Object, e As EventArgs) Handles BtnDataSourceDownload.Click
        ConnectLock = True
        Try
            GenerateErrorCodes()
            Await Download()
        Catch Exception As Exception
            SocketState(Exception)
        End Try
        Dim CaptureIndex As Integer = LstRecords.SelectedIndex
        ShowStatistics()
        RecordsSearch(
            Function(Exception As Exception) As Integer
                Return CaptureIndex
            End Function
        )
        ConnectLock = False
    End Sub

    Private Sub BtnRecordsAdd_Click(sender As Object, e As EventArgs) Handles BtnRecordsAdd.Click
        If Temp.StudentName = String.Empty Then
            MessageBox.Show(Me, "The record to be inserted into the local records should be a record that has a non-empty ""StudentName"".", Text, MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return
        End If
        If Not ChkRecords.Checked Then
            For Each Record As Record In Data
                If Record.StudentName = Temp.StudentName Then
                    MessageBox.Show(Me, "The record to be inserted into the local records should not match the same ""StudentName"". If you would like to suppress the restriction, you have to tick out the ""Allow Duplicated Names"" checkbox.", Text, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Return
                End If
            Next
        End If
        Data.Add(Temp)
        Temp = New Record()
        ShowStatistics()
        RecordsSearch(
            Function(Exception As Exception) As Integer
                Return 0
            End Function
        )
    End Sub

    Private Sub BtnRecordsRemove_Click(sender As Object, e As EventArgs) Handles BtnRecordsRemove.Click
        Data.Remove(SelectedRecord)
        ShowStatistics()
        Dim CaptureIndex As Integer = LstRecords.SelectedIndex
        RecordsSearch(
            Function(Exception As Exception) As Integer
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
            Function(Exception As Exception) As Integer
                Return CaptureIndex
            End Function
        )
    End Sub

    Private Sub BtnRecordsSquare_Click(sender As Object, e As EventArgs) Handles BtnRecordsSquare.Click
        Data.Sort()
        Dim CaptureIndex As Integer = LstRecords.SelectedIndex
        RecordsSearch(
            Function(Exception As Exception) As Integer
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
            Function(Exception As Exception) As Integer
                Return CaptureIndex
            End Function
        )
    End Sub

    Private Sub Btn_Enter(sender As Object, e As EventArgs) Handles BtnDataSourceConnect.Enter, BtnDataSourceUpload.Enter, BtnDataSourceDownload.Enter, BtnRecordsAdd.Enter, BtnRecordsRemove.Enter, BtnRecordsUp.Enter, BtnRecordsSquare.Enter, BtnRecordsDown.Enter
        If sender Is Nothing OrElse TypeOf sender IsNot MetroButton Then
            Throw New BranchesShouldNotBeInstantiatedException("Type not matching!")
        End If
        CType(sender, MetroButton).Highlight = True
    End Sub

    Private Sub Btn_Leave(sender As Object, e As EventArgs) Handles BtnDataSourceConnect.Leave, BtnDataSourceUpload.Leave, BtnDataSourceDownload.Leave, BtnRecordsAdd.Leave, BtnRecordsRemove.Leave, BtnRecordsUp.Leave, BtnRecordsSquare.Leave, BtnRecordsDown.Leave
        If sender Is Nothing OrElse TypeOf sender IsNot MetroButton Then
            Throw New BranchesShouldNotBeInstantiatedException("Type not matching!")
        End If
        CType(sender, MetroButton).Highlight = False
    End Sub

    Private Sub ChkRecords_Enter(sender As Object, e As EventArgs) Handles ChkRecords.Enter, ChkRecordsSearch.Enter
        If sender Is Nothing OrElse TypeOf sender IsNot MetroCheckBox Then
            Throw New BranchesShouldNotBeInstantiatedException("Type not matching!")
        End If
        CType(sender, MetroCheckBox).CustomBackground = True
        CType(sender, MetroCheckBox).BackColor = SystemColors.ControlLight
    End Sub

    Private Sub ChkRecords_Leave(sender As Object, e As EventArgs) Handles ChkRecords.Leave, ChkRecordsSearch.Leave
        If sender Is Nothing OrElse TypeOf sender IsNot MetroCheckBox Then
            Throw New BranchesShouldNotBeInstantiatedException("Type not matching!")
        End If
        CType(sender, MetroCheckBox).CustomBackground = False
    End Sub

    Private Sub ChkRecords_CheckedChanged(sender As Object, e As EventArgs) Handles ChkRecords.CheckedChanged
        If Not ChkRecords.Checked AndAlso Not HaveUniqueNames(Data) Then
            MessageBox.Show(Me, "Some of the local records have the same ""StudentName""! Whenever a record is inserted into the local records, it is still avoided all these records match the same ""StudentName"".", Text, MessageBoxButtons.OK, MessageBoxIcon.Information)
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
                BtnRecordsSquare.Enabled = Not RecordsSorted()
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
        FulfilledInputs(SelectedRecord)
        ShowResult(SelectedRecord)
    End Sub

    Private Sub AnyRecordsSearch_Event(sender As Object, e As EventArgs) Handles TxtRecordsSearch.TextChanged, ChkRecordsSearch.CheckedChanged
        Dim CaptureIndex As Integer = LstRecords.SelectedIndex
        RecordsSearch(
            Function(Exception As Exception) As Integer
                Return If(CaptureIndex < LstRecords.Items.Count, CaptureIndex, LstRecords.Items.Count - 1)
            End Function
        )
    End Sub

    Private Sub ChkEnterKeys_Event(sender As Object, e As KeyEventArgs) Handles ChkRecords.KeyDown, ChkRecordsSearch.KeyDown
        If sender Is Nothing OrElse TypeOf sender IsNot MetroCheckBox Then
            Throw New BranchesShouldNotBeInstantiatedException("Type not matching!")
        End If
        If e.KeyCode = Keys.Enter Then '（實現透過鍵盤 Enter 按鍵改變核對方塊的狀態）
            CType(sender, MetroCheckBox).Checked = Not CType(sender, MetroCheckBox).Checked
        End If
    End Sub

    Private Sub TmrMain_Tick(sender As Object, e As EventArgs) Handles TmrMain.Tick
        If DrawingProgress AndAlso Tag = IsResizing.No Then
            Dim Graphics As Graphics = CreateGraphics() '（透過繪製進度條，保持視窗的標題列位置能夠捕獲相認的訊息）
            DrawProgressTrack(Graphics)
            DrawProgressBar(Graphics)
            Graphics.Dispose()
            ProgressOffset += ProgressBarDeltaX
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
                MetroFormButtonTag.GetEnumValues()(1)
            )
            '（修復對於在視窗空白位置雙擊從而改變視窗狀態時，最大化或一般按鈕樣式無法改變樣式的問題）
            If WindowState = FormWindowState.Normal Then
                Owner.Show() '（對於視窗由最大化即 Form.WindowState 為 FormWindowState.Maximized 變為一般即 Form.WindowState 為 FormWindowState.Normal 會失去分層視窗之底層陰影的修復）
                Activate() '（對於視窗由最大化即 Form.WindowState 為 FormWindowState.Maximized 變為一般即 Form.WindowState 為 FormWindowState.Normal 會失去焦點的修復）
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

    ''' <summary>
    ''' 委託調用決定 LstRecords.SelectedIndex
    ''' </summary>
    ''' <param name="Thrown">正則表達式匹配的異常對象</param>
    ''' <returns></returns>
    Private Delegate Function GetSelectedIndex(Thrown As Exception) As Integer

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

        Public Const TestScale As Double = 0.5
        Public Const QuizzesScale As Double = 0.2
        Public Const ProjectScale As Double = 0.3
        Public Const CAScale As Double = 0.4
        Public Const ExamScale As Double = 0.6
        Private Const Invalid As String = "[Invalid]"

#End Region

#Region "Fields"

        Private Name As String
        Private Test As Double
        Private Quizzes As Double
        Private Project As Double
        Private Exam As Double
        Private Code As Integer
        Public Shared ReadOnly Null As New Record(Invalid, Double.NaN, Double.NaN, Double.NaN, Double.NaN)

#End Region

#Region "Constructors"

        <JsonConstructor>
        Public Sub New()
            Name = String.Empty
            Test = 0
            Quizzes = 0
            Project = 0
            Exam = 0
            Code = 0
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
                If Not IsReal Then
                    Return Invalid
                ElseIf ModuleGrade <> "F" Then
                    Return "Pass"
                ElseIf ModuleMarks >= 30 Then
                    Return "Resit Exam"
                Else
                    Return "Restudy"
                End If
            End Get
        End Property

        <JsonProperty>
        Public Property StudentName As String Implements IRecord.StudentName
            Get
                Return Name
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

        Private Shared Function Enumerate(Separator As String, ParamArray Params() As String) As String
            If Params Is Nothing OrElse Params.Length = 0 Then
                Return String.Empty
            End If
            Dim Result As String = Params(0)
            For Each Param As String In Params.Skip(1)
                Result += Separator + Param
            Next
            Return Result
        End Function

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

        Public Shared Widening Operator CType(Value As (Name As String, Test As Double, Quizzes As Double, Project As Double, Exam As Double)) As Record
            Return New Record(Value.Name, Value.Test, Value.Quizzes, Value.Project, Value.Exam)
        End Operator

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

<Serializable>
Friend Class BranchesShouldNotBeInstantiatedException
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
