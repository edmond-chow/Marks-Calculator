' Name:         Marks Calculator Project
' Purpose:      Used to calculate student grades.
' Programmer:   Chow Cheuk Hei, Tse Ka Yu on 2 / 10 / 2022

Public Class FrmConnect

#Region "Constants"

    ''' <summary>
    ''' 用於表示密碼輸入文字
    ''' </summary>
    Const PasswordChar As String = "•"

#End Region

#Region "Fields"

    ''' <summary>
    ''' 表示數據的讀取運算子
    ''' </summary>
    Private ReadOnly Getter As GetterType

    ''' <summary>
    ''' 表示數據的寫入運算子
    ''' </summary>
    Private ReadOnly Setter As SetterType

#End Region

#Region "Constructors"

    Public Sub New(Getter As GetterType, Setter As SetterType)
        ' 設計工具需要此呼叫。
        InitializeComponent()
        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。
        MinimumSize = Size
        Me.Getter = Getter
        Me.Setter = Setter
    End Sub

#End Region

#Region "Enumerations"

    ''' <summary>
    ''' 表示控制項正在輸入數據
    ''' </summary>
    Private Enum IsTyping
        Yes = 1
        No = 2
    End Enum

#End Region

#Region "Handles"

    Private Sub FrmConnect_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim Info As (Server As String, Port As String, Username As String, Password As String, Schema As String) = Getter?.Invoke()
        TxtServer.Text = Info.Server
        TxtPort.Text = Info.Port
        TxtUsername.Text = Info.Username
        TxtPassword.Text = Info.Password
        TxtSchema.Text = Info.Schema
        TxtPassword.PasswordChar = PasswordChar
    End Sub

    Private Sub BtnOK_Click(sender As Object, e As EventArgs) Handles BtnOK.Click
        DialogResult = DialogResult.OK
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        DialogResult = DialogResult.Cancel
    End Sub

    Private Sub FrmConnect_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If DialogResult = Nothing Then
            DialogResult = DialogResult.Cancel
        End If
        Setter?.Invoke((TxtServer.Text, TxtPort.Text, TxtUsername.Text, TxtPassword.Text, TxtSchema.Text))
    End Sub

    Private Sub TxtPassword_Enter(sender As Object, e As EventArgs) Handles TxtPassword.Enter
        TxtPassword.Tag = IsTyping.Yes
    End Sub

    Private Sub TxtPassword_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtPassword.KeyDown
        If TxtPassword.Tag = IsTyping.Yes AndAlso e.KeyCode = Keys.ShiftKey Then
            TxtPassword.PasswordChar = String.Empty
        ElseIf e.Shift Then
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TxtPassword_KeyUp(sender As Object, e As KeyEventArgs) Handles TxtPassword.KeyUp
        If TxtPassword.Tag = IsTyping.Yes AndAlso e.KeyCode = Keys.ShiftKey Then
            TxtPassword.PasswordChar = PasswordChar
        End If
    End Sub

    Private Sub TxtPassword_Leave(sender As Object, e As EventArgs) Handles TxtPassword.Leave
        TxtPassword.Tag = IsTyping.No
        TxtPassword.PasswordChar = PasswordChar
    End Sub

    Private Sub TxtSource_KeyDown(sender As Object, e As KeyEventArgs) Handles TxtServer.KeyDown, TxtPort.KeyDown, TxtUsername.KeyDown, TxtPassword.KeyDown, TxtSchema.KeyDown
        If sender Is Nothing OrElse TypeOf sender IsNot TextBox Then
            Throw New BranchesShouldNotBeInstantiatedException("Type not matching!")
        End If
        If e.KeyCode = Keys.Enter Then '（實現輸入資料的快速跳轉 Enter 按鍵）
            SelectNextControl(ActiveControl, True, True, True, True)
            e.SuppressKeyPress = True
        End If
    End Sub

    Private Sub TxtSource_Leave(sender As Object, e As EventArgs) Handles TxtServer.Leave, TxtPort.Leave, TxtUsername.Leave, TxtPassword.Leave, TxtSchema.Leave
        If sender Is Nothing OrElse TypeOf sender IsNot TextBox Then
            Throw New BranchesShouldNotBeInstantiatedException("Type not matching!")
        End If
        CType(sender, TextBox).Text = CType(sender, TextBox).Text.Replace("'", "") '（移除無效字元）
    End Sub

    Protected Overrides Sub WndProc(ByRef m As Message)
        Const WM_PAINT As Integer = &HF
        Select Case m.Msg
            Case WM_PAINT
                DefWndProc(m) '（對於這個表單的 Form.ShowDialog(owner As IWin32Window) As DialogResult 方法被調用時，採用預設的繪製方式，移除右下角由點陣組成的三角形）
            Case Else
                MyBase.WndProc(m)
        End Select
    End Sub

#End Region

#Region "Delegates"

    ''' <summary>
    ''' 表示數據的讀取運算子的類型
    ''' </summary>
    Public Delegate Function GetterType() As (Server As String, Port As String, Username As String, Password As String, Schema As String)

    ''' <summary>
    ''' 表示數據的寫入運算子的類型
    ''' </summary>
    Public Delegate Sub SetterType(Info As (Server As String, Port As String, Username As String, Password As String, Schema As String))

#End Region

End Class
