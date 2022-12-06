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
    Private ReadOnly Getter As Func(Of (Host As String, Username As String, Password As String))

    ''' <summary>
    ''' 表示數據的寫入運算子
    ''' </summary>
    Private ReadOnly Setter As Action(Of (Host As String, Username As String, Password As String))

#End Region

#Region "Constructors"

    Public Sub New(Getter As Func(Of (Host As String, Username As String, Password As String)), Setter As Action(Of (Host As String, Username As String, Password As String)))
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
        Dim Data As (Host As String, Username As String, Password As String) = Getter?.Invoke()
        TxtHost.Text = Data.Host
        TxtUsername.Text = Data.Username
        TxtPassword.Text = Data.Password
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
        Setter?.Invoke((TxtHost.Text, TxtUsername.Text, TxtPassword.Text))
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

End Class
