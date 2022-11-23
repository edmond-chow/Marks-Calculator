' Name:         Marks Calculator Project
' Purpose:      Used to calculate student grades.
' Programmer:   Chow Cheuk Hei, Tse Ka Yu on 2 / 10 / 2022

Public Class FrmConnect

#Region "Fields"

    Private ReadOnly Action As Action(Of (Host As String, Username As String, Password As String))

#End Region

#Region "Constructors"

    Public Sub New(Action As Action(Of (Host As String, Username As String, Password As String)))
        ' 設計工具需要此呼叫。
        InitializeComponent()
        ' 在 InitializeComponent() 呼叫之後加入所有初始設定。
        MinimumSize = Size
        Me.Action = Action
    End Sub

#End Region

#Region "Handles"

    Private Sub BtnOK_Click(sender As Object, e As EventArgs) Handles BtnOK.Click
        DialogResult = DialogResult.OK
        Action?.Invoke((TxtHost.Text, TxtUsername.Text, TxtPassword.Text))
    End Sub

    Private Sub BtnCancel_Click(sender As Object, e As EventArgs) Handles BtnCancel.Click
        DialogResult = DialogResult.Cancel
    End Sub

    Private Sub FrmConnect_FormClosed(sender As Object, e As FormClosedEventArgs) Handles Me.FormClosed
        If DialogResult = Nothing Then
            DialogResult = DialogResult.Cancel
        End If
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
