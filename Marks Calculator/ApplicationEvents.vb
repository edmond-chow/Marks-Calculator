'   Marks Calculator Project
'   Purpose: Used to calculate student grades.
'   Copyright © 2022, Edmond Chow and its contributors, All rights reserved.
'
'   This program is free software; you can redistribute it and/or modify
'   it under the terms of the GNU General Public License as published by
'   the Free Software Foundation; either version 2 of the License, or
'   (at your option) any later version.
'
'   This program is distributed in the hope that it will be useful,
'   but WITHOUT ANY WARRANTY; without even the implied warranty of
'   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'   GNU General Public License for more details.
'
'   You should have received a copy of the GNU General Public License along
'   with this program; if not, write to the Free Software Foundation, Inc.,
'   51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.
'
'   For more information, please contact 69935538chow@gmail.com.

Imports Microsoft.VisualBasic.ApplicationServices

Namespace My

    ' MyApplication 可以使用下列事件:
    ' Startup:在應用程式啟動時，但尚未建立啟動表單之前引發。
    ' Shutdown:在所有應用程式表單關閉之後引發。如果應用程式不正常終止，就不會引發此事件。
    ' UnhandledException:在應用程式發生未處理的例外狀況時引發。
    ' StartupNextInstance:在啟動單一執行個體應用程式且應用程式已於使用中時引發。 
    ' NetworkAvailabilityChanged:在建立或中斷網路連線時引發。

    Partial Friend Class MyApplication

        Private Shared Sub ShowException(Exception As Exception)
            While Exception IsNot Nothing
                MessageBox.Show(Nothing,
                Exception.Message + Environment.NewLine +
                Environment.NewLine +
                Exception.TargetSite.ToString() + Environment.NewLine +
                Environment.NewLine +
                Exception.StackTrace.ToString() _
                , Exception.Source, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                Exception = Exception.InnerException
            End While
        End Sub

        Private Sub MyApplication_UnhandledException(sender As Object, e As UnhandledExceptionEventArgs) Handles Me.UnhandledException
            ShowException(e.Exception)
        End Sub

    End Class

End Namespace
