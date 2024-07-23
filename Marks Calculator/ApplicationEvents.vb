'   Marks Calculator Project
'   Purpose: Used to calculate student grades.
'   Copyright © 2022, Edmond Chow and its contributors. All rights reserved.
'   
'   Marks Calculator is available under the Apache License 2.0, in which the version
'   8.0.31 of MySql.Data which is under GNU General Public License version 2 grants
'   you an additional permission to link the program and your derivative works with
'   the separately licensed software. For more information, please see
'   https://github.com/mysql/mysql-connector-net/blob/8.0.31/LICENSE.
'   
'   Licensed under the Apache License, Version 2.0 (the "License");
'   you may not use this file except in compliance with the License.
'   You may obtain a copy of the License at
'   
'   http://www.apache.org/licenses/LICENSE-2.0
'   
'   Unless required by applicable law or agreed to in writing, software
'   distributed under the License is distributed on an "AS IS" BASIS,
'   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
'   See the License for the specific language governing permissions and
'   limitations under the License.
'   
'   For more information, please contact edmond-chow@outlook.com.

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
