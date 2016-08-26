
Imports Microsoft.Office.Interop.Excel

Public Class Sheet1

    Private Sub Sheet1_Startup() Handles Me.Startup

    End Sub

    Private Sub Sheet1_Shutdown() Handles Me.Shutdown

    End Sub

    Private Sub Sheet1_BeforeDoubleClick(Target As Range, ByRef Cancel As Boolean) Handles Me.BeforeDoubleClick
        If Target.Address = "$D$1:$E$2" Then
            Call Main.Main()
        End If
    End Sub
End Class
