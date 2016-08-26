Public Class AssetAllocation
    Private clsMinAge As Long
    Private clsMaxAge As Long
    Private ClsAssetAll As Object

    Public Function GetAssetAll(iage As Long) As Dictionary(Of String, Double)
        Dim i As Long
        Dim j As Long
        Dim lout As Dictionary(Of String, Double)

        If iage <= clsMinAge Then
            i = 2
        ElseIf iage >= clsMaxAge Then
            i = UBound(ClsAssetAll, 2)
        Else
            For i = 2 To UBound(ClsAssetAll, 2) - 1
                If iage >= ClsAssetAll(1, i) And iage < ClsAssetAll(1, i + 1) Then
                    Exit For
                End If
            Next i
        End If
        lout = New Dictionary(Of String, Double)
        For j = 2 To UBound(ClsAssetAll, 1)
            lout.Add(ClsAssetAll(j, 1), ClsAssetAll(j, i))
        Next j

        GetAssetAll = lout
    End Function

    Public Sub LoadMatrix(iAssetAll As Object)
        ClsAssetAll = iAssetAll
        clsMinAge = iAssetAll(1, 2)
        clsMaxAge = iAssetAll(1, UBound(iAssetAll, 2))
    End Sub
End Class
