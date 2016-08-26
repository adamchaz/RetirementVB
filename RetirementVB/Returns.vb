Option Explicit On
Imports Microsoft.Office.Interop.Excel
Imports System.Math
Public Class Returns
    Private clsReturns As Dictionary(Of String, Double)
    Private clsVols As Dictionary(Of String, Double)
    Private clschoskey(,) As Double
    Private clsNumAssets As Long

    Public Sub Load(iReturns As Object, iCorrelation As Object)
        Dim i As Long
        Dim j As Long
        Dim lmat(,) As Double


        clsReturns = New Dictionary(Of String, Double)
        clsVols = New Dictionary(Of String, Double)
        clsNumAssets = UBound(iCorrelation, 1) - 1

        For i = 2 To UBound(iReturns, 1)
            clsReturns.Add(iReturns(i, 1), iReturns(i, 2))
            clsVols.Add(iReturns(i, 1), iReturns(i, 3))
        Next i
        ReDim lmat(0 To clsNumAssets, 0 To clsNumAssets)
        For i = 2 To UBound(iCorrelation, 1)
            For j = 2 To UBound(iCorrelation, 2)
                lmat(i - 1, j - 1) = iCorrelation(i, j)
            Next j
        Next i
        clschoskey = MatrixMath.MatrixCholesky(lmat)
    End Sub
    Public Function Returns() As Dictionary(Of String, Double)
        Dim lrnd(,) As Double
        Dim lKey As String
        Dim i As Long
        Dim lOut As Dictionary(Of String, Double)
        Try
            lOut = New Dictionary(Of String, Double)
            ReDim lrnd(0 To clsNumAssets, 0 To 1)
            For i = 1 To clsNumAssets

                lrnd(i, 1) = ((-2 * Log(Rand())) ^ 0.5) * Cos(2 * 3.141592654 * Rand())
            Next i
            lrnd = MatrixMath.MatrixMultiply(clschoskey, lrnd)
            i = 0
            For Each lKey In clsReturns.Keys
                lOut.Add(lKey, clsReturns(lKey) + lrnd(i + 1, 1) * clsVols(lKey))
                i = i + 1
            Next

            Returns = lOut
            Exit Function
        Catch e As Exception
            Throw
        End Try
    End Function


End Class
