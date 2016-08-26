Option Explicit On
Public Class SimResults
    Private clsSimNum As Long
    Private clsMinSim As Long
    Private clsMaxSim As Long
    Private clsMedianSim As Long
    Private cls25Sim As Long
    Private cls75Sim As Long
    Private clsEndBalDict As SortedDictionary(Of Double, Long)
    Private clsSimDict As SortedDictionary(Of Long, SortedDictionary(Of Long, ReturnItem))

    Public Sub New()
        clsEndBalDict = New SortedDictionary(Of Double, Long)
        clsSimDict = New SortedDictionary(Of Long, SortedDictionary(Of Long, ReturnItem))
    End Sub

    Public Sub Add(iSimDict As SortedDictionary(Of Long, ReturnItem))
        Dim lAge As Long
        Dim lendBal As Double
        Dim lcount As Long

        clsSimNum = clsSimNum + 1
        clsSimDict.Add(clsSimNum, iSimDict)

        lcount = iSimDict.Count
        lAge = iSimDict.Keys(lcount - 1)
        lendBal = iSimDict(lAge).EndBalance
        If clsEndBalDict.ContainsKey(lendBal) Then
            lendBal = lendBal + 0.001
        End If
        clsEndBalDict.Add(lendBal, clsSimNum)

    End Sub
    Public Sub PrepareResults()
        Dim lkeys As SortedDictionary(Of Double, Long).ValueCollection

        lkeys = clsEndBalDict.Values
        clsMinSim = lkeys(0)
        clsMaxSim = lkeys(clsSimNum - 1)


        If clsSimNum = 1 Then
                clsMedianSim = lKeys(0)
            ElseIf clsSimNum = 2 Then
                clsMinSim = lKeys(0)
                clsMaxSim = lKeys(1)
            ElseIf clsSimNum = 3 Then
                clsMinSim = lKeys(0)
                clsMedianSim = lKeys(1)
                clsMaxSim = lKeys(2)
            ElseIf clsSimNum = 4 Then
                clsMinSim = lKeys(0)
                clsMedianSim = lKeys(1)
                clsMaxSim = lKeys(3)
            ElseIf clsSimNum = 5 Then
                clsMinSim = lKeys(0)
                cls25Sim = lKeys(1)
                clsMedianSim = lKeys(2)
                cls75Sim = lKeys(3)
                clsMaxSim = lKeys(4)
            Else
                clsMinSim = lKeys(0)
                clsMaxSim = lKeys(clsSimNum - 1)
                clsMedianSim = lKeys(CInt(clsSimNum * 0.5) - 1)
                cls25Sim = lKeys(CInt(clsSimNum * 0.25) - 1)
                cls75Sim = lKeys(CInt(clsSimNum * 0.75) - 1)
            End If
    End Sub
    Public Function GetValue(iSim As String, iAge As Long, iValue As String) As Double
        Dim lOut As Double
        Dim iSimNUM As Long
        Dim lReturnItem As ReturnItem

        Select Case UCase(iSim)
            Case "MIN"
                iSimNUM = clsMinSim
            Case "MAX"
                iSimNUM = clsMaxSim
            Case "MEDIAN"
                iSimNUM = clsMedianSim
            Case "25TH"
                iSimNUM = cls25Sim
            Case "75TH"
                iSimNUM = cls75Sim
        End Select

        lReturnItem = clsSimDict(iSimNUM)(iAge)

        Select Case UCase(iValue)
            Case "BEGBALANCE"
                lOut = lReturnItem.BegBalance
            Case "EMPLOYEECONT"
                lOut = lReturnItem.EmployeeCont
            Case "EMPLOYERCONT"
                lOut = lReturnItem.EmployerCont
            Case "ASSETRETURN"
                lOut = lReturnItem.AssetReturn
            Case "ENDBALANCE"
                lOut = lReturnItem.EndBalance
            Case "RETURNPCT"
                lOut = lReturnItem.ReturnPct
            Case "SALARY"
                lOut = lReturnItem.Salary
        End Select
        GetValue = lOut
    End Function
    Public Function GetAllValue(iAge As Long, iValue As String) As Double(,)
        Dim lOut(,) As Double
        Dim iSimNUM As Long
        Dim lReturnItem As ReturnItem

        lOut = New Double(clsSimNum - 1, 1) {}
        ReDim lOut(clsSimNum - 1, 1)

        For iSimNUM = 1 To clsSimNum
            lReturnItem = clsSimDict(iSimNUM)(iAge)
            lOut(iSimNUM - 1, 0) = iSimNUM
            Select Case UCase(iValue)
                Case "BEGBALANCE"
                    lOut(iSimNUM - 1, 1) = lReturnItem.BegBalance
                Case "EMPLOYEECONT"
                    lOut(iSimNUM - 1, 1) = lReturnItem.EmployeeCont
                Case "EMPLOYERCONT"
                    lOut(iSimNUM - 1, 1) = lReturnItem.EmployerCont
                Case "ASSETRETURN"
                    lOut(iSimNUM - 1, 1) = lReturnItem.AssetReturn
                Case "ENDBALANCE"
                    lOut(iSimNUM - 1, 1) = lReturnItem.EndBalance
                Case "RETURNPCT"
                    lOut(iSimNUM - 1, 1) = lReturnItem.ReturnPct
                Case "SALARY"
                    lOut(iSimNUM - 1, 1) = lReturnItem.Salary
            End Select

        Next iSimNUM
        GetAllValue = lOut
    End Function


    Private Function GetSimNum(iSim As String) As Long
        Dim iSimNUM As Long
        Select Case UCase(iSim)
            Case "MIN"
                iSimNUM = clsMinSim
            Case "MAX"
                iSimNUM = clsMaxSim
            Case "MEDIAN"
                iSimNUM = clsMedianSim
            Case "25TH"
                iSimNUM = cls25Sim
            Case "75TH"
                iSimNUM = cls75Sim
        End Select

        GetSimNum = iSimNUM
    End Function
    Public Function GetSimOutput(iSim As String) As Object
        Dim lSimNum As Long
        Dim lOut As Object
        Dim lSimData As SortedDictionary(Of Long, ReturnItem)
        Dim lRI As ReturnItem
        Dim lKey As Long
        Dim lColumn As Long
        Dim lRow As Long

        lSimNum = GetSimNum(iSim)
        lSimData = clsSimDict(lSimNum)
        ReDim lOut(0 To lSimData.Count, 0 To 7)

        lOut(0, lColumn) = "Age"
        lColumn = lColumn + 1
        lOut(0, lColumn) = "Salary"
        lColumn = lColumn + 1
        lOut(0, lColumn) = "Return Pct"
        lColumn = lColumn + 1
        lOut(0, lColumn) = "Beginning Balance"
        lColumn = lColumn + 1
        lOut(0, lColumn) = "Employee Contribution"
        lColumn = lColumn + 1
        lOut(0, lColumn) = "Employer Contribution"
        lColumn = lColumn + 1
        lOut(0, lColumn) = "Asset Return"
        lColumn = lColumn + 1
        lOut(0, lColumn) = "Ending Balance"
        lColumn = lColumn + 1

        For Each lKey In lSimData.Keys
            lRow = lRow + 1
            lColumn = 0
            lRI = lSimData(lKey)
            With lRI
                lOut(lRow, lColumn) = lKey
                lColumn = lColumn + 1
                lOut(lRow, lColumn) = .Salary
                lColumn = lColumn + 1
                lOut(lRow, lColumn) = .ReturnPct
                lColumn = lColumn + 1
                lOut(lRow, lColumn) = .BegBalance
                lColumn = lColumn + 1
                lOut(lRow, lColumn) = .EmployeeCont
                lColumn = lColumn + 1
                lOut(lRow, lColumn) = .EmployerCont
                lColumn = lColumn + 1
                lOut(lRow, lColumn) = .AssetReturn
                lColumn = lColumn + 1
                lOut(lRow, lColumn) = .EndBalance
                lColumn = lColumn + 1
            End With
        Next lKey
        GetSimOutput = lOut
    End Function

    Public Function GetTotalValue(iSim As String, iValue As String) As Double
        Dim lSimNum As Long
        Dim lSimData As SortedDictionary(Of Long, ReturnItem)
        Dim lRI As ReturnItem
        Dim lKey As Long
        Dim lOut As Double

        lSimNum = GetSimNum(iSim)
        lSimData = clsSimDict(lSimNum)
        For Each lKey In lSimData.Keys
            lRI = lSimData(lKey)
            With lRI
                Select Case UCase(iValue)
                    Case "BEGBALANCE"
                        lOut = lOut + .BegBalance
                    Case "EMPLOYEECONT"
                        lOut = lOut + .EmployeeCont
                    Case "EMPLOYERCONT"
                        lOut = lOut + .EmployerCont
                    Case "ASSETRETURN"
                        lOut = lOut + .AssetReturn
                    Case "ENDBALANCE"
                        lOut = lOut + .EndBalance
                    Case "RETURNPCT"
                        lOut = lOut + .ReturnPct
                    Case "SALARY"
                        lOut = lOut + .Salary
                End Select
            End With
        Next lKey

        GetTotalValue = lOut
    End Function



End Class
