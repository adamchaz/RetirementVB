Option Explicit On
Imports Microsoft.Office.Interop.Excel
Module Main
    Private mAssetAllocation As AssetAllocation
    Private mReturns As Returns
    Private mSimResults As SimResults
    Private mScenarioResults As Dictionary(Of String, SimResults)

    Private mName As String
    Private mBegSalary As Double
    Private mBegBal As Double
    Private mSaveRate As Double
    Private mEmployeeMatch As Double
    Private mRetirmentAge As Long
    Private mCurAge As Long
    Private mNumSims As Long

    Public Sub RunSimulation(iCurAge As Long, iRetAge As Long, iBegBal As Double, iBegSalary As Double, iSaveRate As Double, iEmployeeMatch As Double)
        Dim iAge As Long
        Dim lReturns As SortedDictionary(Of Long, ReturnItem)
        Dim lRetItem As ReturnItem
        Dim lBegBal As Double
        Dim lSalary As Double
        Dim lReturn As Double

        lBegBal = iBegBal
        lSalary = iBegSalary
        lReturns = New SortedDictionary(Of Long, ReturnItem)
        For iAge = iCurAge To iRetAge
            lRetItem = New ReturnItem
            lReturn = AssetReturn(iAge)
            With lRetItem
                .BegBalance = lBegBal
                .EmployeeCont = iSaveRate * lSalary
                If iEmployeeMatch <= iSaveRate Then
                    .EmployerCont = lSalary * iEmployeeMatch
                Else
                    .EmployerCont = lSalary * iSaveRate
                End If
                .Salary = lSalary
                .ReturnPct = lReturn
                .AssetReturn = .BegBalance * lReturn
                .EndBalance = .BegBalance + .EmployeeCont + .EmployerCont + .AssetReturn
                lBegBal = .EndBalance
            End With
            lReturns.Add(iAge, lRetItem)
        Next iAge
        Call mSimResults.Add(lReturns)
    End Sub
    Public Sub Main()
        Call Setup

        Call RunUserSIM
        Call RunRiskScenario
        Call RunSaveRateScenario
        Call OutputScenarios()
        Call Cleanup()
        MsgBox("You smell", vbOKOnly, "Fatty")
    End Sub
    Private Sub Setup()
        Call LoadData
        Call DeleteOutputTab("Scenario-")
        Call DeleteOutputTab("MC-")
    End Sub
    Private Sub Cleanup()
        mAssetAllocation = Nothing
        mReturns = Nothing
        mSimResults = Nothing
        mScenarioResults.Clear()
        mScenarioResults = Nothing
    End Sub
    Private Sub LoadData()
        Dim lvar As Object
        Dim lRiskType As String


        Globals.ThisWorkbook.Activate()
        Globals.ThisWorkbook.Sheets("Inputs").Activate
        lvar = Globals.ThisWorkbook.Sheets("Inputs").Range("B1:B9").Value
        mName = lvar(1, 1)
        mCurAge = lvar(2, 1)
        mBegSalary = lvar(3, 1)
        mBegBal = lvar(4, 1)
        lRiskType = lvar(5, 1) & "AssetAll"
        mSaveRate = lvar(6, 1)
        mEmployeeMatch = lvar(7, 1)
        mRetirmentAge = lvar(8, 1)
        mNumSims = lvar(9, 1)


        mReturns = New Returns
        Call mReturns.Load(Globals.ThisWorkbook.Sheets("Returns and Correlation").Range("AssetReturns").Value, Globals.ThisWorkbook.Sheets("Returns and Correlation").range("AssetCorr").Value)
        mAssetAllocation = New AssetAllocation
        Call mAssetAllocation.LoadMatrix(Globals.ThisWorkbook.Sheets("Asset Allocation").Range(lRiskType).Value)

        mSimResults = New SimResults
        mScenarioResults = New Dictionary(Of String, SimResults)

    End Sub
    Public Function AssetReturn(iAge As Long) As Double
        Dim lAssetall As Dictionary(Of String, Double)
        Dim lReturn As Dictionary(Of String, Double)
        Dim lReturnMat(,) As Double  'ReturnMatrix
        Dim lAssetMat(,) As Double 'AssetMatrix
        Dim lAssetReturn(,) As Double
        Dim i As Long
        Dim lKey As String

        lAssetall = mAssetAllocation.GetAssetAll(iAge)
        lReturn = mReturns.Returns

        ReDim lAssetMat(0 To 1, 0 To lAssetall.Count)
        ReDim lReturnMat(0 To lReturn.Count, 0 To 1)
        i = 1
        For Each lKey In lAssetall.Keys
            lAssetMat(1, i) = lAssetall(lKey)
            lReturnMat(i, 1) = lReturn(lKey)
            i = i + 1
        Next lKey
        lAssetReturn = MatrixMath.MatrixMultiply(lAssetMat, lReturnMat)
        AssetReturn = lAssetReturn(1, 1)
    End Function

    Private Sub RunUserSIM()
        Dim i As Long
        mSimResults = New SimResults
        Dim lMCOUT As Object
        For i = 1 To mNumSims
            Call RunSimulation(mCurAge, mRetirmentAge, mBegBal, mBegSalary, mSaveRate, mEmployeeMatch)
        Next i
        Call mSimResults.PrepareResults()
        lMCOUT = mSimResults.GetAllValue(mRetirmentAge, "ENDBALANCE")

        Call OutputDataToSheet("EndingBlances", lMCOUT, False, "MC-")
        Call OutputDataToSheet("25TH", mSimResults.GetSimOutput("25TH"), False, "MC-")
        Call OutputDataToSheet("Median", mSimResults.GetSimOutput("MEDIAN"), False, "MC-")
        Call OutputDataToSheet("75TH", mSimResults.GetSimOutput("75TH"), False, "MC-")
    End Sub

    Private Sub RunSaveRateScenario()
        Dim i As Long
        Dim j As Long

        Dim lSaveRate(6) As Double
        lSaveRate(0) = 0.06
        lSaveRate(1) = 0.07
        lSaveRate(2) = 0.08
        lSaveRate(3) = 0.09
        lSaveRate(4) = 0.1
        lSaveRate(5) = 0.125
        lSaveRate(6) = 0.15

        For j = 0 To UBound(lSaveRate)
            mSimResults = New SimResults
            For i = 1 To mNumSims
                Call RunSimulation(mCurAge, mRetirmentAge, mBegBal, mBegSalary, lSaveRate(j), mEmployeeMatch)
            Next i
            Call mSimResults.PrepareResults()
            Call mScenarioResults.Add("Savings Rate-" & Format(lSaveRate(j), "0.00%"), mSimResults)
        Next j

    End Sub
    Private Sub RunRiskScenario()
        Dim i As Long
        Dim j As Long
        Dim lRiskType As String
        Dim lRisk(2) As String

        lRisk(0) = "Low"
        lRisk(1) = "Mid"
        lRisk(2) = "High"


        For j = 0 To UBound(lRisk)
            mSimResults = New SimResults
            lRiskType = lRisk(j) & "AssetAll"
            mAssetAllocation = Nothing
            mAssetAllocation = New AssetAllocation
            Call mAssetAllocation.LoadMatrix(Globals.ThisWorkbook.Worksheets("Asset Allocation").Range(lRiskType).Value)
            For i = 1 To mNumSims
                Call RunSimulation(mCurAge, mRetirmentAge, mBegBal, mBegSalary, mSaveRate, mEmployeeMatch)
            Next i
            Call mSimResults.PrepareResults()
            Call mScenarioResults.Add("Risk Level-" & lRisk(j), mSimResults)
        Next j

    End Sub

    Public Sub OutputScenarios()
        Call OutputDataToSheet("Ending Blances", GetScenarioOutput("Ending Balance"), False, "Scenario-")
        Call OutputDataToSheet("Years in Retirement", GetScenarioOutput("YEARS IN RETIREMENT"), False, "Scenario-")
        Call OutputDataToSheet("Total Employee Contributions", GetScenarioOutput("EMPLOYEE CONT"), False, "Scenario-")
        Call OutputDataToSheet("Total Employer Contributions", GetScenarioOutput("EMPLOYER CONT"), False, "Scenario-")
        Call OutputDataToSheet("Total Asset Returns", GetScenarioOutput("ASSETRETURN"), False, "Scenario-")
    End Sub
    Public Function GetScenarioOutput(iValueType As String) As Object
        Dim lOut As Object
        Dim lNumScenarios As Long
        Dim lColumn As Long
        Dim lRow As Long
        Dim lKey As String
        Dim lSimResult As SimResults
        Dim lColumns(4) As String
        Dim i As Long


        lColumns(0) = "MIN"
        lColumns(1) = "25TH"
        lColumns(2) = "MEDIAN"
        lColumns(3) = "75TH"
        lColumns(4) = "MAX"

        lNumScenarios = mScenarioResults.Count
        ReDim lOut(0 To lNumScenarios, 0 To 5)
        lOut(0, 0) = "Scenario"
        lOut(0, 1) = "Min"
        lOut(0, 2) = "25th"
        lOut(0, 3) = "Median"
        lOut(0, 4) = "75th"
        lOut(0, 5) = "Max"

        lRow = 1
        For Each lKey In mScenarioResults.Keys
            lSimResult = mScenarioResults(lKey)
            lOut(lRow, 0) = lKey
            lColumn = 1
            For i = 0 To UBound(lColumns)
                Select Case UCase(iValueType)
                    Case "ENDING BALANCE"
                        lOut(lRow, lColumn) = lSimResult.GetValue(lColumns(i), mRetirmentAge, "ENDBALANCE")
                    Case "YEARS IN RETIREMENT"
                        lOut(lRow, lColumn) = lSimResult.GetValue(lColumns(i), mRetirmentAge, "ENDBALANCE") / lSimResult.GetValue(lColumns(i), mRetirmentAge, "Salary")
                    Case "EMPLOYEE CONT"
                        lOut(lRow, lColumn) = lSimResult.GetTotalValue(lColumns(i), "EmployeeCont")
                    Case "EMPLOYER CONT"
                        lOut(lRow, lColumn) = lSimResult.GetTotalValue(lColumns(i), "EmployeRCont")
                    Case "ASSETRETURN"
                        lOut(lRow, lColumn) = lSimResult.GetTotalValue(lColumns(i), "ASSETRETURN")
                End Select
                lColumn = lColumn + 1
            Next i
            lRow = lRow + 1
        Next lKey

        GetScenarioOutput = lOut
    End Function

End Module
