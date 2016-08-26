Option Explicit On
Imports Microsoft.Office.Interop.Excel
Imports Microsoft.Office.Interop.Excel.XlSearchOrder
Imports Microsoft.Office.Interop.Excel.XlFindLookIn
Imports Microsoft.Office.Interop.Excel.XlLookAt
Imports Microsoft.Office.Interop.Excel.XlSearchDirection
Imports Microsoft.Office.Interop.Excel.XlBordersIndex
Imports Microsoft.Office.Interop.Excel.XlBorderWeight
Imports Microsoft.Office.Interop.Excel.XlLineStyle
Imports Microsoft.Office.Interop.Excel.Constants

Module UserInterface
    Public Sub DeleteOutputTab(iPrefix As String)
        On Error Resume Next
        Dim lsheet As Worksheet
        Dim lObject As Object
        Globals.ThisWorkbook.Application.DisplayAlerts = False

        For Each lObject In Globals.ThisWorkbook.Worksheets
            'Debug.Print lObject.Name
            If UCase(Left(lObject.Name, Len(iPrefix))) = UCase(iPrefix) Then
                lObject.Delete
            End If
        Next lObject
        Globals.ThisWorkbook.Application.DisplayAlerts = True
    End Sub
    Public Sub OutputDataToSheet(iName As String, lOutput As Object, iNewTab As Boolean, iPrefix As String)
        Static lColumn As Long
        Dim lRange As Range
        Dim lRowoffset As Long
        Dim lsheet As Worksheet
        Dim lTabName As String
        Dim lsheetfound As Boolean
        Dim lObject As Object
        Dim lActivesheet As Worksheet
        Dim i As Long

        'lRange = Range("A1").Offset(0, lColumn)

        If iNewTab Then
            lTabName = iPrefix & iName
        Else
            lTabName = iPrefix & "Output"
        End If


        'Find Tab or create tab
        For Each lObject In Globals.ThisWorkbook.Sheets
            If TypeName(lObject) = "Worksheet" Then
                lsheet = lObject
                If UCase(lsheet.Name) = UCase(lTabName) Then
                    lsheetfound = True
                    Exit For
                End If
            End If
        Next lObject
        If lsheetfound = False Then
            lActivesheet = Globals.ThisWorkbook.Sheets.Add(, Globals.ThisWorkbook.Sheets(Globals.ThisWorkbook.Sheets.Count), , "Worksheet")
            lActivesheet.Name = lTabName
            lActivesheet.Tab.Color = 5287936
            'Range("F3").Activate
            'ActiveWindow.FreezePanes = True
            lColumn = 1
        Else
            With Globals.ThisWorkbook.Sheets(lTabName)
                lColumn = .Cells.Find("*", After:= .Cells(1),
                            LookIn:=xlFormulas, LookAt:=xlWhole,
                            SearchDirection:=xlPrevious,
                            SearchOrder:=xlByColumns).Column
            End With
            lColumn = lColumn + 1
        End If
        If iNewTab Then
            Globals.ThisWorkbook.Sheets(lTabName).Cells.Clear
            lRowoffset = 0
        Else
            lRowoffset = 1
        End If

        lRange = Globals.ThisWorkbook.Sheets(lTabName).Range("A1").Offset(lRowoffset, lColumn)
        lsheet = Globals.ThisWorkbook.Sheets(lTabName)
        Dim lRows As Long
        Dim lColumns As Long
        lRows = UBound(lOutput, 1) - LBound(lOutput, 1)
        lColumns = UBound(lOutput, 2) - LBound(lOutput, 2)

        lsheet.Range(lRange, lRange.Offset(lRows, lColumns)).Value = lOutput
        '    For i = 0 To lRows
        '        For j = 0 To lColumns
        '            lRange.Offset(i, j).Value = lOutput(i, j)
        '        Next j
        '    Next i

        For i = LBound(lOutput, 2) To UBound(lOutput, 2)
            Select Case VarType(lOutput(1, i))


                Case vbDouble
                    'lrange.Offset(0, i).Columns.Style = "Comma"

                    'lrange.Offset(0, 1).Columns
                    'thisworkbook.Sheets(itabname).
                    Globals.ThisWorkbook.Sheets(lTabName).Columns(lRange.Offset(1, i).Column).Style = "Comma"
                Case vbDate

            End Select
        Next i
        If iNewTab = False Then
            lColumn = lColumn + UBound(lOutput, 2) + 2
            lRange.Offset(-lRowoffset, 0).Value = iName
            lsheet.Range(lRange.Offset(-lRowoffset, 0), lRange.Offset(-lRowoffset, UBound(lOutput, 2))).HorizontalAlignment = xlCenterAcrossSelection
            lsheet.Range(lRange.Offset(-lRowoffset, 0), lRange.Offset(-lRowoffset, UBound(lOutput, 2))).Borders(xlEdgeRight).LineStyle = xlContinuous
            lsheet.Range(lRange.Offset(-lRowoffset, 0), lRange.Offset(-lRowoffset, UBound(lOutput, 2))).Borders(xlEdgeRight).Weight = xlMedium
            lsheet.Range(lRange.Offset(-lRowoffset, 0), lRange.Offset(-lRowoffset, UBound(lOutput, 2))).Borders(xlEdgeLeft).LineStyle = xlContinuous
            lsheet.Range(lRange.Offset(-lRowoffset, 0), lRange.Offset(-lRowoffset, UBound(lOutput, 2))).Borders(xlEdgeLeft).Weight = xlMedium

        Else

        End If
        With lsheet.Range(lRange, lRange.Offset(UBound(lOutput, 1), UBound(lOutput, 2)))
            .Borders(xlEdgeRight).LineStyle = xlContinuous
            .Borders(xlEdgeRight).Weight = xlMedium
            .Borders(xlEdgeLeft).LineStyle = xlContinuous
            .Borders(xlEdgeLeft).Weight = xlMedium
            .Borders(xlEdgeBottom).LineStyle = xlContinuous
            .Borders(xlEdgeBottom).Weight = xlMedium
            .Borders(xlEdgeTop).LineStyle = xlContinuous
            .Borders(xlEdgeTop).Weight = xlMedium
        End With

        Globals.ThisWorkbook.Sheets(lTabName).Cells.Columns.AutoFit
    End Sub
End Module
