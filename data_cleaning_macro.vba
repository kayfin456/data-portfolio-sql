'---------------------------------------------------------------
' Project: Excel Data Cleaning and Interval Automation
' Author: Kayden Finlayson
'
' Description:
'   Automates the cleaning, formatting, and intervalization of
'   raw data exported from a reporting system. This macro removes
'   unnecessary rows/columns, parses timestamps, and expands data
'   into 15-minute intervals for analysis.
'---------------------------------------------------------------

Option Explicit

Sub CleanPaste()

    Dim Lr As Long  ' Last Row
    Dim i As Long

    ' Copy data columns from the original sheet into a new one
    Sheet1.Columns("B:M").Copy
    Sheets.Add , After:=Sheet1
    Selection.PasteSpecial Paste:=xlPasteValues

    ' Remove header rows and unwanted columns
    Rows("1:6").Delete Shift:=xlUp
    Range("C:C,E:E,F:F,G:G,I:I,J:J,K:K").Delete Shift:=xlToLeft

    ' Add helper and header columns
    Range("A1").Value = "Name"
    Range("A2").Formula = "=IF(LEFT(C2,3)=""ID:"",B2,A1)"
    Range("E1").Value = "Interval Start"
    Range("F1").Value = "Interval End"
    Range("G1").Value = "# Rows"
    Range("E2").Formula = "=FLOOR(C2,TIME(0,15,0))"
    Range("F2").Formula = "=FLOOR(D2,TIME(0,15,0))"
    Range("G2").Formula = "=IF(E2=F2,0,ROUND((((F2-E2)*24)/0.25)+1,0))"

    ' Autofill helper formulas
    Lr = Range("B2").End(xlDown).Row
    Range("A2").AutoFill Destination:=Range("A2:A" & Lr)
    Range("E2:G2").AutoFill Destination:=Range("E2:G" & Lr)

    ' Replace formulas with values
    Range("A1").CurrentRegion.Copy
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' Remove error rows
    For i = Range("G1").End(xlDown).Row To 1 Step -1
        If IsError(Cells(i, 7).Value) Then
            Cells(i, 7).EntireRow.Delete
        End If
    Next i

    ' Expand into 15-minute intervals and finalize
    Call IntervalizeData
    Call FinalClean

End Sub

'---------------------------------------------------------------
' Expands each record into 15-minute intervals based on duration
'---------------------------------------------------------------
Sub IntervalizeData()

    Dim i As Long
    Dim j As Long

    For i = Range("G1").End(xlDown).Row To 2 Step -1
        If Cells(i, 7).Value > 0 Then

            ' Insert new rows equal to number of intervals
            Cells(i, 7).EntireRow.Offset(1).Resize(Cells(i, 7).Value).Insert Shift:=xlDown

            ' Populate interval rows
            For j = 1 To Cells(i, 7).Value
                Cells(i, 1).Offset(j).Value = Cells(i, 1).Value
                Cells(i, 2).Offset(j).Value = Cells(i, 2).Value
                Cells(i, 5).Offset(j).Value = 0
                Cells(i, 6).Offset(j).Value = 0
                Cells(i, 7).Offset(j).Value = 0

                Select Case j
                    Case 1
                        Cells(i, 3).Offset(j).Value = Cells(i, 3).Value
                        Cells(i, 4).Offset(j).Value = (Cells(i, 5).Value + (15 / 1440))
                    Case 2 To (Cells(i, 7).Value - 1)
                        Cells(i, 3).Offset(j).Value = Cells(i, 4).Offset(j - 1).Value
                        Cells(i, 4).Offset(j).Value = (Cells(i, 3).Offset(j).Value + (15 / 1440))
                    Case Cells(i, 7).Value
                        Cells(i, 3).Offset(j).Value = Cells(i, 4).Offset(j - 1).Value
                        Cells(i, 4).Offset(j).Value = Cells(i, 4).Value
                End Select
            Next j

            ' Remove original aggregate row
            Cells(i, 1).EntireRow.Delete
        End If
    Next i

End Sub

'---------------------------------------------------------------
' Cleans and formats final dataset for analysis
'---------------------------------------------------------------
Sub FinalClean()

    Dim Lr As Long

    ' Remove temporary columns
    Columns("E:G").Delete

    ' Create final header structure
    Range("E1").Value = "Date"
    Range("F1").Value = "DayOfWeek"
    Range("G1").Value = "Interval"
    Range("H1").Value = "Duration"

    ' Add formulas
    Range("E2").Formula = "=TEXT(C2,""MM/DD/YYYY"")"
    Range("F2").Formula = "=TEXT(C2,""DDDD"")"
    Range("G2").Formula = "=ROUND(FLOOR(C2,TIME(0,15,0))-INT(FLOOR(C2,TIME(0,15,0))),5)"
    Range("H2").Formula = "=(D2-C2)*1440"

    ' Autofill down
    Lr = Range("B2").End(xlDown).Row
    Range("E2:H2").AutoFill Destination:=Range("E2:H" & Lr)

    ' Finalize layout
    Rows("1:1").Font.Bold = True
    Range("A1").CurrentRegion.Copy
    Range("A1").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False
    Columns("C:D").EntireColumn.Delete
    Columns("E").NumberFormat = "[$-x-systime]h:mm:ss AM/PM"
    Columns("F").NumberFormat = "0.00"
    Columns("A:F").AutoFit

End Sub
