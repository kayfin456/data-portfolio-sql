'---------------------------------------------------------------
' Project: Data Cleaning and Interval Automation in Excel
' Author: Kayden Finlayson
' Description:
'   Automates cleaning, formatting, and intervalizing raw Excel data exports.
'   Demonstrates VBA logic for data transformation, looping, and dynamic range handling.
'---------------------------------------------------------------

Option Explicit

Sub CleanPaste()

    Dim Lr As Long 'Last Row
    Dim Fr As Long 'First Row
    Dim Row As Long
    Dim i As Long
    
    'Copy data and paste in  new sheet
    Sheet1.Columns("B:M").Copy
        Sheets.Add , After:=Sheet1
        Selection.PasteSpecial Paste:=xlPasteValues
        
    'Start to remove useless columns and rows
    Rows("1:6").Delete Shift:=xlUp
    Range("C:C,E:E,F:F,G:G,I:I,J:J,K:K").Delete Shift:=xlToLeft
    
    'Create name column and interval help columns
    Range("A1") = "Name"
    Range("A2") = "=IF(LEFT(C2,3) = ""ID:"",B2,A1)"
    Range("E1") = "Interval Start"
    Range("F1") = "Interval End"
    Range("G1") = "# Rows"
    Range("E2") = "=FLOOR(C2,TIME(0,15,0))"
    Range("F2") = "=Floor(D2,TIME(0,15,0))"
    Range("G2") = "=IF(E2=F2,0,ROUND((((F2-E2)*24)/0.25)+1,0))"
    Lr = Range("B2").End(xlDown).Row
    Range("A2").AutoFill Range("A2", "A" & Lr)
    Range("E2:G2").AutoFill Range("E2:G2", "E" & Lr & ":" & "G" & Lr)
    
    'Remove forumulas from cells
    Range("A1").CurrentRegion.Copy
    Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    
    'Delete rows with no data
    For i = Range("G1").End(xlDown).Row To 1 Step -1
        If IsError(Cells(i, 7).Value) Then
        Cells(i, 7).EntireRow.Delete
        End If
    Next i
    
    IntervalizeData
    FinalClean
    
    
End Sub

Sub IntervalizeData()

    Dim i As Long
    Dim j As Long
    
    For i = Range("G1").End(xlDown).Row To 2 Step -1
        If Cells(i, 7) > 0 Then
        Cells(i, 7).EntireRow.Offset(1).Resize(Cells(i, 7).Value).Insert Shift:=xlDown
            For j = 1 To Cells(i, 7).Value Step 1
                Cells(i, 1).Offset(j).Value = Cells(i, 1).Value
                Cells(i, 2).Offset(j).Value = Cells(i, 2).Value
                Cells(i, 5).Offset(j).Value = 0
                Cells(i, 6).Offset(j).Value = 0
                Cells(i, 7).Offset(j).Value = 0
                Select Case j
                    Case 1
                        Cells(i, 3).Offset(j) = Cells(i, 3).Value
                        Cells(i, 4).Offset(j) = (Cells(i, 5).Value + (15 / 1440))
                    Case 2 To (Cells(i, 7).Value - 1)
                        Cells(i, 3).Offset(j) = Cells(i, 4).Offset((j - 1)).Value
                        Cells(i, 4).Offset(j) = (Cells(i, 3).Offset(j).Value + (15 / 1440))
                    Case Cells(i, 7).Value
                        Cells(i, 3).Offset(j) = Cells(i, 4).Offset((j - 1)).Value
                        Cells(i, 4).Offset(j) = Cells(i, 4).Value
                        
                End Select
                
            Next j
        Cells(i, 1).EntireRow.Delete
        End If
    Next i
    
End Sub

Sub FinalClean()
    
    Dim Lr As Long
    
    Columns("E:G").Delete
    Range("E1") = "Date"
    Range("F1") = "DOW"
    Range("G1") = "Interval"
    Range("H1") = "Duration"
    Range("E2") = "=TEXT(C2,""MM/DD/YYYY"")"
    Range("F2") = "=TEXT(C2,""DDDD"")"
    Range("G2") = "=ROUND(FLOOR(C2,TIME(0,15,0))-INT(FLOOR(C2,TIME(0,15,0))),5)"
    Range("H2") = "=(D2-C2)*1440"
    Lr = Range("B2").End(xlDown).Row
    Range("E2:H2").AutoFill Range("E2:H2", "E" & Lr & ":" & "H" & Lr)
    Rows("1:1").Font.Bold = True
    Range("A1").CurrentRegion.Copy
    Range("A1").PasteSpecial xlPasteValues
    Application.CutCopyMode = False
    Columns("C:D").EntireColumn.Delete
    Columns("E").NumberFormat = "[$-x-systime]h:mm:ss AM/PM"
    Columns("F").NumberFormat = "0.00"
    Columns("A:F").EntireColumn.AutoFit
End Sub
