Attribute VB_Name = "Module1"
Sub SplitLinesDown_InsertRows()
    Dim c As Range, arr As Variant
    Dim i As Long, n As Long
    Dim txt As String
    
    If TypeName(Selection) <> "Range" Then
        MsgBox "Please select the cell(s) you want to process.", vbExclamation
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Work from bottom to top to avoid shifting problems when inserting rows
    For Each c In Selection.Areas
        Dim r As Range
        For i = c.Cells.Count To 1 Step -1
            Set r = c.Cells(i)
            If Len(r.Value2) > 0 Then
                ' Normalize line breaks to vbLf
                txt = CStr(r.Value2)
                txt = Replace(txt, vbCrLf, vbLf)
                txt = Replace(txt, vbCr, vbLf)
                
                arr = Split(txt, vbLf)
                n = UBound(arr) ' zero-based
                
                If n >= 0 Then
                    ' First line remains in the original cell
                    r.Value = arr(0)
                    
                    ' If more lines exist, insert rows and place them below
                    If n >= 1 Then
                        ' Insert exactly n rows below this row
                        r.Offset(1).Resize(n).EntireRow.Insert
                        
                        ' Write subsequent lines into the newly inserted rows, same column
                        Dim k As Long
                        For k = 1 To n
                            r.Offset(k, 0).Value = arr(k)
                        Next k
                    End If
                End If
            End If
        Next i
    Next c
    
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
