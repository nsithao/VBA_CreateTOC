Attribute VB_Name = "Module1"
#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr) 'For 64 Bit Systems
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) 'For 32 Bit Systems
#End If

Function OpenOneExcelFile() As Workbook
' Date:     2016-01-30_16-13-00
' Author:   Si Thao
' Function: open 1 excel file then pass to other function for processing

    Dim FilesToOpen
    Dim wkbTemp As Workbook
             
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
     
    '--- change directory folder to current workbook running the vba code ---
    ChDir ThisWorkbook.Path
    
    '--- open dialog for user to select file ---
    FilesToOpen = Application.GetOpenFilename _
                    (FileFilter:="Text Files (*.*), *.*", _
                     MultiSelect:=False, Title:="Text Files to Open")
    
    If TypeName(FilesToOpen) = "Boolean" Then
        MsgBox "No Files were selected"
        GoTo ExitHandler
    End If
     
    '--- Assign wkbTemp to the selected workbook by user and open that workbook ---
    Set wkbTemp = Workbooks.Open(Filename:=FilesToOpen)
    
    Set OpenOneExcelFile = wkbTemp
    
ExitHandler:
    Application.ScreenUpdating = True
    Set wkbAll = Nothing
    Set wkbTemp = Nothing
    Exit Function
     
ErrHandler:
    MsgBox Err.Description
    Resume ExitHandler
    
End Function

Sub addNewSheetAtBegin(tmpWB As Workbook)
' Date:     2016-02-01_
' Author:   Si Thao
' Function: add new worksheet in front of current workbook
    Dim WS As Worksheet
    
    If tmpWB.Sheets(1).Name <> "TOC" Then
        Set WS = tmpWB.Sheets.Add(Before:=Worksheets(1))
        WS.Name = "TOC"
        WS.Tab.Color = RGB(255, 165, 0)
    Else
        ' workbook has TOC, change color if not yet
        tmpWB.Sheets("TOC").Tab.Color = RGB(255, 165, 0)
    End If
End Sub

Sub listAndChangeSheetName(tmpWB As Workbook)
' Date:     2016-02-01_
' Author:   Si Thao
' Function: contruct a Table of Content (TOC) of the current workbook and hyperlink
    
    Dim counterWorkSheet As Integer
    Dim counterWorkSheetTemp As Integer
    Dim iTotalWorkSheets As Integer
    Dim strTempFunction As String
    Dim iCurrentRow As Integer
    Dim flagNewSheet As Boolean
    Dim flagOldSheet As Boolean
    
    iTotalWorkSheets = tmpWB.Worksheets.Count
    
    tmpWB.Sheets("TOC").Range("D2") = "Table of Content"
    tmpWB.Sheets("TOC").Range("B4") = "Sheet No"
    tmpWB.Sheets("TOC").Range("B4").HorizontalAlignment = xlCenter
    tmpWB.Sheets("TOC").Range("C4") = "Go to"
    tmpWB.Sheets("TOC").Range("C4").HorizontalAlignment = xlCenter
    tmpWB.Sheets("TOC").Range("D4") = "Sheet content"
    
    tmpWB.Sheets("TOC").Range("B5") = "1"
    tmpWB.Sheets("TOC").Range("B5").HorizontalAlignment = xlCenter
    tmpWB.Sheets("TOC").Range("C5") = "1"
    tmpWB.Sheets("TOC").Range("C5").HorizontalAlignment = xlCenter
    tmpWB.Sheets("TOC").Range("D5") = "TOC"
    
Recheck:
    ' the first loop to verify that there is no extra new sheet was inserted in between of processed sheets from last run
    flagNewSheet = False
    flagOldSheet = False
    
    For counterWorkSheetTemp = 2 To iTotalWorkSheets
        If Not IsNumeric(tmpWB.Worksheets(counterWorkSheetTemp).Name) Then ' all the actual text will be converted into 0
            flagNewSheet = True
        Else
            flagOldSheet = True
            If flagNewSheet = True Then
                MsgBox "Move the below worksheet to the end before doing TOC: " & tmpWB.Worksheets(counterWorkSheetTemp - 1).Name
                tmpWB.Worksheets(counterWorkSheetTemp - 1).Move After:=Worksheets(Worksheets.Count)
                GoTo Recheck
            End If
        End If
    Next counterWorkSheetTemp
    
    ' actual loop to process TOC
    For counterWorkSheet = 2 To iTotalWorkSheets
        strReturnTOC = "=HYPERLINK(""[.\]TOC!A1"",""return to TOC"")"
        tmpWB.Sheets("TOC").Range("B" & counterWorkSheet).Offset(4, 0).HorizontalAlignment = xlCenter
        
        If counterWorkSheet <> tmpWB.Worksheets(counterWorkSheet).Name Then
            tmpWB.Sheets("TOC").Range("B" & counterWorkSheet).Offset(4, 0) = counterWorkSheet
            
            '    =HYPERLINK(CONCATENATE("#",$B5,"!A1"),$B5)
            'strTempFunction = "=HYPERLINK(CONCATENATE(""#"",$B5,""!A1""),$B5)"
            
            iCurrentRow = 4 + counterWorkSheet
            strTempFunction = "=HYPERLINK(CONCATENATE(""#"",$B" & iCurrentRow & ",""!A1""),$B" & iCurrentRow & ")"
            tmpWB.Sheets("TOC").Range("C" & counterWorkSheet).Offset(4, 0).Formula = strTempFunction
            tmpWB.Sheets("TOC").Range("C" & counterWorkSheet).Offset(4, 0).HorizontalAlignment = xlCenter
            tmpWB.Worksheets(counterWorkSheet).Range("D1") = tmpWB.Worksheets(counterWorkSheet).Name
            tmpWB.Worksheets(counterWorkSheet).Range("D1").Font.Color = RGB(255, 0, 0)
            tmpWB.Sheets("TOC").Range("D" & counterWorkSheet).Offset(4, 0) = tmpWB.Worksheets(counterWorkSheet).Name
            tmpWB.Worksheets(counterWorkSheet).Name = counterWorkSheet
        Else
            'MsgBox "Same name, skip"
        End If
        
        ' create a hyperlink back to the TOC sheet
        If tmpWB.Worksheets(counterWorkSheet).Range("B1") = "" Then
            tmpWB.Worksheets(counterWorkSheet).Range("B1").Formula = strReturnTOC
        Else
            'do nothing
        End If
        
        ' change text color to red
        If tmpWB.Worksheets(counterWorkSheet).Range("D1").Value = "" Then
            tmpWB.Worksheets(counterWorkSheet).Range("D1") = tmpWB.Sheets("TOC").Range("D" & counterWorkSheet).Offset(4, 0).Value
            tmpWB.Worksheets(counterWorkSheet).Range("D1").Font.Color = RGB(255, 0, 0)
        Else
            tmpWB.Worksheets(counterWorkSheet).Range("D1").Font.Color = RGB(255, 0, 0)
        End If
        
    Next counterWorkSheet

End Sub

Sub createTOC()
Dim tmpWB As Workbook

    Set tmpWB = OpenOneExcelFile()

    Call addNewSheetAtBegin(tmpWB)
    
    Call listAndChangeSheetName(tmpWB)
          
    MsgBox "Finish !"
End Sub
