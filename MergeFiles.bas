Attribute VB_Name = "modMergeFiles"
Option Explicit

'==========================================================================
' Merge all Excel files and all visible sheets from a folder into one
' master sheet. Header is copied once; two tracking columns are added.
'==========================================================================
Sub Merge_All_Files_All_Sheets()

    Dim FolderPath As String, FileName As String
    Dim wbSource As Workbook
    Dim wsSource As Worksheet, wsDest As Worksheet
    Dim rngLast As Range
    Dim LastRow As Long, LastCol As Long
    Dim DestRow As Long, DataRows As Long
    Dim HeaderCols As Long
    Dim HeaderCopied As Boolean
    Dim FileCount As Long, SheetCount As Long, RowCount As Long
    Dim SkippedList As String

    'Step 1: Select the folder that contains the Excel files
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Select Folder Containing Excel Files"
        If .Show <> -1 Then Exit Sub
        FolderPath = .SelectedItems(1) & "\"
    End With

    On Error GoTo CleanUp
    Application.ScreenUpdating = False

    'Step 2: Create the destination sheet
    Set wsDest = ThisWorkbook.Sheets.Add
    wsDest.Name = "Combined_Data_" & Format(Now, "hhmmss")
    DestRow = 1
    HeaderCopied = False

    'Step 3: Loop through every Excel file in the folder
    FileName = Dir(FolderPath & "*.xls*")

    Do While FileName <> ""

        'Skip temporary lock files (~$...) and this macro workbook itself
        If Left(FileName, 2) <> "~$" And FileName <> ThisWorkbook.Name Then

            Set wbSource = Workbooks.Open(FolderPath & FileName, ReadOnly:=True)
            FileCount = FileCount + 1

            'Step 4: Loop through every visible sheet in the file
            For Each wsSource In wbSource.Worksheets

                If wsSource.Visible = xlSheetVisible Then

                    'Last used row and column, checked across the whole sheet
                    Set rngLast = wsSource.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlPrevious)

                    If Not rngLast Is Nothing Then

                        LastRow = rngLast.Row
                        LastCol = wsSource.Cells.Find(What:="*", LookIn:=xlFormulas, _
                                  SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

                        If LastRow > 1 Then

                            'Copy the header only once, and add tracking columns
                            If Not HeaderCopied Then
                                HeaderCols = LastCol
                                wsSource.Range(wsSource.Cells(1, 1), wsSource.Cells(1, LastCol)).Copy _
                                    wsDest.Cells(1, 1)
                                wsDest.Cells(1, HeaderCols + 1).Value = "Source File"
                                wsDest.Cells(1, HeaderCols + 2).Value = "Sheet Name"
                                DestRow = 2
                                HeaderCopied = True
                            End If

                            'Merge only if the structure matches the first sheet
                            If LastCol = HeaderCols Then

                                DataRows = LastRow - 1

                                wsSource.Range(wsSource.Cells(2, 1), wsSource.Cells(LastRow, LastCol)).Copy _
                                    wsDest.Cells(DestRow, 1)

                                wsDest.Cells(DestRow, HeaderCols + 1).Resize(DataRows, 1).Value = wbSource.Name
                                wsDest.Cells(DestRow, HeaderCols + 2).Resize(DataRows, 1).Value = wsSource.Name

                                DestRow = DestRow + DataRows
                                RowCount = RowCount + DataRows
                                SheetCount = SheetCount + 1

                            Else
                                SkippedList = SkippedList & vbLf & wbSource.Name & " / " & wsSource.Name
                            End If

                        End If
                    End If
                End If

            Next wsSource

            wbSource.Close SaveChanges:=False

        End If

        FileName = Dir
    Loop

    wsDest.Columns.AutoFit

    'Step 5: Summary, useful for checking the merged row count
    Dim Msg As String
    Msg = "Files merged: " & FileCount & vbLf & _
          "Sheets merged: " & SheetCount & vbLf & _
          "Data rows merged: " & RowCount
    If SkippedList <> "" Then Msg = Msg & vbLf & vbLf & "Skipped (different column count):" & SkippedList
    MsgBox Msg, vbInformation, "Merge complete"

CleanUp:
    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        Dim ErrText As String
        ErrText = Err.Description
        On Error Resume Next
        If Not wbSource Is Nothing Then wbSource.Close SaveChanges:=False
        MsgBox "Macro stopped: " & ErrText, vbExclamation
    End If

End Sub
