
varInputFolder = "C:\Users\DELL\Desktop\Data"
varOutputFolder = "C:\Users\DELL\Desktop\Data"
varFileNames = "DDAOTU.xlsx,DDASSR.xlsx,DDAUFT.xlsx,DDAVVV.xlsx,DDAWW.xlsx"
status = MergeExcelFilesData(varInputFolder, varOutputFolder, varFileNames)


Function MergeExcelFilesData(varInputFolder, varOutputFolder, varFileNames)
  Dim objFSO, objFolder, objFile, objExcel1, objWorkSheet, objRange, objRows, strExtension, fileNamesCount, aryFileNames, loopC
  Dim varExpFileName
  varFileCount = 0
  reportCode = 2
  ' Specify folder.
  Set objExcel1 = CreateObject("Excel.Application")
  Set objExcel2 = CreateObject("Excel.Application")
  objExcel1.Visible = False
  objExcel2.Visible = False
  Set objWorkbook2  = objExcel2.Workbooks.Add
  Set objWorksheet2 = objWorkbook2.Worksheets("Sheet1")
  ' Enumerate files in the folder.
  Set objFSO = CreateObject("Scripting.FileSystemObject")
  Set objFolder = objFSO.GetFolder(varInputFolder)
  
  aryFileNames = Split(varFileNames,",",-1,1)
  fileNamesCount = UBound(aryFileNames)+1
  For loopC = 0 to fileNamesCount-1 
    varExpFileName = aryFileNames(loopC)
    For Each objFile In objFolder.Files
      varFileName = objFile.Name
      If (varExpFileName = varFileName) Then
        varFileCount = varFileCount + 1
        ' Select only Excel spreadsheet file.
        strExtension = objFSO.GetExtensionName(objFile.Path)
        If (strExtension = "xls") Or (strExtension = "xlsx") Then
          ' Open each spreadsheet and count the number of rows.
          Set objWorkBook = objExcel1.Workbooks.Open (objFile.Path)
          Set objWorkSheet  = objWorkBook.Worksheets(1)
          varRowsCount  = objWorkSheet.UsedRange.Rows.Count
          varColumnsCount = objWorkSheet.UsedRange.Columns.Count
          varEndColAlpha = ConvertToLetter(varColumnsCOunt)
          If (varFileCount = 1) Then
            varRowsRange  = "A1" & ":" & varEndColAlpha & varRowsCount
          Else
            varRowsRange  = "A2" & ":" & varEndColAlpha & varRowsCount
          End If
          Set objRange  = objWorkSheet.Range(varRowsRange)
          objRange.Copy
          
          If (varFileCount = 1) Then
            varPasteRowRange  = "A1"
            varPasteRowCount = varRowsCount + 1
          Else
            If (reportCode = 2) Then
              reportCode = 1
            Else
              varPasteRowCount = varPasteRowCount + varRowsCount-1
            End If
            varPasteRowRange   = "A" & varPasteRowCount
          End If
          
          objWorksheet2.Range(varPasteRowRange).PasteSpecial -4163
          objWorkBook.Close
          Exit For
        End If
      End If
    Next
  Next
  varTimeStamp = Replace(Replace(Now(), "/", "-"), ":", ".")
  for col=1 to 4
    objWorksheet2.columns(col).AutoFit()
  next
  objWorkbook2.SaveAs(varOutputFolder & "\" & "OutputFile_" & varTimeStamp & ".xlsx")
  objWorkbook2.Close

  TerminateExcelProcess()
End Function

Function ConvertToLetter(iCol)
   Dim iAlpha
   Dim iRemainder
   iAlpha = Int(iCol / 27)
   iRemainder = iCol - (iAlpha * 26)
   If iAlpha > 0 Then
      ConvertToLetter = UCase(Chr(iAlpha + 64))
   End If
   If iRemainder > 0 Then
      ConvertToLetter = UCase(ConvertToLetter & Chr(iRemainder + 64))
   End If
End Function

Function TerminateExcelProcess()
  Dim Process 
  For Each Process In GetObject("winmgmts:").ExecQuery("Select Name from Win32_Process Where Name = 'EXCEL.EXE'")
     Process.Terminate
  Next
End Function
