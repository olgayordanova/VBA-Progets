Attribute VB_Name = "Module_ZG"


Sub OpenFile() '(ByVal i, ByRef SecondFileName)
Dim oExcel As Excel.Application
Dim iFileNum As Integer
Dim sFileName As String
Dim fd As FileDialog
Dim sBuf As String
Dim FirstFile As String
Dim SecondFileName As String
Dim Lend As Long
Dim tod As Long

FirstFile = ThisWorkbook.Name

Response = MsgBox("Ще прехвърляте ли данни за предходен ден", vbYesNoCancel + vbInformation, "Импорт на Данни")
If Response = vbYes Then
    Call TransferDataPrevDay(FirstFile)
    GoTo startcopydata
ElseIf Response = vbCancel Then
    Exit Sub
ElseIf Response = vbNo Then
    GoTo startcopydata
End If

startcopydata:
Response = MsgBox("Изберете директорията с файловете за качване", vbYesNoCancel + vbInformation, "Импорт на Данни")
If Response = vbNo Or Response = vbCancel Then Exit Sub



tod = Workbooks(FirstFile).Worksheets("Assets").Cells(1, 20).Value
Workbooks(FirstFile).Worksheets("Assets").Activate
Workbooks(FirstFile).Worksheets("Assets").Range(Worksheets("Assets").Cells(2, 1), Worksheets("Assets").Cells(tod, 16)).ClearContents

Set fd = Application.FileDialog(msoFileDialogFilePicker)
Dim vrtSelectedItem As Variant
With fd
    If .Show = -1 Then
        For Each vrtSelectedItem In .SelectedItems
            sFileName = vrtSelectedItem
            Lend = Len(Dir$(sFileName))
            SecondFileName = Right(sFileName, Lend)
            'MsgBox "The path is: " & vrtSelectedItem
            If Len(Dir$(sFileName)) = 0 Then
                Exit Sub
            End If
            iFileNum = FreeFile()
            Set oExcel = New Excel.Application
            Workbooks.Open Filename:=sFileName
            Call GetDataAssets(FirstFile, SecondFileName)
            
            Application.DisplayAlerts = False
            Workbooks(SecondFileName).Close SaveChanges:=False
            Application.DisplayAlerts = True
            
            Close iFileNum
        Next vrtSelectedItem
    Else
    End If
End With
Set fd = Nothing
Set oExcel = Nothing
End Sub


Sub GetDataAssets(ByVal FirstFile, ByVal SecondFileName)
Dim oExcel As Excel.Application
Dim oWB As Workbook
Dim i As Long
Dim k As Long
Dim Row As Long
Dim str1 As String
Dim Arr As Long

Windows(SecondFileName).Activate
Worksheets("Sheet2").Cells(1, 1).Value = "=COUNTA((Sheet1!A:A))"
Arr = Worksheets("Sheet2").Cells(1, 1).Value
Sheets("Sheet1").Select
Worksheets("Sheet1").Range(Worksheets("Sheet1").Cells(2, 1), Worksheets("Sheet1").Cells(Arr, 13)).Select
Selection.Copy

Windows(FirstFile).Activate
k = Workbooks(FirstFile).Worksheets("Assets").Cells(1, 20).Value
Worksheets("Assets").Select
Worksheets("Assets").Cells(k, 1).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
For Row = k To Workbooks(FirstFile).Worksheets("Assets").Cells(1, 20).Value - 1
    Worksheets("Assets").Cells(Row, 14).Value = GetPortfoloiCode(SecondFileName)
    Worksheets("Assets").Cells(Row, 15).Value = GetPortfoloiGroup(SecondFileName, FirstFile)
    Worksheets("Assets").Cells(Row, 16).Value = Worksheets("Portfolios").Cells(2, 7).Value
Next

End Sub


Function GetPortfoloiCode(ByVal SecondFileName) As String
Dim str1 As String
Dim l As Long

l = Len(SecondFileName)
str1 = Left(SecondFileName, l - 5)
GetPortfoloiCode = str1

End Function


Function GetPortfoloiGroup(ByVal SecondFileName, ByRef FirstFile) As String
Dim i As Long
Dim str1 As String

For i = 2 To Workbooks(FirstFile).Worksheets("Portfolios").Cells(1, 12).Value
    If Workbooks(FirstFile).Worksheets("Portfolios").Cells(i, 5).Value = SecondFileName Then
        str1 = Workbooks(FirstFile).Worksheets("Portfolios").Cells(i, 4).Value
    End If
Next i
GetPortfoloiGroup = str1

End Function


Sub TransferDataPrevDay(ByVal FirstFile)
Dim tod As Long

If Workbooks(FirstFile).Worksheets("Portfolios").Cells(2, 7).Value = Workbooks(FirstFile).Worksheets("AssetsPrev").Cells(2, 16).Value Then
    MsgBox ("Данните от предходен ден са вече прехвърляни")
    Exit Sub
End If


tod = Workbooks(FirstFile).Worksheets("AssetsPrev").Cells(1, 20).Value
Workbooks(FirstFile).Worksheets("AssetsPrev").Activate
Workbooks(FirstFile).Worksheets("AssetsPrev").Range(Worksheets("AssetsPrev").Cells(2, 1), Worksheets("AssetsPrev").Cells(tod, 16)).ClearContents

tod = Workbooks(FirstFile).Worksheets("Assets").Cells(1, 20).Value
Workbooks(FirstFile).Worksheets("Assets").Activate
Workbooks(FirstFile).Worksheets("Assets").Range(Worksheets("Assets").Cells(2, 1), Worksheets("Assets").Cells(tod, 16)).Copy

Workbooks(FirstFile).Worksheets("AssetsPrev").Activate
Worksheets("AssetsPrev").Select
Worksheets("AssetsPrev").Cells(2, 1).Select
Worksheets("AssetsPrev").Cells(2, 1).PasteSpecial

End Sub


Sub GetPortfolios()

Dim i As Long
Dim k As Long
Dim j As Long
Dim LastRowIC As Long


LastRowIC = Worksheets("Z_Grupa_Report").Cells(1, 15).Value

Worksheets("Z_Grupa_Report").Range(Worksheets("Z_Grupa_Report").Cells(7, 1), Worksheets("Z_Grupa_Report").Cells(LastRowIC, 2)).ClearContents
Worksheets("Z_Grupa_Report").Range(Worksheets("Z_Grupa_Report").Cells(7, 4), Worksheets("Z_Grupa_Report").Cells(LastRowIC, 6)).ClearContents
Worksheets("Z_Grupa_Report").Range(Worksheets("Z_Grupa_Report").Cells(7, 14), Worksheets("Z_Grupa_Report").Cells(LastRowIC, 14)).ClearContents

PortfolioNumbers = Worksheets("Portfolios").Cells(1, 12).Value

k = 7
    For j = 2 To Worksheets("Assets").Cells(1, 20).Value
        If Worksheets("Assets").Cells(j, 1).Value = "Акции" And ISINinList(Worksheets("Assets").Cells(j, 2).Value) = True Then
            Worksheets("Z_Grupa_Report").Cells(k, 1).Value = Worksheets("Assets").Cells(j, 2).Value
            Worksheets("Z_Grupa_Report").Cells(k, 2).Value = Worksheets("Assets").Cells(j, 3).Value
            If Worksheets("Assets").Cells(j, 1).Value = "Àêöèè" And Worksheets("Assets").Cells(j, 2).Value = "BG1100111111" And Worksheets("Assets").Cells(j, 14).Value = "SSSSSS" Then
                Worksheets("Z_Grupa_Report").Cells(k, 4).Value = Worksheets("Assets").Cells(j, 7).Value + Worksheets("Exception").Cells(1, 7).Value
            Else
                Worksheets("Z_Grupa_Report").Cells(k, 4).Value = Worksheets("Assets").Cells(j, 7).Value
            End If
            Worksheets("Z_Grupa_Report").Cells(k, 5).Value = Worksheets("Assets").Cells(j, 14).Value
            Worksheets("Z_Grupa_Report").Cells(k, 6).Value = Worksheets("Assets").Cells(j, 15).Value
            Worksheets("Z_Grupa_Report").Cells(k, 14).Value = Worksheets("Z_Grupa_Report").Cells(k, 4).Value - GetSharesRepo(Worksheets("Z_Grupa_Report").Cells(k, 1).Value, Worksheets("Z_Grupa_Report").Cells(k, 5).Value)
            k = k + 1
        End If
    Next j

End Sub


Function GetSharesRepo(ByVal ISIN As String, ByVal PortfolioCode As String) As Double
Dim i As Long
Dim j As Long
Dim RepoShares As Double
RepoShares = 0#

For i = 2 To Worksheets("Assets").Cells(1, 20).Value
    If Worksheets("Assets").Cells(i, 2).Value = ISIN And Worksheets("Assets").Cells(i, 14).Value = PortfolioCode And Worksheets("Assets").Cells(i, 1).Value = "Акции - Репо" And Worksheets("Assets").Cells(i, 7).Value < 0 Then
        RepoShares = RepoShares + Worksheets("Assets").Cells(i, 7).Value
    End If
Next i

RepoShares = -RepoShares
GetSharesRepo = RepoShares

End Function


Function ISINinList(ByVal ISIN As String) As Boolean
Dim i As Long
Dim bIsin As Boolean

bIsin = False

For i = 2 To Worksheets("Emission").Cells(1, 15).Value
    If Worksheets("Emission").Cells(i, 2).Value = ISIN Then
        bIsin = True
        GoTo endfunct
    End If
Next i
endfunct:
ISINinList = bIsin

End Function


