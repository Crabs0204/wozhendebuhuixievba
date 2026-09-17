Sub GenerateFundingReportAndEmail()

    Dim fxCsvPath As String
    Dim settlementCsvPath As String
    Dim templatePath As String
    Dim outputPath As String

    Dim batchNo As String

    Dim xlApp As Object
    Dim xlWb As Object
    Dim xlWs As Object

    fxCsvPath = "C:\FundingAuto\FX.csv"
    settlementCsvPath = "C:\FundingAuto\Settlement.csv"
    templatePath = "C:\FundingAuto\Funding Template.xlsx"
    outputPath = "C:\FundingAuto\Funding Report_" & Format(Date, "yyyymmdd") & ".xlsx"

    batchNo = InputBox("Enter batch number: 1, 2 or 3")

    If batchNo <> "1" And batchNo <> "2" And batchNo <> "3" Then
        MsgBox "Invalid batch number"
        Exit Sub
    End If

    Call ClearTempTables

    Call ImportFXCsv(fxCsvPath)

    Call ImportSettlementCsv(settlementCsvPath)

    Set xlApp = CreateObject("Excel.Application")
    Set xlWb = xlApp.Workbooks.Open(templatePath)
    Set xlWs = xlWb.Sheets("Report")

    Call WriteFundingNumbersToExcel(xlWs, batchNo)

    xlWb.SaveAs outputPath
    xlWb.Close
    xlApp.Quit

    Set xlWs = Nothing
    Set xlWb = Nothing
    Set xlApp = Nothing

    Call CreateFundingEmail(outputPath)

    MsgBox "Funding report and email have been generated."

End Sub
