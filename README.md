Sub GenerateFundingReportAndEmail()

    Dim fxCsvPath As String
    Dim settlementCsvPath As String

    fxCsvPath = "C:\FundingAuto\FX.csv"
    settlementCsvPath = "C:\FundingAuto\Settlement.csv"

    Call ClearTempTables
    Call ImportFXCsv(fxCsvPath)
    Call ImportSettlementCsv(settlementCsvPath)

    MsgBox "CSV files imported."

End Sub
