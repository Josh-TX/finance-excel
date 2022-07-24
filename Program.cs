using FinanceExcel;

var simpleTrxns = TrxnReader.GetSimpleTrxns();
var existingTrxnRows = ExcelReader.ReadFromExcel();
var categoryRules = Categorizor.GetCategoryRules();
var trxnRows = TrxnRowManager.GetTrxnRows(simpleTrxns, existingTrxnRows, categoryRules);
ExcelWriter.WriteToExcel(trxnRows);

#if !DEBUG
    Console.Write($"{Environment.NewLine}Press any key to exit...");
    Console.ReadKey(true);
#endif
