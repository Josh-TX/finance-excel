using TransactionCat;

var trxns = TrxnReader.GetTrxns();
var rules = Categorizor.GetCategoryRules();
Categorizor.ApplyCategoryRules(trxns, rules);
ExcelWriter3.WriteToExcel(trxns);

#if !DEBUG
    Console.Write($"{Environment.NewLine}Press any key to exit...");
    Console.ReadKey(true);
#endif
