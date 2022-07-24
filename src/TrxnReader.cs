
using Microsoft.VisualBasic.FileIO;
using System.Linq;
using System.Collections.Generic;

namespace FinanceExcel
{
    public static class TrxnReader
    {
        public static List<SimpleTrxn> GetSimpleTrxns()
        {
            var folderExists = System.IO.Directory.Exists("transactions");
            if (!folderExists)
            {
                Console.WriteLine("missing transactions folder");
                return new List<SimpleTrxn>();
            }
            var filenames = System.IO.Directory.GetFiles("transactions");
            if (filenames.Count() == 0){
                Console.WriteLine("no files in transactions folder");
                return new List<SimpleTrxn>();
            }
            var trxns = new List<SimpleTrxn>();
            foreach (var filename in filenames)
            {
                Console.WriteLine($"Loading Transactions from " + filename);
                var rowNum = 1;
                var dateIndex = 0;
                var nameIndex = 0;
                var amountIndex = 0;
                using (var parser = new TextFieldParser(filename))
                {
                    parser.Delimiters = new[] { ",", "\t" };
                    while (!parser.EndOfData)
                    {
                        try
                        {
                            var currentRow = parser.ReadFields();
                            if (currentRow != null)
                            {
                                if (rowNum == 1)
                                {
                                    var headers = currentRow.ToList();
                                    dateIndex = headers.FindIndex(z => z.ToLower() == "date");
                                    nameIndex = headers.FindIndex(z => z.ToLower() == "description");
                                    amountIndex = headers.FindIndex(z => z.ToLower() == "amount" || z.ToLower() == "debit");
                                    if (dateIndex == -1)
                                    {
                                        Console.WriteLine($"First row should be headers. Date header not found");
                                        break;
                                    }
                                    if (nameIndex == -1)
                                    {
                                        Console.WriteLine($"First row should be headers. Name header not found");
                                        break;
                                    }
                                    if (amountIndex == -1)
                                    {
                                        Console.WriteLine($"First row should be headers. Amount nor debit header not found");
                                        break;
                                    }
                                }
                                else
                                {
                                    var success = decimal.TryParse(currentRow[amountIndex].Replace("$", ""), out var amount);
                                    if(!success){
                                        Console.WriteLine($"Error parsing amount '{currentRow[amountIndex]}' in row {rowNum}");
                                        continue;
                                    }
                                    var date = DateTime.Parse(currentRow[dateIndex]);
                                    if (amount != 0){
                                        trxns.Add(new SimpleTrxn(currentRow[nameIndex], date, amount));
                                    }
                                }
                            }
                        }
                        catch (Exception)
                        {

                        }
                        rowNum++;
                    }
                }
            }
            return trxns;
        }
    }
}
