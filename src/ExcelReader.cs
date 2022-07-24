using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Chart;
using System.Linq;
using System.Collections.Generic;

namespace FinanceExcel
{
    public static class ExcelReader
    {
        public static List<TrxnRow> ReadFromExcel()
        {
            var filename = "report.xlsx";
            List<TrxnRow> trxnRows = new List<TrxnRow>();
            if (!System.IO.File.Exists(filename))
            {
                return trxnRows;
            }
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage(filename))
            {
                var transactionsSheet = excelPackage.Workbook.Worksheets.First();
                for (var i = 1; i < transactionsSheet.Dimension.Rows; i++)
                {
                    var rowIndex = i + 1;

                    var trxn = TryGetTrxn(
                        transactionsSheet.Cells[rowIndex, 1].Value?.ToString(), //oadate
                        transactionsSheet.Cells[rowIndex, 2].Value?.ToString(), //name
                        transactionsSheet.Cells[rowIndex, 3].Value?.ToString(), //amount
                        transactionsSheet.Cells[rowIndex, 4].Value?.ToString(), //category
                        transactionsSheet.Cells[rowIndex, 5].Value?.ToString()  //subcategory
                    );
                    if (trxn == null)
                    {
                        continue;
                    }

                    var notes = transactionsSheet.Cells[rowIndex, 6].Value as string;
                    var modified = transactionsSheet.Cells[rowIndex, 7].Value as string;
                    var insertOADate = transactionsSheet.Cells[rowIndex, 8].Value as double?;
                    DateTime? insertDate = insertOADate.HasValue ? DateTime.FromOADate(insertOADate.Value) : null;

                    var originalTrxn = TryGetTrxn(
                        transactionsSheet.Cells[rowIndex, 9].Value?.ToString(),  //oadate
                        transactionsSheet.Cells[rowIndex, 10].Value?.ToString(), //name
                        transactionsSheet.Cells[rowIndex, 11].Value?.ToString(), //amount
                        transactionsSheet.Cells[rowIndex, 12].Value?.ToString(), //category
                        transactionsSheet.Cells[rowIndex, 13].Value?.ToString()  //subcategory
                    );
                    trxnRows.Add(new TrxnRow(trxn, notes ?? "", modified ?? "", insertDate, originalTrxn));
                }
            }
            return trxnRows;
        }

        private static DateTime tryGetDate(object dateObj){
            if (dateObj is DateTime){

            }
            return DateTime.Now;
        }

        private static Trxn? TryGetTrxn(string? dateStr, string? name, string? amountStr, string? category, string? subcategory)
        {
            if (string.IsNullOrEmpty(dateStr) ||  string.IsNullOrEmpty(name) || string.IsNullOrEmpty(amountStr))
            {
                return null;
            }
            var amount = decimal.Parse(amountStr);
            category = !string.IsNullOrEmpty(category) ? category : "uncatagorized";
            subcategory = !string.IsNullOrEmpty(subcategory) ? subcategory : "uncatagorized";
            DateTime date;
            var isOADate = int.TryParse(dateStr, out var dateInt);
            if (isOADate && dateInt != 0){
                date = DateTime.FromOADate(dateInt);
                if (date.Year < 1900 || date.Year > 2100)
                {
                    return null;
                }
            } else {
                var isDateStr = DateTime.TryParse(dateStr, out date);
                if (!isDateStr || date.Year < 1900 || date.Year > 2100){
                    return null;
                }
            }
            return new Trxn(name, date, amount, category, subcategory);
        }
    }
}
