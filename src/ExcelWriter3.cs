using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Chart;
using System.Linq;
using System.Collections.Generic;

namespace TransactionCat
{
    public static class ExcelWriter3
    {
        public static void WriteToExcel(List<Trxn> trxns)
        {
            var filename = "report3.xlsx";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                var transactionsSheet = excelPackage.Workbook.Worksheets.Add("transactions");
                transactionsSheet.Row(1).Style.Font.Bold = true;
                transactionsSheet.Cells[1, 1].Value = "date";
                transactionsSheet.Cells[1, 2].Value = "name";
                transactionsSheet.Cells[1, 3].Value = "amount";
                transactionsSheet.Cells[1, 4].Value = "catagory";
                transactionsSheet.Cells[1, 5].Value = "subcatagory";
                for (var i = 0; i < trxns.Count(); i++)
                {
                    var rowIndex = i + 2;
                    transactionsSheet.Cells[rowIndex, 1].Value = trxns[i].Date;
                    transactionsSheet.Cells[rowIndex, 1].Style.Numberformat.Format = "yyyy-mm-dd";
                    transactionsSheet.Cells[rowIndex, 2].Value = trxns[i].Name;
                    transactionsSheet.Cells[rowIndex, 3].Value = trxns[i].Amount;
                    transactionsSheet.Cells[rowIndex, 4].Value = trxns[i].Category;
                    transactionsSheet.Cells[rowIndex, 5].Value = trxns[i].SubCategory;
                }
                transactionsSheet.Cells.AutoFitColumns(1, 40);

                AddCatagoryLinesSheet(excelPackage, trxns);

                var subCatSheet = excelPackage.Workbook.Worksheets.Add("subcat");
                subCatSheet.Row(1).Style.Font.Bold = true;
                subCatSheet.Column(1).Style.Font.Bold = true;
                subCatSheet.Cells[1, 1].Value = "month";
                var subcatMonthGroups = trxns.GroupBy(z => new
                {
                    Month = new DateTime(z.Date.Year, z.Date.Month, 1),
                    z.SubCategory
                }).OrderByDescending(z => z.Key.Month).ToList();
                var months = subcatMonthGroups.Select(z => z.Key.Month).Distinct().ToList();
                var subcats = subcatMonthGroups.Select(z => z.Key.SubCategory).Distinct().ToList();
                for (var i = 0; i < subcats.Count(); i++)
                {
                    var colIndex = i + 2;
                    subCatSheet.Cells[1, colIndex].Value = subcats[i];
                }
                for (var i = 0; i < months.Count(); i++)
                {
                    var rowIndex = i + 2;
                    subCatSheet.Cells[rowIndex, 1].Value = months[i];
                    subCatSheet.Cells[rowIndex, 1].Style.Numberformat.Format = "yyyy-mm";
                    for (var j = 0; j < subcats.Count(); j++)
                    {
                        var colIndex = j + 2;
                        var matchingGroup = subcatMonthGroups.FirstOrDefault(z => months[i] == z.Key.Month && subcats[j] == z.Key.SubCategory);
                        subCatSheet.Cells[rowIndex, colIndex].Value = matchingGroup?.Sum(z => z.Amount) ?? 0;
                    }
                }
                subCatSheet.Cells.AutoFitColumns(1, 40);

                ExcelLineChart lineChart = (ExcelLineChart)subCatSheet.Drawings.AddChart("lineChart", eChartType.Line);
                for (var i = 0; i < subcats.Count(); i++)
                {
                    var colIndex = i + 2;
                    lineChart.Series.Add(ExcelRange.GetAddress(2, colIndex, 1 + months.Count(), colIndex), ExcelRange.GetAddress(2, 1, 1 + months.Count(), 1));
                    lineChart.Series[i].Header = subcats[i];
                }
                lineChart.SetPosition(months.Count() + 2, 0, 0, 0);
                lineChart.SetSize(1500, 400);

                AddSubCatagoryComparisonSheet(excelPackage, trxns);

                excelPackage.SaveAs(filename);
            }
        }

        private static void AddCatagoryLinesSheet(ExcelPackage excelPackage, List<Trxn> trxns){
            var includedTrxns = trxns.ToList();
            if (includedTrxns.First().Date.Day < 22)
            { //assume there's half a month included
                var firstOfMonth = new DateTime(includedTrxns.First().Date.Year, includedTrxns.First().Date.Month, 1);
                includedTrxns = includedTrxns.Where(z => z.Date < firstOfMonth).ToList();
            }
            var catSheet = excelPackage.Workbook.Worksheets.Add("catagory lines");
            catSheet.Row(1).Style.Font.Bold = true;
            catSheet.Column(1).Style.Font.Bold = true;
            catSheet.Cells[1, 1].Value = "month";
            var catMonthGroups = includedTrxns.GroupBy(z => new
            {
                Month = new DateTime(z.Date.Year, z.Date.Month, 1),
                z.Category
            }).OrderByDescending(z => z.Key.Month).ToList();
            var months = catMonthGroups.Select(z => z.Key.Month).Distinct().ToList();
            var cats = catMonthGroups.Select(z => z.Key.Category).Distinct().ToList();
            for (var i = 0; i < cats.Count(); i++)
            {
                var colIndex = i + 2;
                catSheet.Cells[1, colIndex].Value = cats[i];
            }
            catSheet.Cells[1, cats.Count() + 2].Value = "total";

            catSheet.Cells[2, 1].Value = "average";
            for (var i = 0; i < cats.Count(); i++)
            {
                var colIndex = i + 2;
                var sum = catMonthGroups.Where(z => z.Key.Category == cats[i]).Sum(groups => groups.Sum(z => z.Amount));
                catSheet.Cells[2, colIndex].Value = Math.Round(sum / months.Count(), 2);
            }
            catSheet.Cells[2, cats.Count() + 2].Value = Math.Round(catMonthGroups.Sum(groups => groups.Sum(z => z.Amount)) / months.Count(), 2);

            for (var i = 0; i < months.Count(); i++)
            {
                var rowIndex = i + 3;
                catSheet.Cells[rowIndex, 1].Value = months[i];
                catSheet.Cells[rowIndex, 1].Style.Numberformat.Format = "yyyy-mm";
                var sumAmount = 0m;
                for (var j = 0; j < cats.Count(); j++)
                {
                    var colIndex = j + 2;
                    var matchingGroup = catMonthGroups.FirstOrDefault(z => months[i] == z.Key.Month && cats[j] == z.Key.Category);
                    var amount = matchingGroup?.Sum(z => z.Amount) ?? 0;
                    sumAmount += amount;
                    catSheet.Cells[rowIndex, colIndex].Value = amount;
                }
                catSheet.Cells[rowIndex, cats.Count() + 2].Value = sumAmount;
            }
            catSheet.Cells.AutoFitColumns(1, 40);
            
            var start = months.Count() + 3;
            for (var i = 0; i < cats.Count(); i++)
            {
                ExcelLineChart lineChart = (ExcelLineChart)catSheet.Drawings.AddChart("lineChart-" + (i+1), eChartType.Line);
                var colIndex = i + 2;
                lineChart.Series.Add(ExcelRange.GetAddress(2, colIndex, 1 + months.Count(), colIndex), ExcelRange.GetAddress(2, 1, 1 + months.Count(), 1));
                lineChart.Series[0].Header = cats[i];
                lineChart.SetPosition(start + i*22, 0, 0, 0);
                lineChart.SetSize(1500, 400);
            }

            // ExcelPieChart pieChart = (ExcelPieChart)catSheet.Drawings.AddChart("pieChart", eChartType.Pie3D);
            // pieChart.Series.Add(ExcelRange.GetAddress(2, 2, 8, 2), ExcelRange.GetAddress(2, 1, 8, 1));
            // pieChart.DataLabel.ShowPercent = true;
            // pieChart.SetPosition(4, 0, 2, 0);
        }

        private static void AddSubCatagoryComparisonSheet(ExcelPackage excelPackage, List<Trxn> trxns)
        {
            var subCatSheet = excelPackage.Workbook.Worksheets.Add("change this month");
            var includedTrxns = trxns.ToList();
            if (includedTrxns.First().Date.Day < 22)
            { //assume there's half a month included
                var firstOfMonth = new DateTime(includedTrxns.First().Date.Year, includedTrxns.First().Date.Month, 1);
                includedTrxns = includedTrxns.Where(z => z.Date < firstOfMonth).ToList();
            }
            var prevMonthCount = 8;
            var currentMonth = new DateTime(includedTrxns.First().Date.Year, includedTrxns.First().Date.Month, 1);
            var prevMonths = Enumerable.Range(1, prevMonthCount).Reverse().Select(z => currentMonth.AddMonths(-z)).ToList();
            subCatSheet.Row(1).Style.Font.Bold = true;
            subCatSheet.Column(1).Style.Font.Bold = true;
            subCatSheet.Cells[1, 1].Value = "subcategory";
            for (var i = 0; i < prevMonthCount; i++)
            {
                subCatSheet.Cells[1, i + 2].Value = prevMonths[i].ToString("MMM yyyy");
            }
            subCatSheet.Cells[1, prevMonthCount + 2].Value = $"{prevMonths.First().ToString("MMM")} to {prevMonths.Last().ToString("MMM")} avg";
            subCatSheet.Cells[1, prevMonthCount + 3].Value = currentMonth.ToString("MMM yyyy");
            subCatSheet.Cells[1, prevMonthCount + 4].Value = "change";
            subCatSheet.Cells[1, prevMonthCount + 5].Value = "subcategory";

            var subcats = includedTrxns.OrderBy(z => z.Category).Select(z => z.SubCategory!).Distinct().ToList();
            for (var i = 0; i < subcats.Count(); i++)
            {
                AddSubCatagoryComparisonSheetHelper(
                    subCatSheet,
                    includedTrxns.Where(z => z.SubCategory == subcats[i]).ToList(),
                    i + 2,
                    subcats[i],
                    prevMonths,
                    currentMonth
                    );
            }
            AddSubCatagoryComparisonSheetHelper(
                    subCatSheet,
                    includedTrxns.ToList(),
                    subcats.Count() + 2,
                    "Total",
                    prevMonths,
                    currentMonth
                    );
            subCatSheet.Cells.AutoFitColumns(1, 40);
        }

        //made this a function to make it easier to calculate the total
        private static void AddSubCatagoryComparisonSheetHelper(
            ExcelWorksheet sheet,
            List<Trxn> someTrxns,
            int rowIndex,
            string rowName,
            List<DateTime> prevmonths,
            DateTime currentMonth)
        {
            List<decimal> prevMonthAmounts = new List<decimal>();
            for (var i = 0; i < prevmonths.Count(); i++)
            {
                var upperBound = i + 1 < prevmonths.Count() ? prevmonths[i + 1] : currentMonth;
                prevMonthAmounts.Add(someTrxns.Where(z => z.Date >= prevmonths[i] && z.Date < upperBound).Sum(z => z.Amount));
            }
            var avgAmount = Math.Round(prevMonthAmounts.Sum() / prevMonthAmounts.Count(), 2);
            var currentMonthAmount = someTrxns.Where(z => z.Date >= currentMonth).Sum(z => z.Amount);
            var change = currentMonthAmount - avgAmount;
            sheet.Cells[rowIndex, 1].Value = rowName;
            for (var i = 0; i < prevmonths.Count(); i++)
            {
                sheet.Cells[rowIndex, i + 2].Value = prevMonthAmounts[i];
            }
            sheet.Cells[rowIndex, prevmonths.Count() + 2].Value = avgAmount;
            sheet.Cells[rowIndex, prevmonths.Count() + 3].Value = currentMonthAmount;
            sheet.Cells[rowIndex, prevmonths.Count() + 4].Value = change;
            sheet.Cells[rowIndex, prevmonths.Count() + 5].Value = rowName;

            var borderColor = System.Drawing.Color.FromArgb(208, 206, 206);
            var colCount = prevmonths.Count()+5;
            if (change > 5)
            {
                var log = Math.Log2(decimal.ToDouble(change));
                var altColorVal = Math.Max(255 - (int)Math.Round((log * 3)), 0);
                var color = System.Drawing.Color.FromArgb(255, altColorVal, altColorVal);
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Fill.BackgroundColor.SetColor(color);

                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Top.Color.SetColor(borderColor);
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Bottom.Color.SetColor(borderColor);
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Left.Color.SetColor(borderColor);
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Right.Color.SetColor(borderColor);
            }
            if (change < -5)
            {
                var log = Math.Log2(decimal.ToDouble(-change));
                var altColorVal = Math.Max(255 - (int)Math.Round((log * 3)), 0);
                var color = System.Drawing.Color.FromArgb(altColorVal, 255, altColorVal);
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Fill.BackgroundColor.SetColor(color);

                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Top.Style = ExcelBorderStyle.Thin;
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Left.Style = ExcelBorderStyle.Thin;
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Right.Style = ExcelBorderStyle.Thin;
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Top.Color.SetColor(borderColor);
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Bottom.Color.SetColor(borderColor);
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Left.Color.SetColor(borderColor);
                sheet.Cells[rowIndex, 1, rowIndex, colCount].Style.Border.Right.Color.SetColor(borderColor);
            }
        }
    }
}
