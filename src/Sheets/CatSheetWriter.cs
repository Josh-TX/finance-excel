using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Chart;
using System.Linq;
using System.Collections.Generic;

namespace FinanceExcel
{
    public static class CatSheetWriter
    {
        public static void AddCategorySheet(ExcelWorksheet catSheet, List<CatSheetTrxn> trxns){
            var includedTrxns = trxns.ToList();
            if (includedTrxns.First().Date.Day < 22)
            { //if the latest trxn was before the 22nd, then the latest month will be the month PRIOR to the latest trxn. 
                var firstOfMonth = new DateTime(includedTrxns.First().Date.Year, includedTrxns.First().Date.Month, 1);
                includedTrxns = includedTrxns.Where(z => z.Date < firstOfMonth).ToList();
            }
            catSheet.Column(1).Style.Font.Bold = true;
            var catMonthGroups = includedTrxns.GroupBy(z => new
            {
                Month = new DateTime(z.Date.Year, z.Date.Month, 1),
                z.Category
            }).OrderByDescending(z => z.Key.Month).ToList();
            var monthGroups = includedTrxns.GroupBy(z => new DateTime(z.Date.Year, z.Date.Month, 1)).OrderByDescending(z => z.Key.Month).ToList();
            var months = catMonthGroups.Select(z => z.Key.Month).Distinct().ToList();
            var cats = catMonthGroups.Select(z => z.Key.Category).Distinct().ToList();

            //deal with categoryGroups
            var catsToCatGroup = new Dictionary<string, string>();
            foreach (var trxn in trxns.Where(z => z.CategoryGroup != null)){
                catsToCatGroup[trxn.Category] = trxn.CategoryGroup!;
            }
            var hasCG = catsToCatGroup.Count() > 0 ? 1 : 0;
            if (hasCG == 1){
                cats.Sort((z1, z2) => {
                    if (catsToCatGroup.ContainsKey(z1)){
                        if (catsToCatGroup.ContainsKey(z2)){
                            return catsToCatGroup[z1].CompareTo(catsToCatGroup[z2]);
                        }
                        return -1;
                    }
                    if (catsToCatGroup.ContainsKey(z2)){
                        return 1;
                    }
                    return 0;
                });
                //row 1: CategoryGroup headers
                catSheet.Row(1).Style.Font.Bold = true;
                for (var i = 0; i < cats.Count(); i++)
                {
                    var endIndex = i;
                    catsToCatGroup.TryGetValue(cats[i], out string? catGroup);
                    if (catGroup != null){
                        while (endIndex + 1 < cats.Count() && catsToCatGroup.ContainsKey(cats[endIndex + 1]) && catsToCatGroup[cats[endIndex + 1]] == catGroup){
                            endIndex++;
                        }
                        catSheet.Cells[1, i + 2, 1, endIndex + 2].Merge = true;
                        catSheet.Cells[1, i + 2, 1, endIndex + 2].Value = catGroup;
                        i = endIndex;
                    }
                }
            } else {
                cats.Sort((z1, z2) => {
                    return z1.CompareTo(z2);
                });
            }

            //create amounts matrix
            decimal[,] matrix = new decimal[months.Count,cats.Count];
            for (var i = 0; i < months.Count(); i++)
            {
                for (var j = 0; j < cats.Count(); j++)
                {
                    var matchingGroup = catMonthGroups.FirstOrDefault(z => months[i] == z.Key.Month && cats[j] == z.Key.Category);
                    matrix[i,j] = matchingGroup?.Sum(z => z.Amount) ?? 0;
                }
            }
            var averagesAndSds = cats.Select((cat, i) => {
                var amounts = Enumerable.Range(0, matrix.GetLength(0)).Select(rowIndex => matrix[rowIndex, i]).ToList();
                return ChartHelpers.GetAvgAndSD(amounts);
            }).ToList();
            var monthTotalAmounts = Enumerable.Range(0, matrix.GetLength(0))
                .Select(rowIndex => Enumerable.Range(0, matrix.GetLength(1)).Select(colIndex => matrix[rowIndex, colIndex]).Sum()).ToList();
            var totalAverageAndSd = ChartHelpers.GetAvgAndSD(monthTotalAmounts);


            //row 1 (or 2): column headers
         catSheet.Row(1 + hasCG).Style.Font.Bold = true;
            catSheet.Cells[1 + hasCG, 1].Value = "month";
            for (var i = 0; i < cats.Count(); i++)
            {
                var colIndex = i + 2;
                catSheet.Cells[1 + hasCG, colIndex].Value = cats[i];
            }
            catSheet.Cells[1 + hasCG, cats.Count() + 2].Value = "total";

            //row 2 (or 3): average
            catSheet.Cells[2 + hasCG, 1].Value = "average";
            for (var i = 0; i < cats.Count(); i++)
            {
                var colIndex = i + 2;
                catSheet.Cells[2 + hasCG, colIndex].Value = averagesAndSds[i].Item1;
            }
            catSheet.Cells[2 + hasCG, cats.Count() + 2].Value = totalAverageAndSd.Item1;

            //row 3 (or 4): std dev
            catSheet.Cells[3 + hasCG, 1].Value = "std dev";
            for (var i = 0; i < cats.Count(); i++)
            {
                var colIndex = i + 2;
                catSheet.Cells[3 + hasCG, colIndex].Value = averagesAndSds[i].Item2;
            }
            catSheet.Cells[3 + hasCG, cats.Count() + 2].Value = totalAverageAndSd.Item2;

            //remaining rows: monthly amounts
            for (var i = 0; i < months.Count(); i++)
            {
                var rowIndex = i + 4 + hasCG;
                catSheet.Cells[rowIndex, 1].Value = months[i];
                catSheet.Cells[rowIndex, 1].Style.Numberformat.Format = "yyyy-mm";
                for (var j = 0; j < cats.Count(); j++)
                {
                    var colIndex = j + 2;
                    catSheet.Cells[rowIndex, colIndex].Value = matrix[i,j];
                    ChartHelpers.DrawRedGreenBackground(catSheet, matrix[i,j], averagesAndSds[j], rowIndex, colIndex);
                }
                catSheet.Cells[rowIndex, cats.Count() + 2].Value = monthTotalAmounts[i];
                ChartHelpers.DrawRedGreenBackground(catSheet, monthTotalAmounts[i], totalAverageAndSd, rowIndex, cats.Count() + 2);
            }
            catSheet.Cells.AutoFitColumns(1, 40);
            ChartHelpers.DrawDefaultBorders(catSheet, 4, 2,  months.Count() + 3, cats.Count() + 2);
            
            


            // var start = months.Count() + 3;
            // for (var i = 0; i < cats.Count(); i++)
            // {
            //     ExcelLineChart lineChart = (ExcelLineChart)catSheet.Drawings.AddChart("lineChart-" + (i+1), eChartType.Line);
            //     var colIndex = i + 2;
            //     lineChart.Series.Add(ExcelRange.GetAddress(2, colIndex, 1 + months.Count(), colIndex), ExcelRange.GetAddress(2, 1, 1 + months.Count(), 1));
            //     lineChart.Series[0].Header = cats[i];
            //     lineChart.SetPosition(start + i*22, 0, 0, 0);
            //     lineChart.SetSize(1500, 400);
            // }
        }
    }
}
