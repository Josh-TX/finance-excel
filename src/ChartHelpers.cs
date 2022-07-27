using OfficeOpenXml;
using OfficeOpenXml.Table.PivotTable;
using OfficeOpenXml.Style;
using OfficeOpenXml.Drawing.Chart;
using System.Linq;
using System.Collections.Generic;
using System.Linq;

namespace FinanceExcel
{
    public static class ChartHelpers
    {
        private static System.Drawing.Color borderColor = System.Drawing.Color.FromArgb(208, 206, 206);

        public static void DrawBackgroundAndBorders(ExcelWorksheet sheet, System.Drawing.Color color, int fromRow, int fromCol, int toRow, int toCol)
        {
            sheet.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[fromRow, fromCol, toRow, toCol].Style.Fill.BackgroundColor.SetColor(color);
            DrawDefaultBorders(sheet, fromRow, fromCol, toRow, toCol);
        }

        public static void DrawRedGreenBackground(ExcelWorksheet sheet, decimal val, (decimal, decimal) avgAndSD, int row, int col)
        {
            if (val == avgAndSD.Item1){
                return;
            }
            var z = (val - avgAndSD.Item1) / avgAndSD.Item2;
            var SqAbsZ = Sqrt(Math.Abs(z));
            var colorInt = Convert.ToInt32(Math.Round(250 - (SqAbsZ * 40)));
            System.Drawing.Color color;
            if (z > 0){
                color = System.Drawing.Color.FromArgb(255,colorInt,colorInt);
            } else {
                color = System.Drawing.Color.FromArgb(colorInt,255,colorInt);
            }
            sheet.Cells[row, col, row, col].Style.Fill.PatternType = ExcelFillStyle.Solid;
            sheet.Cells[row, col, row, col].Style.Fill.BackgroundColor.SetColor(color);
        }

        public static void DrawDefaultBorders(ExcelWorksheet sheet, int fromRow, int fromCol, int toRow, int toCol)
        {
            sheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Style = ExcelBorderStyle.Thin;
            sheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            sheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Style = ExcelBorderStyle.Thin;
            sheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Style = ExcelBorderStyle.Thin;
            sheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Top.Color.SetColor(borderColor);
            sheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Bottom.Color.SetColor(borderColor);
            sheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Left.Color.SetColor(borderColor);
            sheet.Cells[fromRow, fromCol, toRow, toCol].Style.Border.Right.Color.SetColor(borderColor);
        }

        public static (decimal, decimal) GetAvgAndSD(List<decimal> amounts)
        {
            var average = amounts.Average();
            var sumOfSquaresOfDifferences = amounts.Sum(val => (val - average) * (val - average));
            var variance = sumOfSquaresOfDifferences / amounts.Count;
            var sd = Sqrt(variance);
            return (average, sd);
        }

        //copied from https://stackoverflow.com/questions/4124189/performing-math-operations-on-decimal-datatype-in-c
        public static decimal Sqrt(decimal x, decimal epsilon = 0.0M)
        {
            decimal current = (decimal)Math.Sqrt((double)x), previous;
            do
            {
                previous = current;
                if (previous == 0.0M) return 0;
                current = (previous + x / previous) / 2;
            }
            while (Math.Abs(previous - current) > epsilon);
            return current;
        }
    }
}
