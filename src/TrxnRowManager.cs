
using Microsoft.VisualBasic.FileIO;
using System.Linq;
using System.Collections.Generic;

namespace FinanceExcel
{
    public static class TrxnRowManager
    {
        public static List<TrxnRow> GetTrxnRows(List<SimpleTrxn> simpleTrxns, List<TrxnRow> existingTrxnRows, List<CategoryRule> categoryRules)
        {
            var now = DateTime.Now;
            var newSimpletrxns = simpleTrxns.Where(simpleTrxn => !existingTrxnRows.Any(existingTrxnRow => 
                existingTrxnRow.Originaltrxn != null
                && existingTrxnRow.Originaltrxn.Amount == simpleTrxn.Amount
                && existingTrxnRow.Originaltrxn.Date == simpleTrxn.Date
                && existingTrxnRow.Originaltrxn.Name == simpleTrxn.Name
            ));


            var newTrxnRows = newSimpletrxns.Select(newSimpletrxn => {
                var name = newSimpletrxn.Name.ToLower();
                var rule = categoryRules.FirstOrDefault(rule => rule.Contains.Any(z => z.Length > 0 && name.Contains(z)) && rule.NotContains.All(z => z.Length == 0 || !name.Contains(z)));
                var category = rule != null ? rule.Category : "uncatagorized";
                var subCategory = rule != null ? (!string.IsNullOrEmpty(rule.Subcategory) ? rule.Subcategory : rule.Category) : "uncatagorized";
                var trxn = new Trxn(newSimpletrxn.Name, newSimpletrxn.Date, newSimpletrxn.Amount, category, subCategory);
                return new TrxnRow(trxn, "", "", now, trxn);
            });

            var updatedExistingTrxnRows = existingTrxnRows.Select(existingTrxnRow => {
                if (existingTrxnRow.Originaltrxn == null || existingTrxnRow.Originaltrxn.Category != existingTrxnRow.Trxn.Category || existingTrxnRow.Originaltrxn.SubCategory != existingTrxnRow.Trxn.SubCategory){
                    return existingTrxnRow with { Modifed = "yes" };
                }
                var modified = "";
                if (existingTrxnRow.Originaltrxn.Name != existingTrxnRow.Trxn.Name || existingTrxnRow.Originaltrxn.Amount != existingTrxnRow.Trxn.Amount || existingTrxnRow.Originaltrxn.Date != existingTrxnRow.Trxn.Date){
                    modified = "yes";
                }
                var name = existingTrxnRow.Trxn.Name.ToLower();
                var rule = categoryRules.FirstOrDefault(rule => rule.Contains.Any(z => z.Length > 0 && name.Contains(z)) && rule.NotContains.All(z => z.Length == 0 || !name.Contains(z)));
                var category = rule != null ? rule.Category : "uncatagorized";
                var subCategory = rule != null ? (!string.IsNullOrEmpty(rule.Subcategory) ? rule.Subcategory : rule.Category) : "uncatagorized";
                var trxn = new Trxn(existingTrxnRow.Trxn.Name, existingTrxnRow.Trxn.Date, existingTrxnRow.Trxn.Amount, category, subCategory);
                var originalTrxn = new Trxn(existingTrxnRow.Originaltrxn.Name, existingTrxnRow.Originaltrxn.Date, existingTrxnRow.Originaltrxn.Amount, category, subCategory);
                return new TrxnRow(trxn, existingTrxnRow.Notes, modified, existingTrxnRow.insertDate ?? now, originalTrxn);
            });
            var trxnRows = newTrxnRows.Concat(updatedExistingTrxnRows).OrderByDescending(z => z.Trxn.Date).ToList();
            return trxnRows;
        }
    }
}
