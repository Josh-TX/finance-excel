
using Microsoft.VisualBasic.FileIO;
using System.Linq;
using System.Collections.Generic;

namespace FinanceExcel
{
    public static class Categorizor
    {
        public static List<CategoryRule> GetCategoryRules()
        {
            var filename = "categories.csv";
            if (!System.IO.File.Exists(filename))
            {
                Console.WriteLine("missing " + filename);
                return new List<CategoryRule>();
            }
            var rules = new List<CategoryRule>();
            Console.WriteLine($"Reading categories from " + filename);
            using (var parser = new TextFieldParser(filename))
            {
                var rowNum = 1;
                List<int>? containsIndexes = null;
                List<int>? notContainsIndexes = null;
                int nameIndex = 0;
                int catIndex = 0;
                int subCatIndex = 0;
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
                                containsIndexes = headers.Select((z, i) => !z.ToLower().Contains("not") && z.ToLower().Contains("contain") ? i : -1).Where(i => i != -1).ToList();
                                notContainsIndexes = headers.Select((z, i) => z.ToLower().Contains("not") && z.ToLower().Contains("contain") ? i : -1).Where(i => i != -1).ToList();
                                nameIndex = headers.FindIndex(z => z.ToLower() == "name");
                                catIndex = headers.FindIndex(z => z.ToLower() == "category");
                                subCatIndex = headers.FindIndex(z => z.ToLower() == "subcategory");
                                if (containsIndexes.Count() == 0)
                                {
                                    Console.WriteLine($"First row should be headers. Contains header not found");
                                    break;
                                }
                                if (nameIndex == -1)
                                {
                                    Console.WriteLine($"First row should be headers. Name header not found");
                                    break;
                                }
                                if (catIndex == -1)
                                {
                                    Console.WriteLine($"First row should be headers. category header not found");
                                    break;
                                }
                                if (subCatIndex == -1)
                                {
                                    Console.WriteLine($"First row should be headers. subcategory header not found");
                                    break;
                                }
                            }
                            else
                            {
                                var contains = containsIndexes!.Select(i => currentRow[i].ToLower()).ToList();
                                var notContains = notContainsIndexes!.Select(i => currentRow[i].ToLower()).ToList();
                                var rule = new CategoryRule(
                                    contains,
                                    notContains,
                                    currentRow[nameIndex],
                                    currentRow[catIndex],
                                    currentRow[subCatIndex]
                                    );
                                rules.Add(rule);
                            }
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"Error occurred parsing {filename}: {e.Message}");
                        break;
                    }
                    rowNum++;
                }
            }
            return rules;
        }
    }
}
