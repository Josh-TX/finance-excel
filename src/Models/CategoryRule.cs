namespace FinanceExcel{
    public record CategoryRule(
        List<string> Contains, 
        List<string> NotContains,
        string Name,
        string Category,
        string Subcategory
        );
}