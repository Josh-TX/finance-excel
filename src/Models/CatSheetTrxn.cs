namespace FinanceExcel{

    public record CatSheetTrxn(
        string Name, 
        DateTime Date, 
        decimal Amount,
        string Category,
        string? CategoryGroup
        ) : SimpleTrxn(Name, Date, Amount);

}