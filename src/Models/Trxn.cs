namespace FinanceExcel{

    public record Trxn(
        string Name, 
        DateTime Date, 
        decimal Amount,
        string Category,
        string SubCategory
        ) : SimpleTrxn(Name, Date, Amount);

}