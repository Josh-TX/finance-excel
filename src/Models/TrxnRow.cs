namespace FinanceExcel{
    public record TrxnRow(
        Trxn Trxn,
        string Notes,
        string Modifed,
        DateTime? insertDate,
        Trxn? Originaltrxn
        );
}