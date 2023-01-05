namespace ParserExcel.Models
{
    public class ExcelFileInfo
    {
        public int Id { get; set; }
        public int BalanceAccount { get; set; }
        public double OpBalanceAsset { get; set; }
        public double OpBalanceLiability { get; set; }
        public double TurnoverAsset { get; set; }
        public double TurnoverLiability { get; set; }
        public double ClBalanceAsset { get; set; }
        public double ClBalanceLiability { get; set; }
        public int ExcelFileId { get; set; }
        public ExcelFile? ExcelFile { get; set; }
    }
}
