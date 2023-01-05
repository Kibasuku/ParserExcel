using System.ComponentModel.DataAnnotations;

namespace ParserExcel.Models
{
    public class ExcelFile
    {
        [Key]
        public int Id { get; set; }
        public string? FileName { get; set; }
        public string? BankName { get; set; }
        public DateTime FileDateGenerate { get; set; }
        public DateTime FileDate { get; set; } 


    }
}
