using ParserExcel.Models;
using Microsoft.EntityFrameworkCore;

namespace Accounting.Data
{
    public class ApplicationDbContext : DbContext
    {
        public ApplicationDbContext(DbContextOptions options) : base(options)
        {
        }

        public DbSet<ExcelFile> ExcelFiles { get; set; }
        public DbSet<ExcelFileInfo> ExcelFileInfos { get; set; }

    }
}