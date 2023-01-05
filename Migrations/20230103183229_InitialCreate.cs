using System;
using Microsoft.EntityFrameworkCore.Migrations;

#nullable disable

namespace ParserExcel.Migrations
{
    /// <inheritdoc />
    public partial class InitialCreate : Migration
    {
        /// <inheritdoc />
        protected override void Up(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.CreateTable(
                name: "ExcelFiles",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    FileName = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    BankName = table.Column<string>(type: "nvarchar(max)", nullable: true),
                    FileDateGenerate = table.Column<DateTime>(type: "datetime2", nullable: false),
                    FileDate = table.Column<DateTime>(type: "datetime2", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ExcelFiles", x => x.Id);
                });

            migrationBuilder.CreateTable(
                name: "ExcelFileInfos",
                columns: table => new
                {
                    Id = table.Column<int>(type: "int", nullable: false)
                        .Annotation("SqlServer:Identity", "1, 1"),
                    BalanceAccount = table.Column<int>(type: "int", nullable: false),
                    OpBalanceAsset = table.Column<double>(type: "float", nullable: false),
                    OpBalanceLiability = table.Column<double>(type: "float", nullable: false),
                    TurnoverAsset = table.Column<double>(type: "float", nullable: false),
                    TurnoverLiability = table.Column<double>(type: "float", nullable: false),
                    ClBalanceAsset = table.Column<double>(type: "float", nullable: false),
                    ClBalanceLiability = table.Column<double>(type: "float", nullable: false),
                    ExcelFileId = table.Column<int>(type: "int", nullable: false)
                },
                constraints: table =>
                {
                    table.PrimaryKey("PK_ExcelFileInfos", x => x.Id);
                    table.ForeignKey(
                        name: "FK_ExcelFileInfos_ExcelFiles_ExcelFileId",
                        column: x => x.ExcelFileId,
                        principalTable: "ExcelFiles",
                        principalColumn: "Id",
                        onDelete: ReferentialAction.Cascade);
                });

            migrationBuilder.CreateIndex(
                name: "IX_ExcelFileInfos_ExcelFileId",
                table: "ExcelFileInfos",
                column: "ExcelFileId");
        }

        /// <inheritdoc />
        protected override void Down(MigrationBuilder migrationBuilder)
        {
            migrationBuilder.DropTable(
                name: "ExcelFileInfos");

            migrationBuilder.DropTable(
                name: "ExcelFiles");
        }
    }
}
