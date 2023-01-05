using Accounting.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using ParserExcel.Models;
using System.Diagnostics;

namespace ParserExcel.Controllers
{
    public class HomeController : Controller
    {
        private readonly ApplicationDbContext _dbContext;
        private readonly IWebHostEnvironment _appEnvironment;
        public HomeController(ApplicationDbContext applicationDbContext, IWebHostEnvironment appEnvironment)
        {
            _dbContext = applicationDbContext;
            _appEnvironment = appEnvironment;
        }

        public async Task<IActionResult> Index()
        {
            return View(await _dbContext.ExcelFiles.ToListAsync());
        }

        // добавление файла
        [HttpPost]
        public async Task<IActionResult> AddFile(IFormFile uploadedFile)
        {
            try
            {
                // загрузка файла на сервер
                if (uploadedFile != null)
                {
                    string path = @"\Files\" + uploadedFile.FileName;
                    using (var fileStream = new FileStream(_appEnvironment.WebRootPath + path, FileMode.Create))
                    {
                        await uploadedFile.CopyToAsync(fileStream);
                    }

                    // парсинг excel файла
                    using (SpreadsheetDocument spreadsheetDocument =
                            SpreadsheetDocument.Open(_appEnvironment.WebRootPath + path, false))
                    {

                        var cellExcelFileList = GetExcelFileData(spreadsheetDocument);

                        var excelFile = AddExcelFile(cellExcelFileList, uploadedFile.FileName);
                        _dbContext.ExcelFiles.Add(excelFile);

                        // парсинг информации для ExcelFileInfo
                        var excelFileInfosList = GetExcelFileInfoData(spreadsheetDocument, excelFile);
                        _dbContext.ExcelFileInfos.AddRange(excelFileInfosList);
                        await _dbContext.SaveChangesAsync();

                    }

                    // удаление файла с сервера
                    new FileInfo(_appEnvironment.WebRootPath + path).Delete();
                  
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return StatusCode(500);
            }
            return RedirectToAction("Index");
        }

        // просмотр данных отдельного документа
        public async Task<IActionResult> Review(int? id)
        {
            try
            {
                if (id != null)
                {
                    var excelFileInfo = await _dbContext.ExcelFileInfos.Where(f => f.ExcelFileId == id).ToListAsync();
                    if (excelFileInfo != null)
                    {
                        return View(excelFileInfo);
                    }
                }
                
            }  
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
                return StatusCode(500);
            }
            return NotFound();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        } 

        // получение значения из ячейки
        private static string GetValueFromCell(SpreadsheetDocument doc, Cell cell)
        {
            if(doc.WorkbookPart!=null)
            {
                if (cell.CellValue != null && doc.WorkbookPart.SharedStringTablePart != null)
                {
                    string value = cell.CellValue.InnerText;
                    if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    {
                        return doc.WorkbookPart.SharedStringTablePart.SharedStringTable.ChildElements.GetItem(int.Parse(value)).InnerText;
                    }
                    if (value.Contains('.'))
                    {
                        value = value.Replace('.', ',');
                    }
                    return value;
                }
            }  
            return "";
        }

        // сборка ExcelFileInfo
        private static ExcelFileInfo AddExcelFileInfo(List<string> cellExcelFileInfoList, ExcelFile excelFile)
        {
            
            var excelFileInfo = new ExcelFileInfo()
            {
                BalanceAccount = int.Parse(cellExcelFileInfoList[0]),
                OpBalanceAsset = double.Parse(cellExcelFileInfoList[1]),
                OpBalanceLiability = double.Parse(cellExcelFileInfoList[2]),
                TurnoverAsset = double.Parse(cellExcelFileInfoList[3]),
                TurnoverLiability = double.Parse(cellExcelFileInfoList[4]),
                ClBalanceAsset = double.Parse(cellExcelFileInfoList[5]),
                ClBalanceLiability = double.Parse(cellExcelFileInfoList[6]),
                ExcelFile = excelFile
            };
            return excelFileInfo;
        }

        // сборка ExcelFile
        private static ExcelFile AddExcelFile(List<string> cellExcelFileList, string fileName)
        {
            var excelFile = new ExcelFile()
            {
                FileName = fileName,
                BankName = cellExcelFileList[0],
                FileDateGenerate = DateTime.Now,
                FileDate = DateTime.Parse(cellExcelFileList[1])
            };
            return excelFile;
        }

        // получение информации для ExcelFile
        private static List<string> GetExcelFileData(SpreadsheetDocument spreadsheetDocument)
        {
            
            var cellExcelFileList = new List<string>();
            if (spreadsheetDocument.WorkbookPart != null)
            {
                var sheetData = spreadsheetDocument.WorkbookPart.WorksheetParts.First().Worksheet.Elements<SheetData>().First();
                var rowList = sheetData.Elements<Row>().Where(r => r.RowIndex == 1 || r.RowIndex == 6).ToList<Row>();
                foreach (Row r in rowList)
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        if (c.CellReference != null)
                        {
                            if (c.CellReference.ToString() == "A1" || c.CellReference.ToString() == "A6")
                            {
                                cellExcelFileList.Add(GetValueFromCell(spreadsheetDocument, c));
                            }
                        }

                    }
                }
            }
                
            return cellExcelFileList;
        }

        // получение списка ExcelFileInfo
        private static List<ExcelFileInfo> GetExcelFileInfoData(SpreadsheetDocument spreadsheetDocument, ExcelFile excelFile)
        {
            if (spreadsheetDocument.WorkbookPart != null)
            {
                var sheetData = spreadsheetDocument.WorkbookPart.WorksheetParts.First().Worksheet.Elements<SheetData>().First();
                var excelFileInfosList = new List<ExcelFileInfo>();
                foreach (Row r in sheetData.Elements<Row>())
                {
                    var rowIndex = 0;
                    var cellExcelFileInfoList = new List<string>();

                    foreach (Cell c in r.Elements<Cell>())
                    {
                        if (c.CellValue != null && c.CellReference != null)
                        {

                            if (c.CellReference.ToString().Contains('A'))
                            {
                                if (int.Parse(c.CellValue.Text) >= 1000)
                                {
                                    rowIndex = (int)r.RowIndex.Value;
                                }

                            }
                            if (rowIndex != 0)
                            {
                                cellExcelFileInfoList.Add(GetValueFromCell(spreadsheetDocument, c));
                            }
                        }

                    }
                    if (rowIndex != 0)
                    {
                        excelFileInfosList.Add(AddExcelFileInfo(cellExcelFileInfoList, excelFile));
                    }

                }
                return excelFileInfosList;
            }
            return new List<ExcelFileInfo>();

        }

    }
}