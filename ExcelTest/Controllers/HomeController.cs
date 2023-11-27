using ExcelTest.Models;
using Microsoft.AspNetCore.Mvc;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data;
using Microsoft.Data.SqlClient;
using System.Diagnostics;
using Microsoft.EntityFrameworkCore;
using static Org.BouncyCastle.Math.EC.ECCurve;
using Dapper;
using static Microsoft.EntityFrameworkCore.DbLoggerCategory.Database;
using NPOI.OpenXmlFormats.Spreadsheet;
using Microsoft.AspNetCore.Mvc.TagHelpers;
using MathNet.Numerics.LinearAlgebra.Factorization;
using Org.BouncyCastle.Utilities.Zlib;
using System.Drawing;
using System.Text.RegularExpressions;
using NPOI.POIFS.Storage;
using System.Xml;


namespace ExcelTest.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        private readonly IConfiguration _config;
        public HomeController(ILogger<HomeController> logger, IConfiguration config)
        {
            _logger = logger;
            _config = config;
        }

        public IDbConnection Connection
        {
            get
            {
                return new SqlConnection(_config.GetConnectionString("DefaultConnection"));
            }
        }

        public IActionResult Index(Needed model, string button)
        {
            model.label = "Просмотр файлов";
            using (IDbConnection db = Connection)
            {
                IEnumerable<int> count = db.Query<int>("SELECT COUNT(*) AS TABLE_COUNT FROM INFORMATION_SCHEMA.TABLES");
                if (count.First() == 0)
                {
                    model.label = "Загрузите файлы, файлов нет!";
                }
                else
                {
                    List<string> NamesTables = db.Query<string>("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE'").ToList();
                    List<string> Years = new List<string>();
                    foreach (string name in NamesTables)
                    {
                        Years.Add(name.Substring(name.Length - 4));
                        
                    }
                    Years = Years.Distinct().ToList();
                    Years.Sort();
                    model.buttonYears = Years;

                    if (button != null)
                    {
                        List<string> month_sorted = new List<string>();
                            foreach (string month in NamesTables)
                            {
                                if (month.Contains(button))
                                {
                                    month_sorted.Add(month);
                                    try 
                                    { 
                                        List<Weather_data> result = db.Query<Weather_data>("Select * FROM " + button).ToList();
                                        model.weather = result;
                                    }
                                    catch { }
                                    
                            }
                            }
                        model.buttonMonth = month_sorted;
                    }

                    
                    
                }

                


            }
                return View(model);
        }
                             

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public IActionResult UploadExcel(Models.Label model)
        {
            model.Text = "";
            return View(model);
        }

        [HttpPost]
        public async Task<IActionResult> UploadExcel(List<IFormFile> file, Models.Label model) // Загрузка файлов
        {

            using (IDbConnection db = Connection)
            {
                // очистка БД
                db.Query("declare @sql nvarchar(max); set @sql = N''; Select @sql += 'Drop table' + QUOTENAME(table_name) + ';' from INFORMATION_SCHEMA.TABLES where TABLE_TYPE = 'Base table'; Exec sp_executesql @sql;");
            }

            foreach (IFormFile excel in file)
            {
                try
                {
                    if (excel != null && excel.Length > 0)
                    {
                        var uploadsFolder = $"{Directory.GetCurrentDirectory()}\\wwwroot\\Uploads\\";
                        if (!Directory.Exists(uploadsFolder))
                        {
                            Directory.CreateDirectory(uploadsFolder);
                        }
                        var filePath = Path.Combine(uploadsFolder, excel.FileName);
                        using (var stream = new FileStream(filePath, FileMode.Create))
                        {
                            await excel.CopyToAsync(stream);
                        }

                        // проверить xls или xlsx?

                        IWorkbook workbook;
                        using (FileStream fileStream = new FileStream(Path.GetFullPath(uploadsFolder + excel.FileName), FileMode.Open, FileAccess.Read))
                        {
                            workbook = new XSSFWorkbook(fileStream);
                        }


                        List<DataTable> DataAllSheets = new List<DataTable>();
                        // считаваем sheets

                        List<string> SheetsNames = new List<string>();

                        for (int i = 0; i < workbook.NumberOfSheets; i++)
                        {
                            string SheetName = workbook.GetSheetName(i).Trim();
                            ISheet sheet = workbook.GetSheetAt(i);

                            if (!string.IsNullOrEmpty(SheetName))
                            {
                                SheetsNames.Add(SheetName);
                            }

                            // подумать как парсить?? вообщем скорее всего DataTable 
                            // xlsx 
                            string Title = sheet.GetRow(0).GetCell(0).ToString();

                            var dataSheet = new DataTable(sheet.SheetName);

                            IRow FirstHead = sheet.GetRow(2);
                            IRow SecondHead = sheet.GetRow(3);
                            List<string> Header = new List<string>();
                            List<string> HeaderTypes = new List<string>();

                            // Типы данных
                            foreach (ICell _type in sheet.GetRow(4))
                            {
                                HeaderTypes.Add(_type.CellType.ToString());
                            }

                            // Данные в таблице
                            for (int jj = 0; jj < FirstHead.Cells.Count; jj++)
                            {

                                Header.Add((FirstHead.GetCell(jj) + " " + SecondHead.GetCell(jj).ToString()).Trim());
                                dataSheet.Columns.Add((FirstHead.GetCell(jj).ToString() + " " + SecondHead.GetCell(jj).ToString()).Trim());
                            }

                            // Сами данные + Сразу загрузка в БД  (можно передать сразу DataTable -> MS SQL (Создать процедуру вроде))
                            using (IDbConnection db = Connection)
                            {

                                db.Query("CREATE TABLE " + SheetName.Replace(" ", "") + "(Date TEXT, Time TEXT, Temp TEXT, Humidity TEXT," +
                                                                       "PointDew TEXT, Pressure TEXT, WindOrien TEXT, " +
                                                                       "WindSpeed TEXT, Cloud TEXT, " +
                                                                       "h TEXT, VV TEXT, Weather TEXT);");



                                for (int jj = 4; jj <= sheet.LastRowNum; jj++)
                                {
                                    string _dat = "";
                                    string check;

                                    for (int ii = 0; ii < Header.Count; ii++)
                                    {

                                        try
                                        {
                                            sheet.GetRow(jj).GetCell(ii).ToString();
                                            check = "'" + sheet.GetRow(jj).GetCell(ii).ToString() + "'";
                                        }
                                        catch
                                        {
                                            check = "''";
                                        }

                                        _dat = _dat + check + ", ";

                                    }
                                    _dat = _dat[..^2];
                                    try
                                    {
                                        string str_test = "INSERT INTO " + SheetName.Replace(" ", "") + " (Date, Time, Temp, Humidity, PointDew, Pressure, WindOrien, WindSpeed, Cloud, h, VV, Weather)" + " VALUES (" + _dat + ");";
                                        db.Query("INSERT INTO " + SheetName.Replace(" ", "") + " (Date, Time, Temp, Humidity, PointDew, Pressure, WindOrien, WindSpeed, Cloud, h, VV, Weather)" + " VALUES (" + _dat + ");");
                                    }
                                    catch { }

                                }

                            }

                        }

                    }
                    model.Text = "Успешно!";
                }
                catch 
                {
                    model.Text = "Ошибка при загрузке файла/файлов!";
                    using (IDbConnection db = Connection)
                    {
                        // очистка БД
                        db.Query("declare @sql nvarchar(max); set @sql = N''; Select @sql += 'Drop table' + QUOTENAME(table_name) + ';' from INFORMATION_SCHEMA.TABLES where TABLE_TYPE = 'Base table'; Exec sp_executesql @sql;");
                    }
                };
            }
            string path = "Uploads";
            string fullPath = Path.GetFullPath(path); // "C:\\Users\\PK\\source\\repos\\ExcelTest\\ExcelTest\\Uploads"
            int cc = fullPath.Count() - 1;
            fullPath = fullPath.Remove(Path.GetFullPath(path).Count() - 8, 8) + "\\wwwroot" + "\\Uploads";

            DirectoryInfo di = new DirectoryInfo(fullPath);
            foreach (var fi in di.GetFiles())
            {
                try { fi.Delete(); } catch { }
            }
            try { di.Delete(); } catch { }
            di.Create();



            return View("UploadExcel", model);
        }
            

    }
}
