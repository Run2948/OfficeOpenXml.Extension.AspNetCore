using Microsoft.AspNetCore.Mvc;
using OfficeOpenXml;
using OfficeOpenXml.Extention.AspNetCore;
using OfficeOpenXmlSample.Models;

namespace OfficeOpenXmlSample.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly string _wwwroot;
        private readonly ILogger<ExcelController> _logger;

        private readonly Dictionary<string, IEnumerable<string>> _marketLists = new()
        {
            { "ˮ��", new string[] { "����", "����", "�㽶", "��" } },
            { "�߲�", new string[] { "���", "����", "�ƹ�", "ơ��" } }
        };

        public ExcelController(IWebHostEnvironment env, ILogger<ExcelController> logger)
        {
            _wwwroot = env.WebRootPath;
            _logger = logger;
        }

        #region Read

        [HttpGet("todos", Name = "Todos")]
        public IEnumerable<TodoRow> GetTodos()
        {
            var excelFilePath = Path.Combine(_wwwroot, "templates", "Todos.xlsx");
            var fileStream = new System.IO.FileStream(excelFilePath, FileMode.Open);
            using var excelPackage = new ExcelPackage(fileStream);
            return excelPackage.ParseWorksheet<TodoRow>().ToList();
        }

        [HttpGet("projects", Name = "Projects")]
        public IEnumerable<ProjectRow> GetProjects()
        {
            var excelFilePath = Path.Combine(_wwwroot, "templates", "Projects.xlsx");
            var fileStream = new System.IO.FileStream(excelFilePath, FileMode.Open);
            using var excelPackage = new ExcelPackage(fileStream);
            return excelPackage.ParseWorksheet<ProjectRow>().ToList();
        }

        #endregion

        #region Write

        [HttpGet("lists", Name = "ExportLists")]
        public IActionResult ExportLists()
        {
            var excelFilePath = Path.Combine(_wwwroot, "templates", "tpl.xlsx");
            var fileStream = new System.IO.FileStream(excelFilePath, FileMode.Open);
            using var excelPackage = new ExcelPackage(fileStream);
            var workBook = excelPackage.Workbook;

            var random = new Random();

            //����model
            var model = new
            {
                ProjectName = "��̫��",
                Name = "Jeff",
                CreatedAt = DateTime.Now,
                BuyerName = "Bill",
                Cates = _marketLists.Select(m => new
                {
                    Name = m.Key,
                    Items = m.Value.Select(n => new
                    {
                        Name = n,
                        Price = (decimal)random.Next(1, 100),
                        Amount = random.Next(1, 100)
                    })
                })
            };

            // �����FillModel���� OfficeOpenXml.Extension.AspNetCore �ṩ����չ����
            workBook.Worksheets.First().FillModel(model);

            string fileName = "Lists_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            string exportFilePath = Path.Combine(_wwwroot, "outputs", fileName);
            var exportFile = new FileInfo(exportFilePath);
            excelPackage.SaveAs(exportFile);

            return File(exportFile.OpenRead(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        [HttpGet("lists2", Name = "ExportLists2")]
        public IActionResult ExportLists2()
        {
            var excelFilePath = Path.Combine(_wwwroot, "templates", "tpl2.xlsx");
            var fileStream = new System.IO.FileStream(excelFilePath, FileMode.Open);
            using var excelPackage = new ExcelPackage(fileStream);
            var workBook = excelPackage.Workbook;

            var random = new Random();

            //����model
            var model = new
            {
                ProjectName = "��̫��",
                Name = "Jeff",
                CreatedAt = DateTime.Now,
                BuyerName = "Bill",
                Cates = _marketLists.Select(m => new
                {
                    Name = m.Key,
                    Items = m.Value.Select(n => new
                    {
                        Name = n,
                        Price = (decimal)random.Next(1, 100),
                        Amount = random.Next(1, 100)
                    }).ToList(),
                })
            };

            // �����FillModel���� OfficeOpenXml.Extension.AspNetCore �ṩ����չ����
            workBook.Worksheets.First().FillModel(model);

            string fileName = "Lists2_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx";
            string exportFilePath = Path.Combine(_wwwroot, "outputs", fileName);
            var exportFile = new FileInfo(exportFilePath);
            excelPackage.SaveAs(exportFile);

            return File(exportFile.OpenRead(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileName);
        }

        #endregion
    }
}