using OfficeOpenXml;
using OfficeOpenXml.Extension.AspNetCore;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo info = new FileInfo("tpl.xlsx");
            ExcelPackage packet = new ExcelPackage(info);
            var book = packet.Workbook;

            Random random = new Random();

            var dic = new Dictionary<string, IEnumerable<string>>
            {
                { "水果", new string[] { "桃子", "李子", "香蕉", "梨" } },
                { "蔬菜", new string[] { "青菜", "土豆", "黄瓜", "啤酒" } }
            };

            //构造model
            var model = new
            {
                ProjectName = "灰太狼",
                Name = "Jeff",
                CreatedAt = DateTime.Now,
                BuyerName = "Bill",
                Cates = dic.Select(m => new
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

            book.Worksheets.First().FillModel(model);

            packet.SaveAs(new FileInfo(DateTime.Now.ToString("yyyyMMddHHmmss") + ".xlsx"));
        }
    }
}
