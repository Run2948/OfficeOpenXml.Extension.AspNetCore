using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;

namespace OfficeOpenXml.Extention.AspNetCore
{
    public class ExcelRender
    {
        private readonly string _templateFile;
        private readonly ExcelPackage _excelPackage;
        public ExcelRender(string templateFile)
        {
            _templateFile = templateFile ?? throw new Exception($"File \"{templateFile}\" does not exist");
            _excelPackage = new ExcelPackage(new FileInfo(_templateFile));
        }

        public Dictionary<string, object> KeyValues { get; } = new Dictionary<string, object>();

        public void RenderAndSave(string outputFile)
        {
            ExcelWorksheets worksheets = _excelPackage.Workbook.Worksheets;
            foreach (ExcelWorksheet sheet in worksheets)
            {
                ExcelInterpreter excelInterpreter = new ExcelInterpreter(sheet);
                excelInterpreter.Complie(KeyValues);
            }
            _excelPackage.SaveAs(new FileInfo(outputFile));
            Process.Start(outputFile);
        }
    }
}
