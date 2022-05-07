using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;
using OfficeOpenXml.Extention.AspNetCore.Attributes;

namespace OfficeOpenXml.Extention.AspNetCore
{
    public static class ExcelPackageExtensions
    {
        public static IEnumerable<T> ParseWorksheet<T>(this ExcelPackage excelPackage) where T : new()
        {
            var items = new List<T>();
            var worksheetIndex = 0;
            var startRowNumber = 1;

            var type = typeof(T);
            var worksheetAttribute = type.GetCustomAttribute<WorksheetAttribute>();
            if (worksheetAttribute != null)
            {
                worksheetIndex = worksheetAttribute.Index;
                startRowNumber = worksheetAttribute.HasHeader ? 2 : 1;
            }

            var worksheet = excelPackage.Workbook.Worksheets[worksheetIndex];
            var rowsCount = worksheet.Dimension?.Rows;

            for (var rowNumber = startRowNumber; rowNumber <= rowsCount; rowNumber++)
            {
                var item = new T();

                var properties = type.GetProperties();
                foreach (var property in properties)
                {
                    var columnAttribute = property.GetCustomAttribute<ColumnAttribute>();
                    if (columnAttribute == null) continue;
                    var valueAsString = worksheet.GetValue<string>(Row: rowNumber, Column: columnAttribute.Number);
                    var converter = TypeDescriptor.GetConverter(property.PropertyType);
                    property.SetValue(item, converter.ConvertFromString(valueAsString));
                }

                items.Add(item);
            }

            return items;
        }

        public static void FillModel(this ExcelWorksheet sheet, object model)
        {
            ExcelInterpreter exlInterpreter = new ExcelInterpreter(sheet);
            exlInterpreter.Complie(new Dictionary<string, object>
            {
                {
                    "model",
                    model
                }
            });
        }
    }
}