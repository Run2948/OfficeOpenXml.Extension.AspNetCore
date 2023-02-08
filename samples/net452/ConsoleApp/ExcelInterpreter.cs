using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Extension.AspNetCore
{
    internal class ExcelInterpreter
    {
        private readonly ExcelWorksheet _sheet;

        private readonly List<ExcelRange> _mergedRegions;

        private readonly int _rows;

        private readonly int _columns;

        private readonly Regex regReplace = new Regex("{\\S+?}");

        private readonly Regex regExp = new Regex("([a-z]+)\\s*\\([\\S|\\s]+?\\)\\s*");

        private readonly Regex regBlockStart = new Regex("^{\\s*");

        private readonly Regex regBlockEnd = new Regex("^}\\s*");

        private readonly Regex regDisplay = new Regex("(^{\\s*$|^{\\s*}\\s*|^(\\s*)})");

        private readonly Regex regMethod = new Regex("([a-zA-Z|\\.]+)\\s*\\([\\S|\\s]+?\\)\\s*");

        public ExcelInterpreter(ExcelWorksheet sheet)
        {
            _sheet = sheet;
            _rows = sheet.Dimension.End.Row;
            _columns = sheet.Dimension.End.Column;
            _mergedRegions = new List<ExcelRange>();
        }

        private int interpRow = 1;

        private int interpChar = 0;

        private int outputRow = 1;

        public void Complie(Dictionary<string, object> values)
        {
            outputRow = _rows + 1;
            Run(values);
            foreach (ExcelRange excelRange in _mergedRegions)
            {
                _sheet.Cells[excelRange.Address].Merge = true;
            }
            _sheet.DeleteColumn(1);
            _sheet.DeleteRow(1, _rows, true);
        }

        private void Run(Dictionary<string, object> data)
        {
            while (interpRow <= _rows)
            {
                string text = _sheet.Cells[interpRow, 1].Text.Trim();
                if (string.IsNullOrWhiteSpace(text))
                {
                    SetRow(interpRow, outputRow++, data);
                    interpRow++;
                }
                else
                {
                    string text2 = text.Substring(interpChar, text.Length - interpChar);
                    if (regExp.IsMatch(text2))
                    {
                        string value = regExp.Match(text2).Value;
                        interpChar += value.Length;
                        int num = value.IndexOf('(');
                        int num2 = value.LastIndexOf(')');
                        string text3 = value.Substring(0, num).Trim();
                        string text4 = value.Substring(num + 1, num2 - num - 1);
                        string text5 = text3;
                        if (text5 != "for")
                        {
                            throw new Exception($"Unknown expression: {text3}");
                        }
                        var source = from m in text4.Split(new []{ " in " }, 2, StringSplitOptions.RemoveEmptyEntries) select m.Trim();
                        if (source.Count() != 2)
                        {
                            throw new Exception("Invalid expression：" + text4);
                        }
                        string[] source2 = source.First().Split(',');
                        string key = source2.First();
                        string text6 = source2.ElementAtOrDefault(1);
                        var enumerable = GetValue(source.ElementAt(1), data) as IEnumerable;
                        if (enumerable == null)
                        {
                            throw new Exception($"{source.ElementAt(1)} can not be Enumerated");
                        }
                        int num3 = interpRow;
                        int num4 = interpChar;
                        int num5 = 1;
                        foreach (object item in enumerable)
                        {
                            data.Add(key, item);
                            if (!string.IsNullOrWhiteSpace(text6))
                            {
                                data.Add(text6, num5);
                            }
                            interpRow = num3;
                            interpChar = num4;
                            Run(data);
                            data.Remove(key);
                            if (!string.IsNullOrWhiteSpace(text6))
                            {
                                data.Remove(text6);
                            }
                            num5++;
                        }
                        if (num5 == 1)
                        {
                            interpRow++;
                            interpChar = 0;
                        }
                        Run(data);
                    }
                    else
                    {
                        if (regBlockStart.IsMatch(text2))
                        {
                            if (regDisplay.IsMatch(text2))
                            {
                                SetRow(interpRow, outputRow++, data);
                            }
                            interpChar += regBlockStart.Match(text2).Value.Length;
                        }
                        else
                        {
                            if (regBlockEnd.IsMatch(text2))
                            {
                                if (regDisplay.IsMatch(text))
                                {
                                    SetRow(interpRow, outputRow++, data);
                                }
                                interpChar += regBlockEnd.Match(text2).Value.Length;
                                break;
                            }
                            if (!string.IsNullOrWhiteSpace(text2))
                            {
                                throw new Exception($"Invalid expression: {text2}");
                            }
                            interpChar = 0;
                            interpRow++;
                        }
                    }
                }
            }
        }

        private void SetRow(int fromRow, int newRow, Dictionary<string, object> data)
        {
            Console.WriteLine($"Rending line {fromRow} to line {newRow}");
            _sheet.InsertRow(newRow, 1, fromRow);
            _sheet.Row(newRow).Height = _sheet.Row(fromRow).Height;
            for (int i = 2; i <= _columns; i++)
            {
                Console.WriteLine($"Cell {(char)(65 + i)}{fromRow}");
                string text = _sheet.MergedCells[fromRow, i];
                if (text != null)
                {
                    ExcelAddress excelAddress = new ExcelAddress(text);
                    ExcelRange merCel = _sheet.Cells[excelAddress.Start.Row + (newRow - fromRow), excelAddress.Start.Column, excelAddress.End.Row + (newRow - fromRow), excelAddress.End.Column];
                    if (!merCel.Merge)
                    {
                        Console.WriteLine($"Merge cell {merCel}");
                        if (!_mergedRegions.Any(a => a.Address == merCel.Address))
                        {
                            _mergedRegions.Add(merCel);
                        }
                    }
                }
                ExcelRange excelRange = _sheet.Cells[fromRow, i];
                ExcelRange excelRange2 = _sheet.Cells[newRow, i];
                if (!string.IsNullOrWhiteSpace(excelRange.FormulaR1C1))
                {
                    excelRange2.FormulaR1C1 = excelRange.FormulaR1C1;
                }
                else
                {
                    object value = excelRange.Value;
                    if (value != null)
                    {
                        string text2 = value.ToString();
                        if (text2.StartsWith("@="))
                        {
                            excelRange2.FormulaR1C1 = ReplaceParam(text2.Substring(1), data).ToString();
                        }
                        else
                        {
                            excelRange2.Value = ReplaceParam(value.ToString(), data);
                        }
                    }
                }
            }
        }

        private object ReplaceParam(string name, Dictionary<string, object> data)
        {
            var source = regReplace.Matches(name);
            var enumerable = (from Match m in source select m.Value).Distinct();
            if (enumerable.Count() == 1 && enumerable.First().Length == name.Length)
            {
                string text = enumerable.First();
                return ExecExpression(text.Substring(1, text.Length - 2), data);
            }
            foreach (string item in enumerable)
            {
                string expression = item.Substring(1, item.Length - 2);
                object obj = ExecExpression(expression, data);
                name = name.Replace(item, obj.ToString());
            }
            return name;
        }

        private object ExecExpression(string expression, Dictionary<string, object> data)
        {
            return GetValue(expression, data);
        }

        private object GetValue(string name, Dictionary<string, object> data)
        {
            string[] source = name.Split('.');
            object obj;
            try
            {
                obj = data[source.First()];
            }
            catch
            {
                throw new Exception("Variable not found：" + source.First());
            }
            foreach (string item in source.Skip(1))
            {
                PropertyInfo property = obj.GetType().GetProperty(item);
                if (property == null)
                {
                    throw new Exception($"Type \"{obj.GetType().Name}\" does not contains member: {item}");
                }
                obj = property.GetValue(obj, null);
            }
            return obj;
        }
    }
}
