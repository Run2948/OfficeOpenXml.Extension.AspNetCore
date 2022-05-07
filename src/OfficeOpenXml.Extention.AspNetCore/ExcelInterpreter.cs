using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace OfficeOpenXml.Extention.AspNetCore
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
                    int fromRow = interpRow;
                    int num = outputRow;
                    outputRow = num + 1;
                    SetRow(fromRow, num, data);
                    interpRow++;
                }
                else
                {
                    string text2 = text.Substring(interpChar, text.Length - interpChar);
                    if (regExp.IsMatch(text2))
                    {
                        string value = regExp.Match(text2).Value;
                        interpChar += value.Length;
                        int num2 = value.IndexOf('(');
                        int num3 = value.LastIndexOf(')');
                        string text3 = value.Substring(0, num2).Trim();
                        string text4 = value.Substring(num2 + 1, num3 - num2 - 1);
                        string a = text3;
                        if (!(a == "for"))
                        {
                            throw new Exception(string.Format("Unknown expression: {0}", text3));
                        }
                        IEnumerable<string> source = from m in text4.Split(new string[]
                        {
                            " in "
                        }, 2, StringSplitOptions.RemoveEmptyEntries)
                                                     select m.Trim();
                        if (source.Count() != 2)
                        {
                            throw new Exception("Invalid expression：" + text4);
                        }
                        string[] source2 = source.First().Split(new char[]
                        {
                            ','
                        });
                        string key = source2.First();
                        string text5 = source2.ElementAtOrDefault(1);
                        IEnumerable enumerable = GetValue(source.ElementAt(1), data) as IEnumerable;
                        if (enumerable == null)
                        {
                            throw new Exception(string.Format("{0} can not be Enumerated", source.ElementAt(1)));
                        }
                        int num4 = interpRow;
                        int num5 = interpChar;
                        int num6 = 1;
                        foreach (object value2 in enumerable)
                        {
                            data.Add(key, value2);
                            if (!string.IsNullOrWhiteSpace(text5))
                            {
                                data.Add(text5, num6);
                            }
                            interpRow = num4;
                            interpChar = num5;
                            Run(data);
                            data.Remove(key);
                            if (!string.IsNullOrWhiteSpace(text5))
                            {
                                data.Remove(text5);
                            }
                            num6++;
                        }
                        if (num6 == 1)
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
                                int fromRow2 = interpRow;
                                int num = outputRow;
                                outputRow = num + 1;
                                SetRow(fromRow2, num, data);
                            }
                            interpChar += regBlockStart.Match(text2).Value.Length;
                        }
                        else
                        {
                            if (regBlockEnd.IsMatch(text2))
                            {
                                if (regDisplay.IsMatch(text))
                                {
                                    int fromRow3 = interpRow;
                                    int num = outputRow;
                                    outputRow = num + 1;
                                    SetRow(fromRow3, num, data);
                                }
                                interpChar += regBlockEnd.Match(text2).Value.Length;
                                break;
                            }
                            if (!string.IsNullOrWhiteSpace(text2))
                            {
                                throw new Exception(string.Format("Invalid expression: {0}", text2));
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
            Console.WriteLine(string.Format("Rending line {0} to line {1}", fromRow, newRow));
            _sheet.InsertRow(newRow, 1, fromRow);
            _sheet.Row(newRow).Height = _sheet.Row(fromRow).Height;
            for (int i = 2; i <= _columns; i++)
            {
                Console.WriteLine(string.Format("Cell {0}{1}", ((char)(65 + i)).ToString(), fromRow));
                string text = _sheet.MergedCells[fromRow, i];
                if (text != null)
                {
                    ExcelAddress excelAddress = new ExcelAddress(text);
                    ExcelRange merCel = _sheet.Cells[excelAddress.Start.Row + (newRow - fromRow), excelAddress.Start.Column, excelAddress.End.Row + (newRow - fromRow), excelAddress.End.Column];
                    if (!merCel.Merge)
                    {
                        Console.WriteLine(string.Format("Merge cell {0}", merCel.ToString()));
                        if (!_mergedRegions.Any((a) => a.Address == merCel.Address))
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
            MatchCollection source = regReplace.Matches(name);
            IEnumerable<string> enumerable = (from Match m in source
                                              select m.Value).Distinct();
            object result;
            if (enumerable.Count() == 1 && enumerable.First().Length == name.Length)
            {
                string text = enumerable.First();
                result = ExecExpression(text.Substring(1, text.Length - 2), data);
            }
            else
            {
                foreach (string text2 in enumerable)
                {
                    string expression = text2.Substring(1, text2.Length - 2);
                    object obj = ExecExpression(expression, data);
                    name = name.Replace(text2, obj.ToString());
                }
                result = name;
            }
            return result;
        }

        private object ExecExpression(string expression, Dictionary<string, object> data)
        {
            return GetValue(expression, data);
        }

        private object GetValue(string name, Dictionary<string, object> data)
        {
            object obj = null;
            string[] source = name.Split(new char[]
            {
                '.'
            });
            try
            {
                obj = data[source.First()];
            }
            catch
            {
                throw new Exception("Variable not found：" + source.First());
            }
            foreach (string text in source.Skip(1))
            {
                PropertyInfo property = obj.GetType().GetProperty(text);
                if (property == null)
                {
                    throw new Exception(string.Format("Type \"{0}\" does not contains member: {1}", obj.GetType().Name, text));
                }
                obj = property.GetValue(obj, null);
            }
            return obj;
        }
    }
}
