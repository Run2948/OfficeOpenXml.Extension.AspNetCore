using System;

namespace OfficeOpenXml.Extention.AspNetCore.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnAttribute : Attribute
    {
        public int Number { get; set; }
    }
}