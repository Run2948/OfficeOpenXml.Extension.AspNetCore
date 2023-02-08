using System;

namespace OfficeOpenXml.Extension.AspNetCore.Attributes
{
    [AttributeUsage(AttributeTargets.Property)]
    public class ColumnAttribute : Attribute
    {
        public int Number { get; set; }
    }
}