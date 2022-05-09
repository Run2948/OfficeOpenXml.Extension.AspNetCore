using OfficeOpenXml.Extention.AspNetCore.Attributes;

namespace OfficeOpenXmlSample.Models
{
    [Worksheet]
    public class TodoRow
    {
        public int Id { get; set; }

        [Column(Number = 2)]
        public string Title { get; set; }

        [Column(Number = 3)]
        public int Priority { get; set; }

        [Column(Number = 4)]
        public bool Completed { get; set; }
    }
}
