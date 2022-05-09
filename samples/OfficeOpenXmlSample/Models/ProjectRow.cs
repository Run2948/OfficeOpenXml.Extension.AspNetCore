using OfficeOpenXml.Extention.AspNetCore.Attributes;

namespace OfficeOpenXmlSample.Models
{
    [Worksheet(Index = 1, HasHeader = false)]
    public class ProjectRow
    {
        [Column(Number = 1)]
        public int Id { get; set; }

        [Column(Number = 2)]
        public string Name { get; set; }

        [Column(Number = 3)]
        public string Description { get; set; }
    }
}
