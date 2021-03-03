
using System.Collections.Generic;

namespace ExcelDemo.Entities
{
    public class Attribute
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string DisplayName { get; set; }
        public bool IsRequired { get; set; }
        public string Type { get; set; }
        public bool IsMultiSelect { get; set; }
        public int Column { get; set; }
        public List<AttributeValue> Values { get; set; }
    }
}
