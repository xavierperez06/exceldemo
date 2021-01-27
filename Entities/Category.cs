using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDemo.Entities
{
    public class Category
    {
        public string Name { get; set; }
        public List<Attribute> Attributes { get; set; }
    }
}
