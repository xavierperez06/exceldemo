using System;
using System.Collections.Generic;
using System.Text;

namespace ExcelDemo.Entities
{
    public class Category
    {
        public string Name { get; set; }
        public List<Attribute> Attributes { get; set; }
        //Rows no iria acá, sino que se deberia crear una clase sheetCategory (o buscar algo mas generico)
        public List<Excel.Row> Rows { get; set; }
    }
}
