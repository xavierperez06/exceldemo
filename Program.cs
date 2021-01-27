using ExcelDemo.Entities;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace ExcelDemo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(@"C:\Users\Xavier\Desktop\Excel\ExcelDemo-" + DateTime.Now.ToString("ddMMyyyy HHmmss") + ".xlsx");

            await SavaExcelFile(LoadData(), file);

        }

        private static async Task SavaExcelFile(List<Category> categories, FileInfo file)
        {
            using var package = new ExcelPackage(file);

            foreach(var category in categories)
            {
                //Creo un nuevo sheet para la categoria
                var ws = package.Workbook.Worksheets.Add(category.Name);

                int index = 1;

                // Tiene un ws.Cells["RANGO"].LoadFromCollection(attributes) que te carga toda la data, pero usa el header del nombre de cada propiedad de la clase
                foreach (var attribute in category.Attributes)
                {
                    ws.Cells[1, index].Value = attribute.DisplayName + (attribute.IsRequired ? "*" : "");

                    if (attribute.Type == "literal-array")
                    {
                        ws.Cells[1, index].Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                        ws.Cells[1, index].Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.AliceBlue);
                    }

                    ws.Cells[1, index].AutoFitColumns();

                    index++;
                }
            }

            await package.SaveAsync();
        }

        private static List<Category> LoadData()
        {
            return new List<Category>()
            {
                new Category()
                {
                    Name = "Tractores",
                    Attributes = new List<Attribute>()
                    {
                        new Attribute() {Id = 1, DisplayName = "Nombre", IsMultiSelect = false, IsRequired = true, Name = "name", Type = "literal" },
                        new Attribute() {Id = 2, DisplayName = "Descripción", IsMultiSelect = false, IsRequired = true, Name = "description", Type = "literal" },
                        new Attribute() {Id = 3, DisplayName = "Modelo", IsMultiSelect = true, IsRequired = false, Name = "model", Type = "literal-array" },
                        new Attribute() {Id = 4, DisplayName = "Marca", IsMultiSelect = true, IsRequired = true, Name = "brand", Type = "literal-array" }
                    }
                }
            };    
        }


    }
}
