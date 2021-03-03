using ExcelDemo.Entities;
using OfficeOpenXml;
using OfficeOpenXml.ConditionalFormatting;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Threading.Tasks;

namespace ExcelDemo
{
    class Program
    {
        static async Task Main(string[] args)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(@"C:\Users\Xavier\Desktop\Excel\ExcelDemo-" + DateTime.Now.ToString("ddMMyyyy HHmmss") + ".xlsx");
            var fileRead = new FileInfo(@"C:\Users\Xavier\Desktop\Excel\ExcelDemo-Read.xlsx");

            //await SavaExcelFile(LoadCategoriesData(), file);

            var dataFromExcel = await LoadExcelFile(fileRead);

        }

        private static async Task<List<Product>> LoadExcelFile(FileInfo file)
        {
            List<Product> output = new List<Product>();

            using var package = new ExcelPackage(file);

            await package.LoadAsync(file);

            var sheetCategories = GetCategoriesFromFormatSheet(package.Workbook.Worksheets);

            //Itero todas las sheets
            foreach (var ws in package.Workbook.Worksheets)
            {
                int row = 2; // Comienzo donde comienza la data
                int col = 1;

                var sheetRow = ws.Row(row);
            }

            return output;
        }

        private static List<Category> GetCategoriesFromFormatSheet(ExcelWorksheets worksheets)
        {
            var sheetCategories = new List<Category>();
            var formatSheet = worksheets["format"];

            int categoriesCount = int.Parse(formatSheet.Cells[1, 1].Value.ToString());

            int row = 2;

            for (int i = 0; i < categoriesCount; i ++)
            {
                Category category = new Category()
                {
                    Name = formatSheet.Cells[row, 1].Value.ToString(),
                    Attributes = new List<Entities.Attribute>(),
                    Rows = new List<Entities.Excel.Row>()
                };

                int attributesCount = int.Parse(formatSheet.Cells[row, 3].Value.ToString());

                row++;

                int column = 0;
                int columnCount = 1;

                for (int j = 0; j < attributesCount; j++)
                {
                    Entities.Attribute attribute = new Entities.Attribute()
                    {
                        Name = formatSheet.Cells[row, ++column].Value.ToString(),
                        DisplayName = formatSheet.Cells[row, ++column].Value.ToString(),
                        Column = columnCount
                    };

                    category.Attributes.Add(attribute);
                    columnCount++;
                }

                sheetCategories.Add(category);
                row++;
            }

            foreach (var sheetCategory in sheetCategories)
            {
                var sheetValueRows = worksheets[sheetCategory.Name];

                int highestRow = sheetValueRows.Dimension.End.Row;
                int highetColumn = sheetValueRows.Dimension.End.Column;

                for (int rowIndex = 2; rowIndex <= highestRow; ++rowIndex)
                {
                    var sheetRow = new Entities.Excel.Row();

                    for (int col = 1; col <= highetColumn; ++col)
                    {    
                        sheetRow.addCellValue(col, sheetValueRows.Cells[rowIndex, col].Value.ToString());
                    }

                    if (sheetRow.HasValue())
                    {
                        sheetCategory.Rows.Add(sheetRow);
                    }

                }
            }

            return sheetCategories;
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

        #region Data Population
        private static List<Category> LoadCategoriesData()
        {
            return new List<Category>()
            {
                new Category()
                {
                    Name = "Tractores",
                    Attributes = new List<Entities.Attribute>()
                    {
                        new Entities.Attribute() {Id = 1, DisplayName = "Nombre", IsMultiSelect = false, IsRequired = true, Name = "name", Type = "literal" },
                        new Entities.Attribute() {Id = 2, DisplayName = "Descripción", IsMultiSelect = false, IsRequired = true, Name = "description", Type = "literal" },
                        new Entities.Attribute() {
                            Id = 3,
                            DisplayName = "Modelo",
                            IsMultiSelect = true,
                            IsRequired = false,
                            Name = "model",
                            Type = "literal-array",
                            Values = new List<AttributeValue>()
                            {
                                new AttributeValue() { Id = 1, Value = "Modelo 1" },
                                new AttributeValue() { Id = 2, Value = "Modelo 2" }
                            }},
                        new Entities.Attribute() {
                            Id = 4,
                            DisplayName = "Marca",
                            IsMultiSelect = true,
                            IsRequired = true,
                            Name = "brand",
                            Type = "literal-array",
                            Values = new List<AttributeValue>()
                            {
                                new AttributeValue() { Id = 1, Value = "Marca 1" },
                                new AttributeValue() { Id = 2, Value = "Marca 2" }
                            }}
                    }
                }
            };
        }
        #endregion



    }
}
