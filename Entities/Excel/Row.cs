using System.Collections.Generic;
using System.Linq;

namespace ExcelDemo.Entities.Excel
{
    public class Row
    {
        private Dictionary<int, string> cells;

        public Row()
        {
            cells = new Dictionary<int, string>();
        }

        public void addCellValue(int column, string value)
        {
            cells.Add(column, value);
        }

        public Dictionary<int, string> GetCells() => cells;

        public bool HasValue() => GetCells().Any(pair => pair.Value != null && pair.Value.Any());

    }
}
