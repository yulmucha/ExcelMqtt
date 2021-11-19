using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;

namespace CSharpLibraryForExcel
{
    public static class Extensions
    {
        public static IEnumerable<string> SelectRow(this ExcelWorksheet sheet, int row)
        {
            return sheet
                .Cells[row, 1, row, sheet.Dimension.End.Column]
                .Select(c => c.Value == null ? string.Empty : c.Value.ToString());
        }
    }
}
