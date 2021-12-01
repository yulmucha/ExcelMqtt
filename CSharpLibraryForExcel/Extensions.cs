using OfficeOpenXml;

namespace CSharpLibraryForExcel
{
    public static class Extensions
    {
        public static bool IsEmptyRow(this ExcelWorksheet sheet, int row)
        {
            for (int col = 1; col <= sheet.Dimension.End.Column; col++)
            {
                if (sheet.Cells[row, col].Value != null)
                {
                    return false;
                }
            }

            return true;
        }
    }
}
