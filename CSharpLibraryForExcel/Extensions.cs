using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System.Collections.Generic;
using System.Linq;

namespace CSharpLibraryForExcel
{
    public static class Extensions
    {
        public static List<string> SelectRow(this ExcelWorksheet sheet, int row)
        {
            var result = new List<string>(sheet.Dimension.End.Column);
            for (int col = 1; col <= sheet.Dimension.End.Column; col++)
            {
                var cell = sheet.Cells[row, col];
                result.Add(cell.Value == null ? string.Empty : cell.Value.ToString());
            }
            return result;
        }
        public static List<JObject> ToMessage(this List<JArray> chunks, int totalRecords, int totalColumns, int totalChunks, int chunkSize)
        {
            return chunks.Select((c, i) => new JObject
            {
                { "rows", totalRecords },
                { "chunkSequence", i + 1},
                { "chunks", totalChunks},
                { "chunkSize", chunkSize },
                { "columns", totalColumns},
                { "data", c }
            }).ToList();
        }
    }
}
