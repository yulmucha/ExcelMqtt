using Newtonsoft.Json.Linq;
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
