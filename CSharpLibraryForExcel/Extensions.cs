using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSharpLibraryForExcel
{
    public static class Extensions
    {
        public static List<string> SelectRow(this Excel.Worksheet sheet, int row)
        {
            int lastCol = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column;
            var result = new List<string>(lastCol);
            for (int col = 1; col <= lastCol; col++)
            {
                var cell = sheet.Cells[row, col];
                result.Add(cell.Value == null ? string.Empty : cell.Value.ToString());
            }
            return result;
        }

        private static JObject labelRow(List<string> columnNames, List<string> row)
        {
            var rowObj = new JObject();
            for (int i = 0; i < row.Count(); i++)
            {
                rowObj.Add(columnNames.ElementAt(i), row.ElementAt(i));
            }
            return rowObj;
        }

        public static JArray ToRecordsJson(this Excel.Worksheet sheet, Config config)
        {
            int lastRow = sheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;

            var columnNames = sheet.SelectRow(config.PropertyRow);

            var recordsJson = new JArray();
            for (int rowNum = config.StartRecordRow; rowNum <= lastRow; rowNum++)
            {
                recordsJson.Add(labelRow(columnNames, sheet.SelectRow(rowNum)));
            }

            return recordsJson;
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
