using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CSharpLibraryForExcel
{
    public class ExcelCOM
    {
        private readonly Config mConfig;
        private readonly Application mApp;
        private readonly Workbook mWorkbook;
        private readonly Worksheet mWorksheet;


        public int TotalRecords
        {
            get { return LastRow - mConfig.StartRecordRow + 1; }
        }

        public int LastRow
        {
            get { return mWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row; }
        }

        public int LastColumn
        {
            get { return mWorksheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column; }
        }

        public int TotalChunks
        {
            get
            {
                return (int)Math.Ceiling((double)TotalRecords / mConfig.ChunkSize);
            }
        }

        public List<string> Columns
        {
            get
            {
                var result = new List<string>(LastColumn);
                for (int col = 1; col <= LastColumn; col++)
                {
                    var cell = mWorksheet.Cells[mConfig.PropertyRow, col];
                    result.Add(cell.Value == null ? string.Empty : cell.Value.ToString());
                }
                return result;
            }
        }

        public ExcelCOM(Config config)
        {
            this.mConfig = config;

            try
            {
                mApp = (Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains("0x800401E3 (MK_E_UNAVAILABLE)"))
                {
                    mApp = new Application();
                }
                else
                {
                    throw;
                }
            }

            mWorkbook = mApp.Workbooks.Open(mConfig.ExcelFileName);
            mWorksheet = GetSheet();
        }

        public Worksheet GetSheet()
        {
            if (mWorkbook.Worksheets.Count == 1)
            {
                return mWorkbook.Worksheets[1];
            }
            else
            {
                try
                {
                    return mWorkbook.Worksheets[mConfig.ExcelSheetName];
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    throw new ArgumentException("존재하지 않는 엑셀 시트 이름입니다.");
                }
            }
        }

        public JArray GetRowJson(int row)
        {
            if (row < mConfig.StartRecordRow)
            {
                throw new ArgumentException("변환하려는 행이 시작 행보다 앞섭니다.");
            }

            JArray result = new JArray();
            for (int col = 1; col <= LastColumn; col++)
            {
                result.Add(mWorksheet.Cells[row, col].Value);
            }
            return result;
        }


        public JObject LabelRowJson(JArray rowValues)
        {
            List<string> columns = Columns;
            if (columns.Count != rowValues.Count)
            {
                throw new ArgumentException("컬럼 개수와 행의 데이터 개수가 일치하지 않습니다.");
            }

            JObject result = new JObject();
            for (int i = 0; i < columns.Count; i++)
            {
                result.Add(columns[i], rowValues[i]);
            }

            return result;
        }

        public JArray GetRecordsJson()
        {
            JArray recordsJson = new JArray();
            for (int row = mConfig.StartRecordRow; row <= LastRow; row++)
            {
                recordsJson.Add(LabelRowJson(GetRowJson(row)));
            }
            return recordsJson;
        }

        public List<JArray> ChunkRecordsJson(JArray records)
        {
            List<JArray> chunks = new List<JArray>();
            var chunk = new JArray();
            for (int i = 0; i < records.Count; i++)
            {
                chunk.Add(records[i]);
                if (i == records.Count - 1 || chunk.Count % mConfig.ChunkSize == 0)
                {
                    chunks.Add(chunk);
                    chunk = new JArray();
                }
            }
            return chunks;
        }

        public List<JObject> GetMqttMessages()
        {
            return ChunkRecordsJson(GetRecordsJson()).Select((c, i) => new JObject
            {
                { "rows", TotalRecords },
                { "chunkSequence", i + 1},
                { "chunks", TotalChunks},
                { "chunkSize", mConfig.ChunkSize },
                { "columns", Columns.Count},
                { "data", c }
            }).ToList();
        }

        public int GetColumnNumberOf(string keyColumn)
        {
            int result = Columns.IndexOf(keyColumn);
            
            if (result == -1)
            {
                throw new ArgumentException("일치하는 keyColumn 값이 없음");
            }

            return result + 1;
        }

        public int GetRowNumberOf(dynamic key, int columnNumber)
        {
            for (int row = 1; row <= LastRow; row++)
            {
                if (key.Equals(mWorksheet.Cells[row, columnNumber].Value))
                {
                    return row;
                }
            }

            throw new ArgumentException("일치하는 key 값이 없음");
        }

        public void Dispose()
        {
            System.Runtime.InteropServices.Marshal.ReleaseComObject(mWorksheet);
        }
    }
}
