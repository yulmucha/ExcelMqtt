using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Text;
using uPLibrary.Networking.M2Mqtt;
using uPLibrary.Networking.M2Mqtt.Messages;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSharpLibraryForExcel
{
    public class ExcelMqttHandler
    {
        private readonly Config mConfig;
        private readonly Excel.Application mXlApp;
        private readonly Excel.Workbook mXlWorkbook;
        private readonly Excel.Worksheet mXlWorksheet;
        private readonly MqttClient mMqttClient;

        public int TotalRecords
        {
            get { return XlLastRow - mConfig.StartRecordRow + 1; }
        }

        public int XlLastRow
        {
            get { return mXlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Row;  }
        }

        public int XlLastColumn
        {
            get { return mXlWorksheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing).Column; }
        }

        public ExcelMqttHandler(Config mConfig)
        {
            this.mConfig = mConfig;

            try
            {
                mXlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains("0x800401E3 (MK_E_UNAVAILABLE)"))
                {
                    mXlApp = new Excel.Application();
                }
                else
                {
                    throw;
                }
            }

            mXlWorkbook = mXlApp.Workbooks.Open(mConfig.ExcelFileName);
            
            mXlWorksheet = getSheet();
            
            mMqttClient = new MqttClient(
                brokerHostName: mConfig.BrokerHostName,
                brokerPort: mConfig.BrokerPort,
                secure: false,
                caCert: null);
        }

        public void Publish()
        {
            List<JArray> chunks = chunkRecords(mXlWorksheet.ToRecordsJson(mConfig), mConfig.ChunkSize);

            List<JObject> messages = chunks.ToMessage(
                totalRecords: TotalRecords,
                totalColumns: XlLastColumn,
                totalChunks: (int)Math.Ceiling((double)TotalRecords / mConfig.ChunkSize),
                chunkSize: mConfig.ChunkSize);
            
            mMqttClient.Connect(
                clientId: mConfig.ClientId,
                username: mConfig.Username,
                password: mConfig.Password);

            if (!mMqttClient.IsConnected)
            {
                throw new Exception("MqttClient failed to connect.");
            }

            mMqttClient.MqttMsgPublishReceived += MqttMsgPublishReceived;
            mMqttClient.Subscribe(new string[] { mConfig.Topic + "/REPLY" }, new byte[] { MqttMsgBase.QOS_LEVEL_AT_LEAST_ONCE });

            messages.ForEach(msg => mMqttClient.Publish(
                topic: mConfig.Topic,
                message: Encoding.UTF8.GetBytes(msg.ToString()),
                qosLevel: MqttMsgBase.QOS_LEVEL_AT_LEAST_ONCE,
                retain: false));
        }

        private Excel.Worksheet getSheet()
        {
            if (mXlWorkbook.Worksheets.Count == 1)
            {
                return mXlWorkbook.Worksheets[1];
            }
            else
            {
                try
                {
                    return mXlWorkbook.Worksheets[mConfig.ExcelSheetName];
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    throw new ArgumentException("엑셀 시트 이름이 올바르지 않습니다.");
                }
            }
        }

        private List<JArray> chunkRecords(JArray records, int chunkSize)
        {
            List<JArray> chunks = new List<JArray>();
            var chunk = new JArray();
            for (int i = 0; i < records.Count; i++)
            {
                chunk.Add(records[i]);
                if (i == records.Count - 1 || chunk.Count % chunkSize == 0)
                {
                    chunks.Add(chunk);
                    chunk = new JArray();
                }
            }
            return chunks;
        }

        private bool isValid(JObject response)
        {
            if (response.GetValue("keyColumn") == null) return false;
            if (response.GetValue("key") == null) return false;
            if (response.GetValue("result") == null) return false;
            if (response.GetValue("reason") == null) return false;

            return true;
        }

        private bool tryGetColumnOf(string value, out int? colNum)
        {
            colNum = null;

            for (int c = 1; c <= XlLastColumn; c++)
            {
                var cellValue = mXlWorksheet.Cells[mConfig.PropertyRow, c].Value;
                if (cellValue != null && cellValue.Equals(value))
                {
                    colNum = c;
                    return true;
                }
            }

            return false;
        }

        private bool tryGetRowOf(int column, string value, out int? rowNum)
        {
            rowNum = null;

            for (int r = mConfig.StartRecordRow; r <= XlLastRow; r++)
            {
                var cellValue = mXlWorksheet.Cells[r, column].Value;
                if (cellValue != null && cellValue.Equals(value))
                {
                    rowNum = r;
                    return true;
                }
            }

            return false;
        }

        private void MqttMsgPublishReceived(object sender, MqttMsgPublishEventArgs e)
        {
            var json = JObject.Parse(Encoding.UTF8.GetString(e.Message, 0, e.Message.Length));
            
            if (!isValid(json))
            {
                //throw new InvalidDataException("레코드 전송에 대한 응답 형식이 올바르지 않습니다.");
                return;
            }

            if (!tryGetColumnOf(json.GetValue("keyColumn").ToString(), out int? targetCol))
            {
                return;
            }

            if (!tryGetRowOf(targetCol.Value, json.GetValue("key").ToString(), out int? targetRow))
            {
                return;
            }
                        
            Excel.Range targetCell = mXlWorksheet.Cells[targetRow.Value, targetCol.Value];
            if (json.GetValue("result").ToString().Equals("OK"))
            {
                targetCell.Interior.ColorIndex = 35;
            }
            else
            {
                targetCell.ClearComments();
                targetCell.AddComment(json.GetValue("reason").ToString());
                targetCell.Interior.ColorIndex = 36;
            }
            mXlWorkbook.Save();
        }
    }
}
