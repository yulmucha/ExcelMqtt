using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using uPLibrary.Networking.M2Mqtt;
using uPLibrary.Networking.M2Mqtt.Messages;
using Excel = Microsoft.Office.Interop.Excel;

namespace CSharpLibraryForExcel
{
    public class ExcelMqttHandler
    {
        private readonly Config mConfig;
        private readonly ExcelPackage mExcelPackage;
        private readonly ExcelWorksheet mWorksheet;
        private readonly MqttClient mMqttClient;

        public int TotalRecords
        {
            get { return mWorksheet.Dimension.End.Row - mConfig.StartRecordRow + 1; }
        }

        public JArray Records
        {
            get
            {
                var columnNames = mWorksheet.SelectRow(mConfig.PropertyRow);

                var records = new JArray();
                for (int rowNum = mConfig.StartRecordRow; rowNum <= mWorksheet.Dimension.End.Row; rowNum++)
                {
                    records.Add(labelRow(columnNames, mWorksheet.SelectRow(rowNum)));
                }

                return records;
            }
        }

        public ExcelMqttHandler(Config mConfig)
        {
            this.mConfig = mConfig;
            mExcelPackage = new ExcelPackage(new FileInfo(mConfig.ExcelFileName));
            mWorksheet = getSheet(mExcelPackage);
            mMqttClient = new MqttClient(
                brokerHostName: mConfig.BrokerHostName,
                brokerPort: mConfig.BrokerPort,
                secure: false,
                caCert: null);
        }

        public void Publish()
        {
            List<JArray> chunks = chunkRecords(Records, mConfig.ChunkSize);
            
            var messages = new List<JObject>();
            messages = chunks.ToMessage(
                totalRecords: TotalRecords,
                totalColumns: mWorksheet.Dimension.End.Column,
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

        private int totalRecords()
        {
            return mWorksheet.Dimension.End.Row - mConfig.StartRecordRow + 1;
        }

        private ExcelWorksheet getSheet(ExcelPackage xlPackage)
        {
            if (xlPackage.Workbook.Worksheets.Count == 1)
            {
                return xlPackage.Workbook.Worksheets.First();
            }
            else
            {
                var sheet = xlPackage.Workbook.Worksheets[mConfig.ExcelSheetName];
                if (sheet == null)
                {
                    throw new InvalidDataException("엑셀 시트 이름이 올바르지 않습니다.");
                }

                return sheet;
            }
        }

        private JObject labelRow(IEnumerable<string> columnNames, IEnumerable<string> row)
        {
            var rowObj = new JObject();
            for (int i = 0; i < row.Count(); i++)
            {
                rowObj.Add(columnNames.ElementAt(i), row.ElementAt(i));
            }
            return rowObj;
        }

        private JArray composeRecords()
        {
            var columnNames = mWorksheet.SelectRow(mConfig.PropertyRow);
            var records = new JArray();
            for (int rowNum = mConfig.StartRecordRow; rowNum <= mWorksheet.Dimension.End.Row; rowNum++)
            {
                records.Add(labelRow(columnNames, mWorksheet.SelectRow(rowNum)));
            }

            return records;
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

        private int getColumnOf(string value)
        {
            for (int c = 1; c <= mWorksheet.Dimension.End.Column; c++)
            {
                var cellValue = mWorksheet.Cells[mConfig.PropertyRow, c].Value;
                if (cellValue != null && cellValue.Equals(value))
                {
                    return c;
                }
            }

            throw new InvalidDataException("keyColumn 값과 일치하는 셀이 없습니다.");
        }

        private int getRowOf(int column, string value)
        {
            for (int r = mConfig.StartRecordRow; r <= mWorksheet.Dimension.End.Row; r++)
            {
                var cellValue = mWorksheet.Cells[r, column].Value;
                if (cellValue != null && cellValue.Equals(value))
                {
                    return r;
                }
            }

            throw new InvalidDataException("key 값과 일치하는 셀이 없습니다.");
        }

        private void MqttMsgPublishReceived(object sender, MqttMsgPublishEventArgs e)
        {
            var json = JObject.Parse(Encoding.UTF8.GetString(e.Message, 0, e.Message.Length));
            if (!isValid(json))
            {
                throw new InvalidDataException("레코드 전송에 대한 응답 형식이 올바르지 않습니다.");
            }

            int targetCol = getColumnOf(json.GetValue("keyColumn").ToString());
            int targetRow = getRowOf(targetCol, json.GetValue("key").ToString());

            Excel.Application xlApp;
            try
            {
                xlApp = (Excel.Application)System.Runtime.InteropServices.Marshal.GetActiveObject("Excel.Application");
            }
            catch (Exception ex)
            {
                if (ex.ToString().Contains("0x800401E3 (MK_E_UNAVAILABLE)"))
                {
                    xlApp = new Excel.Application();
                }
                else
                {
                    throw;
                }
            }
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(mConfig.ExcelFileName);
            Excel.Worksheet xlWorksheet = xlWorkbook.Worksheets[mConfig.ExcelSheetName];
            Excel.Range targetCell = xlWorksheet.Cells[targetRow, targetCol];
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
            xlWorkbook.Save();
        }
    }
}
