using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using uPLibrary.Networking.M2Mqtt;
using uPLibrary.Networking.M2Mqtt.Messages;

namespace CSharpLibraryForExcel
{
    [ComVisible(true)]
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ExcelMqtt
    {
        private string mExcelFileName;
        private string mExcelSheetName;
        private string mBrokerHostName;
        private int mBrokerPort;
        private string mClientId;
        private string mUsername;
        private string mPassword;
        private string mTopic;
        private int mPropertyRow;
        private int mStartRecordRow;
        private int mChunkSize;
        private static MqttClient mMqttClient;

        [ComVisible(true)]
        public void SetExcelFileName(string excelFileName)
        {
            if (string.IsNullOrWhiteSpace(excelFileName))
            {
                return;
            }

            mExcelFileName = excelFileName;
        }

        [ComVisible(true)]
        public void SetExcelSheetName(string excelSheetName)
        {
            if (string.IsNullOrWhiteSpace(excelSheetName))
            {
                return;
            }

            mExcelSheetName = excelSheetName;
        }

        [ComVisible(true)]
        public void SetBrokerHostName(string brokerHostName)
        {
            if (string.IsNullOrWhiteSpace(brokerHostName))
            {
                return;
            }

            mBrokerHostName = brokerHostName;
        }

        [ComVisible(true)]
        public void SetBrokerPort(int brokerPort)
        {
            mBrokerPort = brokerPort;
        }

        [ComVisible(true)]
        public void SetClientId(string clientId)
        {
            if (string.IsNullOrWhiteSpace(clientId))
            {
                return;
            }

            mClientId = clientId;
        }

        [ComVisible(true)]
        public void SetUsername(string username)
        {
            if (string.IsNullOrWhiteSpace(username))
            {
                return;
            }

            mUsername = username;
        }

        [ComVisible(true)]
        public void SetPassword(string password)
        {
            mPassword = password;
        }

        [ComVisible(true)]
        public void SetTopic(string topic)
        {
            if (string.IsNullOrWhiteSpace(topic))
            {
                return;
            }

            mTopic = topic;
        }

        [ComVisible(true)]
        public void SetPropertyRow(int rowNumber)
        {
            mPropertyRow = rowNumber;
        }

        [ComVisible(true)]
        public void SetStartRecordRow(int rowNumber)
        {
            mStartRecordRow = rowNumber;
        }

        [ComVisible(true)]
        public void SetChunkSize(int size)
        {
            mChunkSize = size;
        }

        private ExcelWorksheet getSheet(ExcelPackage xlPackage)
        {
            if (xlPackage.Workbook.Worksheets.Count == 1)
            {
                return xlPackage.Workbook.Worksheets.First();
            }
            else
            {
                return xlPackage.Workbook.Worksheets
                    .Where(s => s.Name.Equals(mExcelSheetName, StringComparison.OrdinalIgnoreCase))
                    .Single();
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

        private JArray composeRecords(ExcelWorksheet sheet, int propertyRow, int startRecordRow)
        {
            var columnNames = sheet.SelectRow(propertyRow);
            var records = new JArray();
            for (int rowNum = startRecordRow; rowNum <= sheet.Dimension.End.Row; rowNum++)
            {
                records.Add(labelRow(columnNames, sheet.SelectRow(rowNum)));
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

        private void connect()
        {
            mMqttClient = new MqttClient(
                    brokerHostName: mBrokerHostName,
                    brokerPort: mBrokerPort,
                    secure: false,
                    caCert: null);

            mMqttClient.Connect(
                clientId: mClientId,
                username: mUsername,
                password: mPassword);

            if (!mMqttClient.IsConnected)
            {
                Debug.WriteLine("connection falied");
                return;
            }

            mMqttClient.MqttMsgPublishReceived += MqttMsgPublishReceived;
            mMqttClient.Subscribe(new string[] { mTopic }, new byte[] { MqttMsgBase.QOS_LEVEL_AT_MOST_ONCE });
        }

        private void publish(JObject msg)
        {
            mMqttClient.Publish(
                topic: mTopic,
                message: Encoding.UTF8.GetBytes(msg.ToString()),
                qosLevel: MqttMsgBase.QOS_LEVEL_AT_LEAST_ONCE,
                retain: false);
        }

        [ComVisible(true)]
        public void Publish()
        {
            var messages = new List<JObject>();

            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(mExcelFileName)))
            {
                ExcelWorksheet sheet = getSheet(xlPackage);

                JArray records = composeRecords(sheet, mPropertyRow, mStartRecordRow);

                List<JArray> chunks = chunkRecords(records, mChunkSize);

                int totalRecords = sheet.Dimension.End.Row - mStartRecordRow + 1;
                messages = chunks.ToMessage(
                    totalRecords: totalRecords,
                    totalColumns: sheet.Dimension.End.Column,
                    totalChunks: (int)Math.Ceiling((double)totalRecords / mChunkSize),
                    chunkSize: mChunkSize);
            }

            connect();

            messages.ForEach(o => publish(o));
        }

        private static void MqttMsgPublishReceived(object sender, MqttMsgPublishEventArgs e)
        {
            var json = JObject.Parse(Encoding.UTF8.GetString(e.Message, 0, e.Message.Length));
            if (json.GetValue("chunkSequence").Equals(json.GetValue("chunks")))
            {
                mMqttClient.Disconnect();
            }
        }
    }
}
