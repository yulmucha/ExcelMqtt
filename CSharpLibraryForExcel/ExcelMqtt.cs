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

        private void publish(JObject msg)
        {
            var mqttClient = new MqttClient(
                    brokerHostName: mBrokerHostName,
                    brokerPort: mBrokerPort,
                    secure: false,
                    caCert: null);

            mqttClient.Connect(
                clientId: mClientId,
                username: mUsername,
                password: mPassword);

            if (!mqttClient.IsConnected)
            {
                Debug.WriteLine("connection falied");
                return;
            }

            mqttClient.Publish(
                topic: mTopic,
                message: Encoding.UTF8.GetBytes(msg.ToString()),
                qosLevel: MqttMsgBase.QOS_LEVEL_AT_LEAST_ONCE,
                retain: false);
        }

        [ComVisible(true)]
        public void Publish()
        {
            var msg = new JObject();

            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(mExcelFileName)))
            {
                ExcelWorksheet sheet = getSheet(xlPackage);
                int totalRecords = sheet.Dimension.End.Row - mStartRecordRow + 1;
                int totalColumns = sheet.Dimension.End.Column;

                var columnNames = sheet.SelectRow(mPropertyRow);

                var rows = new JArray();
                for (int rowNum = mStartRecordRow; rowNum <= sheet.Dimension.End.Row; rowNum++)
                {
                    rows.Add(labelRow(columnNames, sheet.SelectRow(rowNum)));
                }

                msg.Add("rows", totalRecords);
                msg.Add("columns", totalColumns);
                msg.Add("data", rows);
            }

            publish(msg);
        }
    }
}
