using Newtonsoft.Json.Linq;
using OfficeOpenXml;
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
        public void Publish()
        {
            using (ExcelPackage xlPackage = new ExcelPackage(new FileInfo(mExcelFileName)))
            {
                ExcelWorksheet sheet = xlPackage.Workbook.Worksheets.First();
                int totalRows = sheet.Dimension.End.Row - mStartRecordRow + 1;
                int totalColumns = sheet.Dimension.End.Column;

                var columnNames = sheet
                    .Cells[mPropertyRow, 1, mPropertyRow, totalColumns]
                    .Select(c => c.Value == null ? string.Empty : c.Value.ToString());

                var rows = new JArray();
                for (int rowNum = mStartRecordRow; rowNum <= sheet.Dimension.End.Row; rowNum++)
                {
                    var row = sheet
                        .Cells[rowNum, 1, rowNum, totalColumns]
                        .Select(c => c.Value == null ? string.Empty : c.Value.ToString());
                    var rowObj = new JObject();
                    for (int i = 0; i < row.Count(); i++)
                    {
                        rowObj.Add(columnNames.ElementAt(i), row.ElementAt(i));
                    }
                    rows.Add(rowObj);
                }

                var msg = new JObject();
                msg.Add("rows", totalRows);
                msg.Add("columns", totalColumns);
                msg.Add("data", rows);

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
        }
    }
}
