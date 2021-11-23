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
using Excel = Microsoft.Office.Interop.Excel;

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

        [ComVisible(true)]
        public void SetExcelFileName(string excelFileName)
        {
            mExcelFileName = excelFileName;
        }

        [ComVisible(true)]
        public void SetExcelSheetName(string excelSheetName)
        {
            mExcelSheetName = excelSheetName;
        }

        [ComVisible(true)]
        public void SetBrokerHostName(string brokerHostName)
        {
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
            mClientId = clientId;
        }

        [ComVisible(true)]
        public void SetUsername(string username)
        {
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

        [ComVisible(true)]
        public void Publish()
        {
            var config = new Config(
                excelFileName: mExcelFileName,
                excelSheetName: mExcelSheetName,
                brokerHostName: mBrokerHostName,
                brokerPort: mBrokerPort,
                clientId: mClientId,
                username: mUsername,
                password: mPassword,
                topic: mTopic,
                propertyRow: mPropertyRow,
                startRecordRow: mStartRecordRow,
                chunkSize: mChunkSize);

            var handler = new ExcelMqttHandler(config);
            handler.Publish();
        }
    }
}
