using System;

namespace CSharpLibraryForExcel
{
    public class Config
    {
        public string ExcelFileName { get; private set; }
        public string ExcelSheetName { get; private set; }
        public string BrokerHostName { get; private set; }
        public int BrokerPort { get; private set; }
        public string ClientId { get; private set; }
        public string Username { get; private set; }
        public string Password { get; private set; }
        public string Topic { get; private set; }
        public int PropertyRow { get; private set; }
        public int StartRecordRow { get; private set; }
        public int ChunkSize { get; private set; }

        public Config(
            string excelFileName,
            string excelSheetName,
            string brokerHostName,
            int brokerPort,
            string clientId,
            string username,
            string password,
            string topic,
            int propertyRow,
            int startRecordRow,
            int chunkSize)
        {
            if (string.IsNullOrWhiteSpace(excelFileName))
            {
                throw new ArgumentException($"'{nameof(excelFileName)}' cannot be null or white space.", nameof(excelFileName));
            }

            if (string.IsNullOrWhiteSpace(excelSheetName))
            {
                throw new ArgumentException($"'{nameof(excelSheetName)}' cannot be null or white space.", nameof(excelSheetName));
            }

            if (string.IsNullOrWhiteSpace(brokerHostName))
            {
                throw new ArgumentException($"'{nameof(brokerHostName)}' cannot be null or white space.", nameof(brokerHostName));
            }

            if (string.IsNullOrWhiteSpace(clientId))
            {
                throw new ArgumentException($"'{nameof(clientId)}' cannot be null or white space.", nameof(clientId));
            }

            if (string.IsNullOrWhiteSpace(username))
            {
                throw new ArgumentException($"'{nameof(username)}' cannot be null or white space.", nameof(username));
            }

            if (string.IsNullOrWhiteSpace(password))
            {
                throw new ArgumentException($"'{nameof(password)}' cannot be null or white space.", nameof(password));
            }

            if (string.IsNullOrWhiteSpace(topic))
            {
                throw new ArgumentException($"'{nameof(topic)}' cannot be null or white space.", nameof(topic));
            }

            if (propertyRow < 1)
            {
                throw new ArgumentException($"'{nameof(propertyRow)}' cannot be zero or nagative.", nameof(propertyRow));
            }

            if (startRecordRow < 1)
            {
                throw new ArgumentException($"'{nameof(startRecordRow)}' cannot be zero or nagative.", nameof(startRecordRow));
            }

            if (startRecordRow <= propertyRow)
            {
                throw new ArgumentException($"'{nameof(startRecordRow)}' cannot be less than or equal to '{nameof(propertyRow)}'.", nameof(startRecordRow));
            }

            if (chunkSize < 1)
            {
                throw new ArgumentException($"'{nameof(chunkSize)}' cannot be zero or nagative.", nameof(chunkSize));
            }

            ExcelFileName = excelFileName;
            ExcelSheetName = excelSheetName;
            BrokerHostName = brokerHostName;
            BrokerPort = brokerPort;
            ClientId = clientId;
            Username = username;
            Password = password;
            Topic = topic;
            PropertyRow = propertyRow;
            StartRecordRow = startRecordRow;
            ChunkSize = chunkSize;
        }
    }
}
