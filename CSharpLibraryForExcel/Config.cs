using System;

namespace CSharpLibraryForExcel
{
    public class Config
    {
        public string ExcelFileName { get; set; }
        public string ExcelSheetName { get; set; }
        public string BrokerHostName { get; set; }
        public int BrokerPort { get; set; }
        public string ClientId { get; set; }
        public string Username { get; set; }
        public string Password { get; set; }
        public string Topic { get; set; }
        public int PropertyRow { get; set; }
        public int StartRecordRow { get; set; }
        public int ChunkSize { get; set; }

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
