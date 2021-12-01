using Newtonsoft.Json.Linq;
using System;
using System.Diagnostics;
using System.Text;
using uPLibrary.Networking.M2Mqtt;
using uPLibrary.Networking.M2Mqtt.Messages;

namespace CSharpLibraryForExcel
{
    public class MqttHandler
    {
        private readonly Config mConfig;
        private readonly ExcelCOM mExcelCOM;
        private readonly MqttClient mMqttClient;

        public MqttHandler(Config config, ExcelCOM excelCOM)
        {
            mConfig = config;
            mExcelCOM = excelCOM;
            mMqttClient = new MqttClient(
                brokerHostName: mConfig.BrokerHostName,
                brokerPort: mConfig.BrokerPort,
                secure: false,
                caCert: null);

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
        }

        public void Publish(JObject msg)
        {
            mMqttClient.Publish(
                topic: mConfig.Topic,
                message: Encoding.UTF8.GetBytes(msg.ToString()),
                qosLevel: MqttMsgBase.QOS_LEVEL_AT_LEAST_ONCE,
                retain: false);
        }

        private void MqttMsgPublishReceived(object sender, MqttMsgPublishEventArgs e)
        {
            var json = JObject.Parse(Encoding.UTF8.GetString(e.Message, 0, e.Message.Length));
            
            if (!isValid(json))
            {
                Debug.WriteLine("invalid json");
                return;
            }

            int targetCol = mExcelCOM.GetColumnNumberOf(json["keyColumn"].ToString());
            if (targetCol == 0)
            {
                Debug.WriteLine("keyColumn error");
                return;
            }

            int targetRow = mExcelCOM.GetRowNumberOf(getKeyValue(json["key"]), targetCol);
            if (targetRow == 0)
            {
                Debug.WriteLine("key error");
                return;
            }

            if (json.GetValue("result").ToString().Equals("OK"))
            {
                mExcelCOM.MarkCellOK(targetRow, targetCol);
            }
            else if (json.GetValue("result").ToString().Equals("NK"))
            {
                mExcelCOM.MarkCellNK(targetRow, targetCol, json.GetValue("reason").ToString());
            }
            else
            {
                Debug.WriteLine("result error");
                return;
            }
        }

        private bool isValid(JObject response)
        {
            if (response.GetValue("keyColumn") == null) return false;
            if (response.GetValue("key") == null) return false;
            if (response.GetValue("result") == null) return false;
            if (response.GetValue("reason") == null) return false;

            return true;
        }

        private dynamic getKeyValue(JToken key)
        {
            dynamic result;
            if (key.Type == JTokenType.String)
            {
                result = (string)key;
            }
            else if (key.Type == JTokenType.Float)
            {
                result = (double)key;
            }
            else if (key.Type == JTokenType.Integer)
            {
                result = (int)key;
            }
            else
            {
                result = null;
            }

            return result;
        }
    }
}
