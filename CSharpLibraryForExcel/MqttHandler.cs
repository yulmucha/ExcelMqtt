using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using uPLibrary.Networking.M2Mqtt;
using uPLibrary.Networking.M2Mqtt.Messages;

namespace CSharpLibraryForExcel
{
    public class MqttHandler
    {
        private readonly Config mConfig;
        private readonly MqttClient mMqttClient;

        public MqttHandler(Config config)
        {
            this.mConfig = config;
            this.mMqttClient = new MqttClient(
                brokerHostName: mConfig.BrokerHostName,
                brokerPort: mConfig.BrokerPort,
                secure: false,
                caCert: null);
        }

        public void Publish(JObject msg)
        {
            mMqttClient.Connect(
                clientId: mConfig.ClientId,
                username: mConfig.Username,
                password: mConfig.Password);

            if (!mMqttClient.IsConnected)
            {
                throw new Exception("MqttClient failed to connect.");
            }

            //mMqttClient.MqttMsgPublishReceived += MqttMsgPublishReceived;
            //mMqttClient.Subscribe(new string[] { mConfig.Topic + "/REPLY" }, new byte[] { MqttMsgBase.QOS_LEVEL_AT_LEAST_ONCE });

            mMqttClient.Publish(
                topic: mConfig.Topic,
                message: Encoding.UTF8.GetBytes(msg.ToString()),
                qosLevel: MqttMsgBase.QOS_LEVEL_AT_LEAST_ONCE,
                retain: false);
        }
    }
}
