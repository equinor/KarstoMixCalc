using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Xml;
using System.Xml.Serialization;

namespace Tests
{
    [TestClass]
    public class IntegrationTests
    {
        [TestMethod]
        public void GenerateAndReadConfigModel()
        {
            // Stream 0
            MixCalc.ConfigModel config = new MixCalc.ConfigModel
            {
                OpcUrl = "opc.tcp://localhost:62548/Quickstarts/DataAccessServer",
                OpcUser = "user",
                OpcPassword = "password",
            };
            config.HistoryMeasurements.MaxAge = 2.0 / 60.0;
            config.HistoryMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Statpipe CO2", Tag = "31AI0157A_K" });
            config.HistoryMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Statpipe N2", Tag = "31AI0157A_J" });

            config.HistoryMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Åsgard CO2", Tag = "31AI0161A_K" });
            config.HistoryMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Åsgard N2", Tag = "31AI0161A_J" });

            XmlWriterSettings writerSettings = new XmlWriterSettings
            {
                Indent = true,
            };
            XmlWriter writer = XmlWriter.Create("MixCalc.config", writerSettings);
            XmlSerializer configSerializer = new XmlSerializer(typeof(MixCalc.ConfigModel));
            configSerializer.Serialize(writer, config);
            writer.Close();

            string file = AppDomain.CurrentDomain.BaseDirectory.ToString(CultureInfo.InvariantCulture) + "\\MixCalc.config";
            MixCalc.ConfigModel readConfig = MixCalc.ConfigModel.ReadConfig(file);

            for (int i = 0; i < config.HistoryMeasurements.Item.Count; i++)
            {
                Assert.AreEqual(config.HistoryMeasurements.Item[i].Name, readConfig.HistoryMeasurements.Item[i].Name);
                Assert.AreEqual(config.HistoryMeasurements.Item[i].Tag, readConfig.HistoryMeasurements.Item[i].Tag);
            }
        }

    }
}
