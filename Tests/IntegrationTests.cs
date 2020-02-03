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

            config.HistoryMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Åsgard molweight", Tag = "31AM0161V" });
            config.HistoryMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Statpipe molweight", Tag = "31AM0157V" });
            config.HistoryMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Statpipe mass flow x-over", Tag = "20FI7195SM" });
            config.HistoryMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Åsgard mass flow før x-over", Tag = "20FI5506" });


            config.AsgardMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Åsgard mass flow før x-over", Tag = "20FI5506" });
            config.AsgardMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Åsgard density før x-over", Tag = "15DY2038" });
            config.AsgardMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Åsgard mass flow x-over", Tag = "15FI5701SM" });
            config.AsgardMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Åsgard density Kårstø", Tag = "15DY2038" });
            config.AsgardMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Åsgard volume flow", Tag = "31FI2038V", Type = "double" });
            config.AsgardMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Åsgard transport time", Tag = "31KY2038", Type = "double" });

            config.AsgardMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Åsgard molweight", Tag = "31AM0161V" });

            config.AsgardComposition.Item.Add(new MixCalc.Component { Name = "Åsgard CO2", Id = 0, Tag = "31AI0161A_K", ScaleFactor = 0.01, WriteTag = "31AY0161A_K", Type = "double" });
            config.AsgardComposition.Item.Add(new MixCalc.Component { Name = "Åsgard N2", Id = 1, Tag = "31AI0161A_J", ScaleFactor = 0.01, WriteTag = "31AY0161A_J", Type = "double" });

            config.StatpipeMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "X-over status", Tag = "15HS0105" });

            config.StatpipeMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Statpipe mass flow x-over", Tag = "20FI7195SM" });
            config.StatpipeMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Statpipe mass flow x-over STP", Tag = "15FY0105" });
            config.StatpipeMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Statpipe density Kårstø", Tag = "31DY0004" });
            config.StatpipeMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Statpipe density Kalstø", Tag = "31DY0157" });
            config.StatpipeMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Statpipe volume flow", Tag = "31FI035V", Type = "double" });
            config.StatpipeMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Statpipe transport time", Tag = "31KY035", Type = "double" });

            config.StatpipeMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "Statpipe molweight", Tag = "31AM0157V" });
            config.StatpipeMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "T100 molweight", Tag = "15AM5626V" });
            config.StatpipeMeasurements.Item.Add(new MixCalc.TimeStampedMeasurement { Name = "T100 mass flow", Tag = "20FI7193SM" });

            config.T100Composition.Item.Add(new MixCalc.Component { Name = "T100 CO2", Id = 0, WriteTag = "15AY5626_K", Type = "double" });
            config.T100Composition.Item.Add(new MixCalc.Component { Name = "T100 N2", Id = 1, WriteTag = "15AY5626_J", Type = "double" });

            config.T400Composition.Item.Add(new MixCalc.Component { Name = "T400 CO2", Id = 0, WriteTag = "20AY7146_K", Type = "double" });
            config.T400Composition.Item.Add(new MixCalc.Component { Name = "T400 N2", Id = 1, WriteTag = "20AY7146_J", Type = "double" });

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
