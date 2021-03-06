﻿using System;
using System.Collections.Generic;
using System.Xml;
using System.Xml.Serialization;

namespace MixCalc
{
    [XmlRoot("configuration")]
    public class ConfigModel
    {
        public enum PressureUnit : int
        {
            barg = 0,
            bara = 1,
        }

        public enum TemperatureUnit : int
        {
            C = 0,
            K = 1,
        }

        [XmlElement]
        public string OpcUrl { get; set; }
        [XmlElement]
        public string OpcUser { get; set; }
        [XmlElement]
        public string OpcPassword { get; set; }

        [XmlElement]
        public bool ReadOnly { get; set; }

        [XmlElement]
        public MeasurementList HistoryMeasurements { get; set; } = new MeasurementList();

        [XmlElement]
        public MeasurementList AsgardMeasurements { get; set; } = new MeasurementList();

        [XmlElement]
        public CompositionList AsgardComposition { get; set; } = new CompositionList();

        [XmlElement]
        public MeasurementList StatpipeMeasurements { get; set; } = new MeasurementList();

        [XmlElement]
        public CompositionList StatpipeComposition { get; set; } = new CompositionList();

        [XmlElement]
        public CompositionList T400Composition { get; set; } = new CompositionList();

        [XmlElement]
        public CompositionList T100Composition { get; set; } = new CompositionList();

        [XmlElement]
        public MeasurementList Validation { get; set; } = new MeasurementList();

        [XmlElement]
        public Measurement WatchDog { get; set; } = new Measurement();

        public static ConfigModel ReadConfig(string file)
        {
            XmlReaderSettings readerSettings = new XmlReaderSettings
            {
                IgnoreComments = true,
                IgnoreProcessingInstructions = true,
                IgnoreWhitespace = true
            };

            XmlReader configFileReader = XmlReader.Create(file, readerSettings);
            XmlSerializer configSerializer = new XmlSerializer(typeof(ConfigModel));
            ConfigModel result = (ConfigModel)configSerializer.Deserialize(configFileReader);
            configFileReader.Close();

            return result;
        }
    }

    public class StreamList
    {
        public StreamList() { Item = new List<Stream>(); }
        [XmlElement("Stream")]
        public List<Stream> Item { get; }
    }

    public class Stream
    {
        [XmlAttribute]
        public string Name { get; set; }

        [XmlElement]
        public CompositionList Composition { get; set; } = new CompositionList();
    }

    public class MeasurementList
    {
        public MeasurementList() { Item = new List<TimeStampedMeasurement>(); }
        [XmlElement("Measurement")]
        public List<TimeStampedMeasurement> Item { get; }
        [XmlAttribute]
        public double MaxAge { get; set; }
    }

    public class CompositionList
    {
        public CompositionList() { Item = new List<Component>(); }
        [XmlElement("Component")]
        public List<Component> Item { get; }

        public double[] GetValues()
        {
            List<double> vs = new List<double>();

            foreach (var component in Item)
            {
                vs.Add(component.Value);
            }

            return vs.ToArray();
        }

        public double[] GetScaledValues()
        {
            List<double> vs = new List<double>();

            foreach (var component in Item)
            {
                vs.Add(component.GetScaledValue());
            }

            return vs.ToArray();
        }

        public Int32[] GetIds()
        {
            List<Int32> vs = new List<Int32>();

            foreach (var component in Item)
            {
                vs.Add(component.Id);
            }

            return vs.ToArray();
        }

        public string[] GetTags()
        {
            List<string> vs = new List<string>();

            foreach (var component in Item)
            {
                vs.Add(component.Tag);
            }

            return vs.ToArray();
        }
    }

    public class Component
    {
        [XmlAttribute]
        public string Name { get; set; }
        [XmlAttribute]
        public Int32 Id { get; set; }
        [XmlAttribute]
        public string Tag { get; set; }
        [XmlAttribute]
        public double ScaleFactor { get; set; }
        [XmlAttribute]
        public string WriteTag { get; set; }
        [XmlAttribute]
        public string Type { get; set; }

        [XmlIgnore]
        public double Value { get; set; }
        [XmlIgnore]
        public double WriteValue { get; set; }

        public double GetScaledValue()
        {
            return Value * ScaleFactor;
        }

        public double GetScaledWriteValue()
        {
            return WriteValue * ScaleFactor;
        }
    }

    public class Measurement
    {
        [XmlAttribute]
        public string Name { get; set; }
        [XmlAttribute]
        public string Tag { get; set; }
        [XmlAttribute]
        public string Type { get; set; }
        [XmlAttribute]
        public bool Output { get; set; }

        [XmlIgnore]
        public double Value { get; set; }

        public object GetTypedValue()
        {
            switch (Type)
            {
                case "single":
                    return Convert.ToSingle(Value);
                case "double":
                    return Convert.ToDouble(Value);
                case "bool":
                    if (Value < 0.5)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                default:
                    return Convert.ToDouble(Value);
            }
        }
    }

    public class PressureMeasurement : Measurement
    {
        [XmlAttribute]
        public ConfigModel.PressureUnit Unit { get; set; }

        public double GetUMRConverted()
        {
            // Convert from Unit to bar absolute
            const double stdAtm = 1.01325;
            double result = 0.0;
            switch (Unit)
            {
                case ConfigModel.PressureUnit.barg:
                    result = (Value + stdAtm);
                    break;
                case ConfigModel.PressureUnit.bara:
                    result = Value;
                    break;
                default:
                    break;
            }

            return (result);
        }

        public double GetUnitConverted()
        {
            // Convert from bar absolute to Unit
            const double stdAtm = 1.01325;
            double result = 0.0;
            switch (Unit)
            {
                case ConfigModel.PressureUnit.barg:
                    result = (Value - stdAtm);
                    break;
                case ConfigModel.PressureUnit.bara:
                    result = Value;
                    break;
                default:
                    break;
            }

            return (result);
        }
    }

    public class TemperatureMeasurement : Measurement
    {
        [XmlAttribute]
        public ConfigModel.TemperatureUnit Unit { get; set; }

        public double GetUMRConverted()
        {
            // Convert from Unit to K
            const double zeroCelsius = 273.15;
            double result = 0.0;
            switch (Unit)
            {
                case ConfigModel.TemperatureUnit.C:
                    result = Value + zeroCelsius;
                    break;
                case ConfigModel.TemperatureUnit.K:
                    result = Value;
                    break;
                default:
                    break;
            }

            return (result);
        }

        public double GetUnitConverted()
        {
            // Convert from K to Unit
            const double zeroCelsius = 273.15;
            double result = 0.0;
            switch (Unit)
            {
                case ConfigModel.TemperatureUnit.C:
                    result = Value - zeroCelsius;
                    break;
                case ConfigModel.TemperatureUnit.K:
                    result = Value;
                    break;
                default:
                    break;
            }

            return (result);
        }
    }

    public class TimeStampedMeasurement: Measurement
    {
        [XmlIgnore]
        public DateTime TimeStamp { get; set; }
    }
}
