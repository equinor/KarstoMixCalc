using Opc.Ua;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Timers;

namespace MixCalc
{
    class MixCalcService : IDisposable
    {
        private readonly Timer timer;
        private readonly ConfigModel config;
        private readonly OpcClient opcClient;
        private static readonly NLog.Logger logger = NLog.LogManager.GetCurrentClassLogger();
        private readonly object WorkerLock = new object();

        public MixCalcService()
        {
            logger.Info("Initializing MixCalcService.");
            string ConfigFile = AppDomain.CurrentDomain.BaseDirectory.ToString(CultureInfo.InvariantCulture) + "MixCalc.config";
            logger.Debug(CultureInfo.InvariantCulture, "Reading configuration from {0}.", ConfigFile);
            try
            {
                config = ConfigModel.ReadConfig(ConfigFile);
            }
            catch (Exception e)
            {
                logger.Fatal(e, "Failed to read configuration.");
                throw;
            }

            timer = new Timer(20_000.0) { AutoReset = true, SynchronizingObject = null };
            timer.Elapsed += Worker;

            opcClient = new OpcClient(config.OpcUrl, config.OpcUser, config.OpcPassword);
        }

        public void Dispose()
        {
            opcClient.Dispose();
            timer.Dispose();
        }

        public async void Start()
        {
            logger.Info("Starting service.");
            await opcClient.Connect().ConfigureAwait(false);
            timer.Start();
        }

        private void Worker(object sender, ElapsedEventArgs ea)
        {
            lock (WorkerLock)
            {
                ReadFromOPC();
                StoreHistoryMeasurements();
                CalculateVolumeFlow();
                CalculateDelays();
                WriteToOPC();
            }
        }

        private void CalculateVolumeFlow()
        {
            foreach (var item in config.AsgardMeasurements.Item)
            {
                if (item.Name == "Åsgard Volume flow")
                {
                    item.Value = CalculateAsgardVolumeFlow();
                }
            }

            foreach (var item in config.StatpipeMeasurements.Item)
            {
                if (item.Name == "Statpipe Volume flow")
                {
                    item.Value = CalculateStatpipeVolumeFlow();
                }
            }
        }

        private void CalculateDelays()
        {
            double asgardPipeVolume = 18085.0;
            double statpipePipeVolume = 7901.0;
            string asgardTag = "";
            string statpipeTag = "";

            foreach (var item in config.AsgardMeasurements.Item)
            {
                if (item.Name == "Åsgard Volume flow")
                {
                    asgardTag = item.Tag;
                }
            }

            foreach (var item in config.StatpipeMeasurements.Item)
            {
                if (item.Name == "Statpipe Volume flow")
                {
                    statpipeTag = item.Tag;
                }
            }

            TimeSpan asgardDelay = CalculateDelay(asgardTag, asgardPipeVolume);
            logger.Debug(CultureInfo.InvariantCulture, "Åsgard delay: {0} h", asgardDelay.TotalHours);

            TimeSpan statpipeDelay = CalculateDelay(statpipeTag, statpipePipeVolume);
            logger.Debug(CultureInfo.InvariantCulture, "Statpipe delay: {0} h", statpipeDelay.TotalHours);

            foreach (var item in config.AsgardMeasurements.Item)
            {
                if (item.Name == "Åsgard Transport time")
                {
                    item.Value = asgardDelay.TotalHours;
                }
            }

            foreach (var item in config.StatpipeMeasurements.Item)
            {
                if (item.Name == "Statpipe Transport time")
                {
                    item.Value = statpipeDelay.TotalHours;
                }
            }
        }

        private TimeSpan CalculateDelay(string Tag, double PipeVolume)
        {
            DateTime timeStamp = DateTime.Now;
            int i = 1;
            DateTime t0 = DateTime.Now;
            double volume = 0.0;
            foreach (var m in DataAccess.GetValueSet(Tag))
            {
                if (i == 1)
                {
                    t0 = m.TimeStamp;
                    i++;
                    volume += m.Value;
                    continue;
                }

                volume += m.Value;
                double accumulatedVolume = (volume / (double)i) * (t0 - m.TimeStamp).TotalHours;

                logger.Debug(CultureInfo.InvariantCulture, "{0} accumulated volume: {1}, delay: {2}", Tag, accumulatedVolume, (t0 - m.TimeStamp).TotalHours);

                if (accumulatedVolume > PipeVolume)
                {
                    timeStamp = m.TimeStamp;
                    break;
                }
                i++;
            }

            return (t0 - timeStamp);
        }

        private void ReadFromOPC()
        {
            NodeIdCollection nodes = new NodeIdCollection();
            List<Type> types = new List<Type>();
            List<object> result = new List<object>();
            List<ServiceResult> errors = new List<ServiceResult>();

            // Make a list of all the OPC item that we want to read
            foreach (var item in config.HistoryMeasurements.Item)
            {
                nodes.Add(item.Tag); types.Add(typeof(object));
            }

            foreach (var item in config.AsgardMeasurements.Item)
            {
                nodes.Add(item.Tag); types.Add(typeof(object));
            }

            foreach (var item in config.StatpipeMeasurements.Item)
            {
                nodes.Add(item.Tag); types.Add(typeof(object));
            }

            foreach (var item in nodes)
            {
                logger.Debug(CultureInfo.InvariantCulture, "Item to read: \"{0}\"", item.ToString());
            }

            // Read all of the items
            try
            {
                opcClient.OpcSession.ReadValues(nodes, types, out result, out errors);
            }
            catch (Exception e)
            {
                logger.Error(e, "Error reading values from OPC.");
            }

            for (int n = 0; n < result.Count; n++)
            {
                logger.Debug(CultureInfo.InvariantCulture, "Item: \"{0}\" Value: \"{1}\" Status: \"{2}\"",
                    nodes[n].ToString(), result[n], errors[n].StatusCode.ToString());
            }

            int it = 0;
            foreach (var meas in config.HistoryMeasurements.Item)
            {
                meas.Value = Convert.ToDouble(result[it++], CultureInfo.InvariantCulture);
                meas.TimeStamp = DateTime.Now;
                logger.Debug(CultureInfo.InvariantCulture,
                    "Measurement: \"{0}\" Value: {1} TimeStamp: {2} Tag: \"{3}\"",
                    meas.Name, meas.Value, meas.TimeStamp.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), meas.Tag);
            }

            foreach (var meas in config.AsgardMeasurements.Item)
            {
                meas.Value = Convert.ToDouble(result[it++], CultureInfo.InvariantCulture);
                meas.TimeStamp = DateTime.Now;
                logger.Debug(CultureInfo.InvariantCulture,
                    "Measurement: \"{0}\" Value: {1} TimeStamp: {2} Tag: \"{3}\"",
                    meas.Name, meas.Value, meas.TimeStamp.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), meas.Tag);
            }

            foreach (var meas in config.StatpipeMeasurements.Item)
            {
                meas.Value = Convert.ToDouble(result[it++], CultureInfo.InvariantCulture);
                meas.TimeStamp = DateTime.Now;
                logger.Debug(CultureInfo.InvariantCulture,
                    "Measurement: \"{0}\" Value: {1} TimeStamp: {2} Tag: \"{3}\"",
                    meas.Name, meas.Value, meas.TimeStamp.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), meas.Tag);
            }
        }

        private void StoreHistoryMeasurements()
        {
            try
            {
                int storedRows = MixCalc.DataAccess.StoreValue(config.HistoryMeasurements.Item);
                logger.Debug(CultureInfo.InvariantCulture,
                    "Wrote {0} values to History database.", storedRows);
                int clearedRows = MixCalc.DataAccess.ClearHistory(DateTime.Now.AddHours(-config.HistoryMeasurements.MaxAge));
                if (clearedRows > 0)
                {
                    logger.Debug(CultureInfo.InvariantCulture,
                        "Cleared {0} values that were older than {1} hours from History database.", clearedRows, config.HistoryMeasurements.MaxAge);
                }
            }
            catch (Exception e)
            {
                logger.Error(e, "Error writing to History database");
            }
        }

        private double CalculateAsgardVolumeFlow()
        {
            double massFlowBeforeXover = 0.0;
            double densityBeforeXover = 0.0;
            double massFlowXover = 0.0;
            double densityKarsto = 0.0;
            double densityKalsto = 0.0;

            foreach (var item in config.AsgardMeasurements.Item)
            {
                switch (item.Name)
                {
                    case "Åsgard Mass flow før x-over":
                        massFlowBeforeXover = item.Value;
                        break;
                    case "Åsgard Density før x-over":
                        densityBeforeXover = item.Value;
                        break;
                    case "Åsgard Mass flow x-over":
                        massFlowXover = item.Value;
                        break;
                    case "Åsgard Density Kårstø":
                        densityKarsto = item.Value;
                        break;
                    case "Åsgard Density Kalstø":
                        densityKalsto = item.Value;
                        break;
                    default:
                        break;
                }
            }

            // Convert from mass flow to diff pressure [mbar]
            double dp = Math.Pow(massFlowBeforeXover * Math.Sqrt(350.0) / 3105.0, 2.0);
            logger.Debug(CultureInfo.InvariantCulture, "Åsgard diff pressure before x-over {0} mbar", dp);

            // Calculate mass flow before x-over [t/h]
            massFlowBeforeXover = 1312.0 * Math.Sqrt(densityBeforeXover * dp * 100.0) / 1000.0;
            logger.Debug(CultureInfo.InvariantCulture, "Åsgard corrected mass flow before x-over {0} t/h", massFlowBeforeXover);

            // Calculate Åsgard mass flow [t/h]
            double asgardMassFlow = massFlowBeforeXover + massFlowXover;
            logger.Debug(CultureInfo.InvariantCulture, "Åsgard mass flow {0} t/h", asgardMassFlow);

            // Calculate average density
            double density = (densityKalsto + densityKarsto) / 2.0;
            // Calculate Åsgard volume flow [m³/h]
            double asgardVolumeFlow = 1000.0 * asgardMassFlow / density;
            logger.Debug(CultureInfo.InvariantCulture, "Åsgard volume flow {0} m³/h", asgardVolumeFlow);

            return asgardVolumeFlow;
        }

        private double CalculateStatpipeVolumeFlow()
        {
            double massFlowXover = 0.0;
            double massFlowXoverSTP = 0.0;
            double densityKarsto = 0.0;
            double densityKalsto = 0.0;

            foreach (var item in config.StatpipeMeasurements.Item)
            {
                switch (item.Name)
                {
                    case "Statpipe Mass flow x-over":
                        massFlowXover = item.Value;
                        break;
                    case "Statpipe Mass flow x-over STP":
                        massFlowXoverSTP = item.Value;
                        break;
                    case "Statpipe Density Kårstø":
                        densityKarsto = item.Value;
                        break;
                    case "Statpipe Density Kalstø":
                        densityKalsto = item.Value;
                        break;
                    default:
                        break;
                }
            }

            // Sum mass flows [t/h]
            if (massFlowXoverSTP < 0.0)
            {
                massFlowXoverSTP = 0.0;
            }
            double massFlow = massFlowXover + massFlowXoverSTP;

            // Calculate average density
            double density = (densityKalsto + densityKarsto) / 2.0;
            // Calculate Statpipe volume flow [m³/h]
            double statpipeVolumeFlow = 1000.0 * massFlow / density;
            logger.Debug(CultureInfo.InvariantCulture, "Statpipe volume flow {0} m³/h", statpipeVolumeFlow);

            return statpipeVolumeFlow;
        }

        private void WriteToOPC()
        {
            // Make a list of all the OPC item that we want to write
            WriteValueCollection wvc = new WriteValueCollection();

            foreach (var item in config.AsgardMeasurements.Item)
            {
                if (!string.IsNullOrEmpty(item.Type))
                {
                    wvc.Add(new WriteValue()
                    {
                        NodeId = item.Tag,
                        AttributeId = Attributes.Value,
                        Value = new DataValue { Value = item.GetTypedValue() }
                    });
                }
            }

            foreach (var item in config.StatpipeMeasurements.Item)
            {
                if (!string.IsNullOrEmpty(item.Type))
                {
                    wvc.Add(new WriteValue()
                    {
                        NodeId = item.Tag,
                        AttributeId = Attributes.Value,
                        Value = new DataValue { Value = item.GetTypedValue() }
                    });
                }
            }

            foreach (var item in wvc)
            {
                logger.Debug(CultureInfo.InvariantCulture, "Item to write: \"{0}\" Value: {1}",
                    item.NodeId.ToString(),
                    item.Value.Value);
            }

            try
            {
                opcClient.OpcSession.Write(null, wvc, out StatusCodeCollection results, out DiagnosticInfoCollection diagnosticInfos);

                for (int i = 0; i < results.Count; i++)
                {
                    if (results[i].Code != 0)
                    {
                        logger.Error(CultureInfo.InvariantCulture, "Write result: \"{0}\" Tag: \"{1}\" Value: \"{2}\" Type: \"{3}\"",
                            results[i].ToString(), wvc[i].NodeId, wvc[i].Value.Value, wvc[i].Value.Value.GetType().ToString());
                    }

                }
            }
            catch (Exception e)
            {
                logger.Error(e, "Error writing OPC items");
            }
        }

        public void Stop()
        {
            logger.Info("Stop service command received.");
            timer.Stop();
            logger.Info("Waiting for current worker.");
            lock (WorkerLock)
            {
                logger.Info("Worker is done.");
                logger.Info("Disconnecting from OPC server.");
                opcClient.DisConnect();
            }
            logger.Info("Stopping service.");
        }
    }
}
