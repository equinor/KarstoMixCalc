﻿using Opc.Ua;
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
                WriteToOPC();
            }
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

        private void CalculateAsgardVolumeFlow()
        {
            double massFlowBeforeXover = 0.0;
            double densityBeforeXover = 0.0;
            double massFlowXover = 0.0;
            double density = 0.0;

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
                    case "Åsgard Mass flow x-over AT->ST":
                        massFlowXover = item.Value;
                        break;
                    case "Åsgard Density":
                        density = item.Value;
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

            // Calculate Åsgard volume flow [m³/h]
            double asgardVolumeFlow = 1000.0 * asgardMassFlow / density;
            logger.Debug(CultureInfo.InvariantCulture, "Åsgard volume flow {0} m³/h", dp);
        }

        private void WriteToOPC()
        {
            // Make a list of all the OPC item that we want to write
            WriteValueCollection wvc = new WriteValueCollection();

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
