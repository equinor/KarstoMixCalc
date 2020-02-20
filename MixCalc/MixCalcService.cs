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

        private double AsgardMolFlow = 0.0;
        private double StatpipeMolFlow = 0.0;
        private double T100MolFlow = 0.0;
        private double StatpipeXoverMolFlow = 0.0;

        // status 0.0 is OK, 1.0 is Bad
        private double Status = 0.0;

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

            config.WatchDog.Value = 0.0;
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
                try
                {
                    if (config.WatchDog.Value > 100.0 || config.WatchDog.Value < 0.0)
                    {
                        config.WatchDog.Value = 0.0;
                    }
                    logger.Debug(CultureInfo.InvariantCulture, "WatchDog: {0}", config.WatchDog.Value);

                    ReadFromOPC();
                    CalculateVolumeFlow();
                    StoreHistoryMeasurements();
                    CalculateDelays();
                    ReadDelayedComposition();
                    CalculateMix();
                    Validate();
                    WriteToOPC();

                    Status = 0.0;
                    config.WatchDog.Value += 1.0;
                }
                catch (Exception e)
                {
                    Status = 1.0;
                    logger.Error(e, "Error in Worker.");
                }
            }
        }

        private static double CheckGcStatus(List<TimeStampedMeasurement> gc)
        {
            // 0 means good, 1 means bad
            double gcStatus = 1.0;
            foreach (var item in gc)
            {
                logger.Debug(CultureInfo.InvariantCulture, "Validation: {0}: {1}", item.Name, item.Value);
                // status value that is read from GC, 1000 means OK
                if (item.Value > 999.9) gcStatus = 0.0;
            }
            return gcStatus;
        }

        private void Validate()
        {
            // GC status
            if (CheckGcStatus(config.Validation.Item.FindAll(x => x.Name.StartsWith("GC 15AI2038 status", StringComparison.InvariantCulture))) > 0.5) Status = 1.0;
            if (CheckGcStatus(config.Validation.Item.FindAll(x => x.Name.StartsWith("GC 15AM5626 status", StringComparison.InvariantCulture))) > 0.5) Status = 1.0;
            if (CheckGcStatus(config.Validation.Item.FindAll(x => x.Name.StartsWith("GC 15AM0004 status", StringComparison.InvariantCulture))) > 0.5) Status = 1.0;
            if (CheckGcStatus(config.Validation.Item.FindAll(x => x.Name.StartsWith("GC Kalstø status", StringComparison.InvariantCulture))) > 0.5) Status = 1.0;

            // Mol flows
            if (double.IsNaN(AsgardMolFlow) || double.IsNaN(StatpipeMolFlow) || double.IsNaN(StatpipeXoverMolFlow) || double.IsNaN(T100MolFlow))
            {
                Status = 1.0;
            }

            // Volume flows
            if (double.IsNaN(config.AsgardMeasurements.Item.Find(x => x.Name.Contains("Åsgard volume flow")).Value) ||
                double.IsNaN(config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("Statpipe volume flow")).Value))
            {
                Status = 1.0;
            }

            // Transport times
            double ttAsgard = config.AsgardMeasurements.Item.Find(x => x.Name.Contains("Åsgard transport time")).Value;
            double ttStatpipe = config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("Statpipe transport time")).Value;
            if (double.IsNaN(ttAsgard) || double.IsNaN(ttStatpipe) || ttAsgard < 0.1 || ttStatpipe < 0.1)
            {
                Status = 1.0;
            }

            // Calculated delayed compositions
            foreach (var item in config.AsgardComposition.Item)
            {
                if (double.IsNaN(item.WriteValue))
                {
                    Status = 1.0;
                }
            }
            foreach (var item in config.StatpipeComposition.Item)
            {
                if (double.IsNaN(item.WriteValue))
                {
                    Status = 1.0;
                }
            }

            config.Validation.Item.Find(x => x.Name.Contains("PhaseOpt status")).Value = Status;
            logger.Debug(CultureInfo.InvariantCulture, "Validation: status {0}", Status);
        }

        private void CalculateMix()
        {
            // Calculate mol flows
            var asgardDelay = config.AsgardMeasurements.Item.Find(x => x.Name.Contains("Åsgard transport time"));
            var asgardFlow = DataAccess.GetValue(new List<string> { config.HistoryMeasurements.Item.Find(x => x.Name.Contains("Åsgard corrected mass flow før x-over")).Tag },
                DateTime.Now.AddHours(-asgardDelay.Value));
            var asgardMolWeight = DataAccess.GetValue(new List<string> { config.HistoryMeasurements.Item.Find(x => x.Name.Contains("Åsgard molweight")).Tag },
                DateTime.Now.AddHours(-asgardDelay.Value));

            AsgardMolFlow = asgardFlow[0] * 1000.0 / asgardMolWeight[0]; // [mol/h]
            logger.Debug(CultureInfo.InvariantCulture, "Åsgard mol flow: {0} mol/h", AsgardMolFlow);

            var statpipeDelay = config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("Statpipe transport time"));
            var statpipeFlow = DataAccess.GetValue(new List<string> { config.HistoryMeasurements.Item.Find(x => x.Name.Contains("Statpipe mass flow x-over")).Tag },
                DateTime.Now.AddHours(-statpipeDelay.Value));
            var statpipeMolWeight = DataAccess.GetValue(new List<string> { config.HistoryMeasurements.Item.Find(x => x.Name.Contains("Statpipe molweight")).Tag },
                DateTime.Now.AddHours(-statpipeDelay.Value));

            StatpipeMolFlow = statpipeFlow[0] * 1000.0 / statpipeMolWeight[0]; // [mol/h]
            logger.Debug(CultureInfo.InvariantCulture, "Statpipe mol flow: {0} mol/h", StatpipeMolFlow);

            T100MolFlow = config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("T100 mass flow")).Value * 1000.0 /
                config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("T100 molweight")).Value; // [mol/h]
            logger.Debug(CultureInfo.InvariantCulture, "T100 mol flow: {0} mol/h", T100MolFlow);

            if (config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("X-over status")).Value > 0.5)
            {
                StatpipeXoverMolFlow = config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("Statpipe mass flow x-over STP")).Value * 1000.0 / statpipeMolWeight[0];
                logger.Debug(CultureInfo.InvariantCulture, "X-over position: Statpipe");
            }
            else
            {
                StatpipeXoverMolFlow = config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("Statpipe mass flow x-over STP")).Value * 1000.0 / asgardMolWeight[0];
                logger.Debug(CultureInfo.InvariantCulture, "X-over position: Åsgard");
            }
            if (StatpipeXoverMolFlow < 0.0)
            {
                StatpipeXoverMolFlow = 0.0;
            }
            logger.Debug(CultureInfo.InvariantCulture, "X-over mol flow: {0} mol/h", StatpipeXoverMolFlow);

            // Calculate mixed compositions
            List<Component> asgardComponentFlow = new List<Component>();
            double asgardComponentFlowSum = 0.0;
            foreach (var item in config.AsgardComposition.Item)
            {
                asgardComponentFlow.Add(new Component() { Id = item.Id, WriteTag = item.WriteTag, WriteValue = AsgardMolFlow * item.GetScaledWriteValue() });
                asgardComponentFlowSum += AsgardMolFlow * item.GetScaledWriteValue();
                logger.Debug(CultureInfo.InvariantCulture, "Åsgard component flow \"{0}\": {1}", item.Name, AsgardMolFlow * item.GetScaledWriteValue());
            }
            logger.Debug(CultureInfo.InvariantCulture, "Åsgard component flow sum: {0}", asgardComponentFlowSum);

            List<Component> statpipeComponentFlow = new List<Component>();
            double statpipeComponentFlowSum = 0.0;
            foreach (var item in config.StatpipeComposition.Item)
            {
                statpipeComponentFlow.Add(new Component() { Id = item.Id, WriteTag = item.WriteTag, WriteValue = StatpipeMolFlow * item.GetScaledWriteValue() });
                statpipeComponentFlowSum += StatpipeMolFlow * item.GetScaledWriteValue();
                logger.Debug(CultureInfo.InvariantCulture, "Statpipe component flow \"{0}\": {1}", item.Name, StatpipeMolFlow * item.GetScaledWriteValue());
            }
            logger.Debug(CultureInfo.InvariantCulture, "Statpipe component flow sum: {0}", statpipeComponentFlowSum);

            if (asgardComponentFlowSum + statpipeComponentFlowSum > 0.0)
            {
                foreach (var item in config.T400Composition.Item)
                {
                    item.WriteValue = (asgardComponentFlow.Find(x => x.Id == item.Id).WriteValue +
                        statpipeComponentFlow.Find(x => x.Id == item.Id).WriteValue) /
                        (asgardComponentFlowSum + statpipeComponentFlowSum) * 100.0;
                }
            }

            List<Component> xOverComponentFlow = new List<Component>();
            double xOverComponentFlowSum = 0.0;
            if (config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("X-over status")).Value > 0.5)
            {
                foreach (var item in config.StatpipeComposition.Item)
                {
                    xOverComponentFlow.Add(new Component() { Id = item.Id, WriteTag = item.WriteTag, WriteValue = StatpipeXoverMolFlow * item.GetScaledWriteValue() });
                    xOverComponentFlowSum += StatpipeXoverMolFlow * item.GetScaledWriteValue();
                }
            }
            else
            {
                foreach (var item in config.AsgardComposition.Item)
                {
                    xOverComponentFlow.Add(new Component() { Id = item.Id, WriteTag = item.WriteTag, WriteValue = StatpipeXoverMolFlow * item.GetScaledWriteValue() });
                    xOverComponentFlowSum += StatpipeXoverMolFlow * item.GetScaledWriteValue();
                }
            }

            foreach (var item in xOverComponentFlow)
            {
                logger.Debug(CultureInfo.InvariantCulture, "X-over component flow \"{0}\": {1}", item.Name, item.WriteValue);
            }
            logger.Debug(CultureInfo.InvariantCulture, "X-over component flow sum: {0}", xOverComponentFlowSum);

            List<Component> dixoComponentFlow = new List<Component>();
            double dixoComponentFlowSum = 0.0;
            foreach (var item in config.T400Composition.Item)
            {
                dixoComponentFlow.Add(new Component() { Id = item.Id, WriteTag = item.WriteTag, WriteValue = T100MolFlow * item.GetScaledWriteValue() });
                dixoComponentFlowSum += T100MolFlow * item.GetScaledWriteValue();
                logger.Debug(CultureInfo.InvariantCulture, "Dixo component flow \"{0}\": {1}", item.Name, T100MolFlow * item.GetScaledWriteValue());
            }
            logger.Debug(CultureInfo.InvariantCulture, "Dixo component flow sum: {0}", dixoComponentFlowSum);

            if (xOverComponentFlowSum + dixoComponentFlowSum > 0.0)
            {
                foreach (var item in config.T100Composition.Item)
                {
                    item.WriteValue = (xOverComponentFlow.Find(x => x.Id == item.Id).WriteValue +
                        dixoComponentFlow.Find(x => x.Id == item.Id).WriteValue) /
                        (xOverComponentFlowSum + dixoComponentFlowSum) * 100.0;
                }
            }
        }
        private void ReadDelayedComposition()
        {
            var asgardDelay = config.AsgardMeasurements.Item.Find(x => x.Name.Contains("Åsgard transport time"));
            var asgardValues = DataAccess.GetValue(new List<string>(config.AsgardComposition.GetTags()), DateTime.Now.AddHours(-asgardDelay.Value));

            int i = 0;
            foreach (var item in asgardValues)
            {
                config.AsgardComposition.Item[i].WriteValue = item;
                logger.Debug(CultureInfo.InvariantCulture, "Åsgard delayed composition tag {0}, value {1}",
                    config.AsgardComposition.GetTags()[i], config.AsgardComposition.Item[i].WriteValue);
                i++;
            }

            var statpipeDelay = config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("Statpipe transport time"));
            var statpipeValues = DataAccess.GetValue(new List<string>(config.StatpipeComposition.GetTags()), DateTime.Now.AddHours(-statpipeDelay.Value));

            i = 0;
            foreach (var item in statpipeValues)
            {
                config.StatpipeComposition.Item[i].WriteValue = item;
                logger.Debug(CultureInfo.InvariantCulture, "Statpipe delayed composition tag {0}, value {1}",
                    config.StatpipeComposition.GetTags()[i], config.StatpipeComposition.Item[i].WriteValue);
                i++;
            }
        }

        private void CalculateVolumeFlow()
        {
            config.AsgardMeasurements.Item.Find(x => x.Name.Contains("Åsgard volume flow")).Value = CalculateAsgardVolumeFlow();
            config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("Statpipe volume flow")).Value = CalculateStatpipeVolumeFlow();
        }

        private void CalculateDelays()
        {
            double asgardPipeVolume = 18085.0;
            double statpipePipeVolume = 7901.0;
            string asgardTag = config.AsgardMeasurements.Item.Find(x => x.Name.Contains("Åsgard volume flow")).Tag;
            string statpipeTag = config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("Statpipe volume flow")).Tag;

            TimeSpan asgardDelay = CalculateDelay(asgardTag, asgardPipeVolume);
            logger.Debug(CultureInfo.InvariantCulture, "Åsgard delay: {0} h", asgardDelay.TotalHours);

            TimeSpan statpipeDelay = CalculateDelay(statpipeTag, statpipePipeVolume);
            logger.Debug(CultureInfo.InvariantCulture, "Statpipe delay: {0} h", statpipeDelay.TotalHours);

            config.AsgardMeasurements.Item.Find(x => x.Name.Contains("Åsgard transport time")).Value = asgardDelay.TotalHours;
            config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("Statpipe transport time")).Value = statpipeDelay.TotalHours;
        }

        private static TimeSpan CalculateDelay(string Tag, double PipeVolume)
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
                if (!item.Output)
                {
                    nodes.Add(item.Tag); types.Add(typeof(object));
                }
            }

            foreach (var item in config.AsgardMeasurements.Item)
            {
                if (!item.Output)
                {
                    nodes.Add(item.Tag); types.Add(typeof(object));
                }
            }

            foreach (var item in config.StatpipeMeasurements.Item)
            {
                if (!item.Output)
                {
                    nodes.Add(item.Tag); types.Add(typeof(object));
                }
            }

            foreach (var item in config.Validation.Item)
            {
                if (!item.Output)
                {
                    nodes.Add(item.Tag); types.Add(typeof(object));
                }
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
                Status = 1.0;
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
                if (!meas.Output)
                {
                    meas.Value = Convert.ToDouble(result[it++], CultureInfo.InvariantCulture);
                    meas.TimeStamp = DateTime.Now;
                    logger.Debug(CultureInfo.InvariantCulture,
                        "Measurement: \"{0}\" Value: {1} TimeStamp: {2} Tag: \"{3}\"",
                        meas.Name, meas.Value, meas.TimeStamp.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), meas.Tag);
                }
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

            foreach (var meas in config.Validation.Item)
            {
                if (!meas.Output)
                {
                    meas.Value = Convert.ToDouble(result[it++], CultureInfo.InvariantCulture);
                    meas.TimeStamp = DateTime.Now;
                    logger.Debug(CultureInfo.InvariantCulture,
                        "Measurement: \"{0}\" Value: {1} TimeStamp: {2} Tag: \"{3}\"",
                        meas.Name, meas.Value, meas.TimeStamp.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture), meas.Tag);
                }
            }
        }

        private void StoreHistoryMeasurements()
        {
            try
            {
                int storedRows = DataAccess.StoreValue(config.HistoryMeasurements.Item);
                logger.Debug(CultureInfo.InvariantCulture,
                    "Wrote {0} values to History database.", storedRows);
                int clearedRows = DataAccess.ClearHistory(DateTime.Now.AddHours(-config.HistoryMeasurements.MaxAge));
                if (clearedRows > 0)
                {
                    logger.Debug(CultureInfo.InvariantCulture,
                        "Cleared {0} values that were older than {1} hours from History database.", clearedRows, config.HistoryMeasurements.MaxAge);
                }
            }
            catch (Exception e)
            {
                Status = 1.0;
                logger.Error(e, "Error writing to History database");
            }
        }

        private double CalculateAsgardVolumeFlow()
        {
            double massFlowBeforeXover = config.AsgardMeasurements.Item.Find(x => x.Name.Contains("Åsgard mass flow før x-over")).Value;
            double densityBeforeXover = config.AsgardMeasurements.Item.Find(x => x.Name.Contains("Åsgard density før x-over")).Value;
            double massFlowXover = config.AsgardMeasurements.Item.Find(x => x.Name.Contains("Åsgard mass flow x-over")).Value;
            double densityKarsto = config.AsgardMeasurements.Item.Find(x => x.Name.Contains("Åsgard density Kårstø")).Value;
            double densityKalsto = config.AsgardMeasurements.Item.Find(x => x.Name.Contains("Åsgard density Kalstø")).Value;

            // Convert from mass flow to diff pressure [mbar]
            double dp = Math.Pow(massFlowBeforeXover * Math.Sqrt(350.0) / 3105.0, 2.0);
            logger.Debug(CultureInfo.InvariantCulture, "Åsgard diff pressure before x-over {0} mbar", dp);

            // Calculate mass flow before x-over [t/h]
            massFlowBeforeXover = 1312.0 * Math.Sqrt(densityBeforeXover * dp * 100.0) / 1000.0;
            logger.Debug(CultureInfo.InvariantCulture, "Åsgard corrected mass flow before x-over {0} t/h", massFlowBeforeXover);
            // Store mass flow in history database
            config.HistoryMeasurements.Item.Find(x => x.Name.Contains("Åsgard corrected mass flow før x-over")).Value = massFlowBeforeXover;
            config.HistoryMeasurements.Item.Find(x => x.Name.Contains("Åsgard corrected mass flow før x-over")).TimeStamp = DateTime.Now;

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
            double massFlowXover = config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("Statpipe mass flow x-over")).Value;
            double massFlowXoverSTP = config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("Statpipe mass flow x-over STP")).Value;
            double densityKarsto = config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("Statpipe density Kårstø")).Value;
            double densityKalsto = config.StatpipeMeasurements.Item.Find(x => x.Name.Contains("Statpipe density Kalstø")).Value;

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

            foreach (var item in config.AsgardComposition.Item)
            {
                if (!string.IsNullOrEmpty(item.WriteTag) && !string.IsNullOrEmpty(item.Type))
                {
                    wvc.Add(new WriteValue()
                    {
                        NodeId = item.WriteTag,
                        AttributeId = Attributes.Value,
                        Value = new DataValue { Value = item.WriteValue }
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

            foreach (var item in config.StatpipeComposition.Item)
            {
                if (!string.IsNullOrEmpty(item.WriteTag) && !string.IsNullOrEmpty(item.Type))
                {
                    wvc.Add(new WriteValue()
                    {
                        NodeId = item.WriteTag,
                        AttributeId = Attributes.Value,
                        Value = new DataValue { Value = item.WriteValue }
                    });
                }
            }

            foreach (var item in config.T100Composition.Item)
            {
                if (!string.IsNullOrEmpty(item.WriteTag) && !string.IsNullOrEmpty(item.Type))
                {
                    wvc.Add(new WriteValue()
                    {
                        NodeId = item.WriteTag,
                        AttributeId = Attributes.Value,
                        Value = new DataValue { Value = item.WriteValue * 10_000.0 }
                    });
                }
            }

            foreach (var item in config.T400Composition.Item)
            {
                if (!string.IsNullOrEmpty(item.WriteTag) && !string.IsNullOrEmpty(item.Type))
                {
                    wvc.Add(new WriteValue()
                    {
                        NodeId = item.WriteTag,
                        AttributeId = Attributes.Value,
                        Value = new DataValue { Value = item.WriteValue * 10_000.0 }
                    });
                }
            }

            var status = config.Validation.Item.Find(x => x.Name.Contains("PhaseOpt status"));
            if (!string.IsNullOrEmpty(status.Tag) && !string.IsNullOrEmpty(status.Type))
            {
                wvc.Add(new WriteValue()
                {
                    NodeId = status.Tag,
                    AttributeId = Attributes.Value,
                    Value = new DataValue { Value = status.GetTypedValue() }
                });
            }

            wvc.Add(new WriteValue()
            {
                NodeId = config.WatchDog.Tag,
                AttributeId = Attributes.Value,
                Value = new DataValue { Value = config.WatchDog.GetTypedValue() }
            });

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
                Status = 1.0;
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
