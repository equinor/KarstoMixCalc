using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace Tests
{
    [TestClass]
    public class UnitTests
    {
        [TestMethod]
        public void DataAccess_GetValue()
        {
            List<MixCalc.TimeStampedMeasurement> value = new List<MixCalc.TimeStampedMeasurement>();
            value.Add(new MixCalc.TimeStampedMeasurement()
            {
                Tag = "Test-Tag",
                Value = System.Math.PI,
                TimeStamp = new System.DateTime(2020, 01, 04, 13, 28, 52)
            });

            value.Add(new MixCalc.TimeStampedMeasurement()
            {
                Tag = "Test-Tag",
                Value = System.Math.E,
                TimeStamp = new System.DateTime(2020, 01, 04, 13, 36, 17)
            });

            try
            {
                MixCalc.DataAccess.StoreValue(value);
            }
            catch
            {
            }

            System.DateTime ts = new System.DateTime(2020, 01, 04, 13, 34, 13);
            List<string> tagList = new List<string>
            {
                "Test-Tag"
            };

            var results = MixCalc.DataAccess.GetValue(tagList, ts);

            MixCalc.DataAccess.ClearHistory(System.DateTime.Now);

            Assert.AreEqual(2.836_238_103_326, results[0], 1.0e-6);
        }

        [TestMethod]
        public void DataAccess_StoreValue()
        {
            List<MixCalc.TimeStampedMeasurement> value = new List<MixCalc.TimeStampedMeasurement>();

            int rows = 5;
            for (int i = 0; i < rows; i++)
            {
                value.Add(new MixCalc.TimeStampedMeasurement()
                {
                    Tag = $@"TestTag{i}",
                    Value = 7357.749,
                    TimeStamp = System.DateTime.Now
                });
            }

            int affectedRows = MixCalc.DataAccess.StoreValue(value);
            Assert.AreEqual(rows, affectedRows);

            var result = MixCalc.DataAccess.GetValue(new List<string> { value[0].Tag }, System.DateTime.Now);
            Assert.AreEqual(value[0].Value, result[0], 1.0e-5);
        }

        [TestMethod]
        public void DataAccess_ClearHistory()
        {
            List<MixCalc.TimeStampedMeasurement> value = new List<MixCalc.TimeStampedMeasurement>();

            int rows = 5;
            for (int i = 0; i < rows; i++)
            {
                value.Add(new MixCalc.TimeStampedMeasurement()
                {
                    Tag = $@"TestTag{i}",
                    Value = 7357.749,
                    TimeStamp = new System.DateTime(2019, 09, 10, 12, 00, 00)
                });
            }

            int storedRows = MixCalc.DataAccess.StoreValue(value);
            Assert.AreEqual(rows, storedRows);

            int clearedRows = MixCalc.DataAccess.ClearHistory(new System.DateTime(2019, 09, 11, 12, 00, 00));
            Assert.AreEqual(storedRows, clearedRows);
        }

        [TestMethod]
        public void DataAccess_GetValue_Exception()
        {
            System.DateTime ts = new System.DateTime(2019, 11, 25, 18, 35, 18);
            List<string> tagList = null;
            System.Action getValue = () => MixCalc.DataAccess.GetValue(tagList, ts);
            Assert.ThrowsException<System.Exception>(getValue);
        }
    }
}
