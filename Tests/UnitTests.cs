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
            System.DateTime ts = new System.DateTime(2019, 11, 25, 18, 35, 18);
            List<string> tagList = new List<string>
            {
                "31AI0161A_05",
                "31AI0161A_06",
                "31AI0161A_07",
                "31AI0161A_08",
                "31AI0161A_09"
            };
            var results = MixCalc.DataAccess.GetValue(tagList, ts);

            Assert.AreEqual(0.45604419999999996, results[0], 1.0e-6);
            Assert.AreEqual(6.5209647511312214, results[1], 1.0e-6);
            Assert.AreEqual(double.NaN, results[2]);
            Assert.AreEqual(4.609825, results[3], 1.0e-6);
            Assert.AreEqual(6.99234, results[4], 1.0e-6);
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
