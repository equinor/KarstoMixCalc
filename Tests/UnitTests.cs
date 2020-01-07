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
    }
}
