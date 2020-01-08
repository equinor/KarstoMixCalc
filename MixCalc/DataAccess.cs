using Dapper;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SQLite;
using System.Globalization;
using System.Linq;

namespace MixCalc
{
    public class DataAccess
    {
        public static List<double> GetValue(List<string> Tag, DateTime TimeStamp, int Threshold = 360)
        {
            List<double> resultValues = new List<double>();
            string tagList = "(FALSE";
            foreach (var item in Tag)
            {
                tagList += "\n OR Tag = '" + item + "'";
            }
            tagList += ")";

            // gets all of the values that are closer than Threshold seconds to TimeStamp
            string query = $@"SELECT
    Tag, TimeStamp, Value
FROM
    History
WHERE
    {tagList}
    AND abs(CAST(strftime('%s', '{TimeStamp.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)}') AS NUMERIC) - CAST(strftime('%s', TimeStamp) AS NUMERIC)) < {Threshold}";

            using (IDbConnection con = new SQLiteConnection(LoadConnectionString()))
            {
                var values = con.Query<PhaseOptDcs.TimeStampedMeasurement>(query, new DynamicParameters());

                foreach (var tagItem in Tag)
                {
                    resultValues.Add(Interpolate(values.Where(v => v.Tag == tagItem).ToList(), TimeStamp));
                }
            }

            return resultValues;
        }

        public static int StoreValue(List<PhaseOptDcs.TimeStampedMeasurement> Value)
        {
            string query = "INSERT INTO History (Tag, Value, TimeStamp) VALUES (@Tag, @Value, @TimeStamp);";

            using (IDbConnection con = new SQLiteConnection(LoadConnectionString()))
            {
                con.Open();

                var affectedRows = con.Execute(query, Value);

                return affectedRows;
            }
        }

        public static int ClearHistory(System.DateTime TimeStamp)
        {
            string query = $@"DELETE
FROM
    History
WHERE
    CAST(strftime('%s', TimeStamp) AS NUMERIC) < CAST(strftime('%s', '{TimeStamp.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture)}') AS NUMERIC)";

            using (IDbConnection con = new SQLiteConnection(LoadConnectionString()))
            {
                con.Open();

                var affectedRows = con.Execute(query);

                return affectedRows;
            }

        }

        private static string LoadConnectionString(string id = "Default")
        {
            return ConfigurationManager.ConnectionStrings[id].ConnectionString;
        }

        private static double Interpolate(List<PhaseOptDcs.TimeStampedMeasurement> Values, DateTime TimeStamp)
        {
            double resultValue;
            if (Values.ToList().Count == 0)
            {
                resultValue = double.NaN;
            }
            else if (Values.ToList().Count == 1)
            {
                resultValue = Values.ToList()[0].Value;
            }
            else
            {

                PhaseOptDcs.TimeStampedMeasurement m0 = new PhaseOptDcs.TimeStampedMeasurement();
                PhaseOptDcs.TimeStampedMeasurement m1 = new PhaseOptDcs.TimeStampedMeasurement();
                double nearestBefore = Double.MaxValue;
                double nearestAfter = Double.MaxValue;
                foreach (var v in Values)
                {
                    if (v.TimeStamp < TimeStamp && (TimeStamp - v.TimeStamp).TotalSeconds < nearestBefore)
                    {
                        nearestBefore = (TimeStamp - v.TimeStamp).TotalSeconds;
                        m0 = v;
                    }
                    else if (v.TimeStamp > TimeStamp && (TimeStamp - v.TimeStamp).TotalSeconds < nearestAfter)
                    {
                        nearestAfter = (TimeStamp - v.TimeStamp).TotalSeconds;
                        m1 = v;
                    }
                }

                resultValue = m0.Value + (TimeStamp - m0.TimeStamp).TotalSeconds * (m1.Value - m0.Value) / ((m1.TimeStamp - m0.TimeStamp).TotalSeconds);
            }

            return resultValue;
        }
    }
}
