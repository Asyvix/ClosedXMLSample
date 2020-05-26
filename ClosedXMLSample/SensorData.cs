using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXMLSample
{
    public class SensorData
    {
        public int TempA { get; set; }
        public int TempB { get; set; }
        public int TempC { get; set; }
        public int Humidity { get; set; }

        public readonly static Random rnd = new Random();
        public static SensorData CreateSensorData()
        {
            SensorData sd = new SensorData();
            sd.TempA = rnd.Next(1, 500);
            sd.TempB = rnd.Next(1, 500);
            sd.TempC = rnd.Next(1, 500);
            sd.Humidity = rnd.Next(1, 500);
            return sd;
        }
    }
}
