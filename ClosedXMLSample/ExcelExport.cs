using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Text;

namespace ClosedXMLSample
{
    public class ExcelExport
    {
        public const string TempA = "A";
        public const string TempB = "B";
        public const string TempC = "C";
        public const string Humidity = "D";

        public static void Export(List<SensorData> sensorDatas)
        {
            int count = 2;
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Sample Sheet");
                worksheet.Cell($"{TempA}1").Value = "온도A";
                worksheet.Cell($"{TempB}1").Value = "온도B";
                worksheet.Cell($"{TempC}1").Value = "온도C";
                worksheet.Cell($"{Humidity}1").Value = "습도";
                foreach (var item in sensorDatas)
                {
                    worksheet.Cell($"{TempA}{count}").Value = item.TempA.ToString();
                    worksheet.Cell($"{TempB}{count}").Value = item.TempB.ToString();
                    worksheet.Cell($"{TempC}{count}").Value = item.TempC.ToString();
                    worksheet.Cell($"{Humidity}{count}").Value = item.Humidity.ToString();
                    count++;
                }
                workbook.SaveAs("HelloWorld.xlsx");
            }

        }
    }
}
