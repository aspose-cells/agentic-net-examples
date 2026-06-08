using System;
using Aspose.Cells;
using Aspose.Cells.Charts;
using Aspose.Cells.Utility;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsChartJsonExport
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate sample data (categories in column A, values in column B)
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["B1"].PutValue("Value");
                sheet.Cells["A2"].PutValue("A");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["A3"].PutValue("B");
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["A4"].PutValue("C");
                sheet.Cells["B4"].PutValue(30);

                // Add a column chart and bind the data series
                int chartIndex = sheet.Charts.Add(ChartType.Column, 5, 0, 15, 5);
                Chart chart = sheet.Charts[chartIndex];
                chart.NSeries.Add("=Sheet1!$B$2:$B$4", true);               // Y values
                chart.NSeries.CategoryData = "=Sheet1!$A$2:$A$4";          // X categories

                // Retrieve the first series (index 0)
                Series series = chart.NSeries[0];

                // The Values property contains the range address (e.g., "=Sheet1!$B$2:$B$4")
                string valuesAddress = series.Values;

                // Remove the leading '=' to obtain a pure range address
                if (valuesAddress.StartsWith("="))
                    valuesAddress = valuesAddress.Substring(1);

                // Create a Range object for the series values
                AsposeRange seriesRange = sheet.Cells.CreateRange(valuesAddress);

                // Configure JSON export options (no header row, export as native types)
                JsonSaveOptions jsonOptions = new JsonSaveOptions
                {
                    HasHeaderRow = false,
                    ExportAsString = false,
                    ExportEmptyCells = false,
                    ToExcelStruct = false
                };

                // Export the range to a JSON string
                string json = JsonUtility.ExportRangeToJson(seriesRange, jsonOptions);

                // Output the resulting JSON
                Console.WriteLine("Chart Series JSON:");
                Console.WriteLine(json);
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}