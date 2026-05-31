using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;
using System.Drawing;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Add sample data for two series
                sheet.Cells["A1"].PutValue("Category");
                sheet.Cells["A2"].PutValue("Jan");
                sheet.Cells["A3"].PutValue("Feb");
                sheet.Cells["A4"].PutValue("Mar");

                sheet.Cells["B1"].PutValue("Series1");
                sheet.Cells["B2"].PutValue(10);
                sheet.Cells["B3"].PutValue(20);
                sheet.Cells["B4"].PutValue(30);

                sheet.Cells["C1"].PutValue("Series2");
                sheet.Cells["C2"].PutValue(15);
                sheet.Cells["C3"].PutValue(25);
                sheet.Cells["C4"].PutValue(35);

                // Add a line chart with data markers
                int chartIdx = sheet.Charts.Add(ChartType.LineWithDataMarkers, 5, 0, 20, 10);
                Chart chart = sheet.Charts[chartIdx];

                // Set the data range for the two series
                chart.NSeries.Add("B2:B4", true); // Series1
                chart.NSeries.Add("C2:C4", true); // Series2
                chart.NSeries.CategoryData = "A2:A4";

                // Hide markers for the first series
                chart.NSeries[0].Marker.MarkerStyle = ChartMarkerType.None;

                // Ensure markers are visible for the second series
                chart.NSeries[1].Marker.MarkerStyle = ChartMarkerType.Circle;
                chart.NSeries[1].Marker.MarkerSize = 8;
                chart.NSeries[1].Marker.ForegroundColor = Color.Blue;

                // Save the workbook
                string outputPath = "LineChartWithSelectiveMarkers.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}