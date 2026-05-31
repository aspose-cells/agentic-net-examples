using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class AddExponentialTrendline
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xls";
            string outputPath = "output.xls";

            // Load existing workbook or create a new one with sample data
            Workbook workbook;
            if (File.Exists(inputPath))
            {
                workbook = new Workbook(inputPath);
            }
            else
            {
                workbook = new Workbook();
                Worksheet ws = workbook.Worksheets[0];
                ws.Cells["A1"].PutValue(1);
                ws.Cells["A2"].PutValue(2);
                ws.Cells["A3"].PutValue(3);
                ws.Cells["A4"].PutValue(4);
                ws.Cells["B1"].PutValue(2);
                ws.Cells["B2"].PutValue(4);
                ws.Cells["B3"].PutValue(6);
                ws.Cells["B4"].PutValue(8);
            }

            // First worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Get existing chart or create a new line chart
            Chart chart;
            if (sheet.Charts.Count > 0)
            {
                chart = sheet.Charts[0];
            }
            else
            {
                int chartIndex = sheet.Charts.Add(ChartType.Line, 5, 0, 20, 8);
                chart = sheet.Charts[chartIndex];
                // Add a dummy series (A1:B4)
                chart.NSeries.Add("A1:B4", true);
            }

            // Add exponential trendline to the first series
            int trendlineIndex = chart.NSeries[0].TrendLines.Add(TrendlineType.Exponential);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIndex];
            trendline.DisplayEquation = true;      // Show equation on chart
            trendline.DisplayRSquared = false;     // Hide R² (optional)

            // Save the workbook
            workbook.Save(outputPath);

            // Equation retrieval via API is not supported; inform the user
            string equationText = "Equation not available via API.";
            Console.WriteLine("Exponential Trendline Equation: " + equationText);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}