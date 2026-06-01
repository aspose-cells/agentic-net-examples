using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Input workbook path
            string inputPath = "input.xlsx";

            // Create a sample workbook if the input file does not exist
            if (!File.Exists(inputPath))
            {
                var wb = new Workbook();
                var ws = wb.Worksheets[0];
                ws.Cells["A1"].PutValue(1);
                ws.Cells["A2"].PutValue(2);
                ws.Cells["A3"].PutValue(3);
                ws.Cells["B1"].PutValue(2);
                ws.Cells["B2"].PutValue(4);
                ws.Cells["B3"].PutValue(6);
                wb.Save(inputPath);
            }

            // Load the workbook (ensure the file exists before loading)
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"Input file not found: {inputPath}");

            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure there is at least one chart; create a simple line chart if none exist
            Chart chart;
            if (sheet.Charts.Count == 0)
            {
                int chartIdx = sheet.Charts.Add(ChartType.Line, 5, 0, 15, 5);
                chart = sheet.Charts[chartIdx];
                chart.NSeries.Add("A1:B3", true);
            }
            else
            {
                chart = sheet.Charts[0];
            }

            // Add a linear trendline to the first series
            int trendlineIdx = chart.NSeries[0].TrendLines.Add(TrendlineType.Linear);
            Trendline trendline = chart.NSeries[0].TrendLines[trendlineIdx];
            trendline.DisplayEquation = true; // Show equation on the chart

            // Save the workbook to a temporary file (required for some operations)
            string tempPath = "temp.xlsx";
            workbook.Save(tempPath);

            // Retrieve the trendline equation.
            // Note: Aspose.Cells may not expose the equation directly via a property in all versions.
            // If unavailable, fall back to a placeholder message.
            string equation = "Equation not available";
            try
            {
                // Attempt to use the TrendlineEquation property if it exists (available in newer versions)
                var prop = typeof(Trendline).GetProperty("TrendlineEquation");
                if (prop != null)
                {
                    var value = prop.GetValue(trendline) as string;
                    if (!string.IsNullOrEmpty(value))
                        equation = value;
                }
            }
            catch
            {
                // Ignore any reflection errors and keep the placeholder.
            }

            // Output the extracted equation to the console
            Console.WriteLine("Trendline equation: " + equation);
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error: " + ex.Message);
        }
    }
}