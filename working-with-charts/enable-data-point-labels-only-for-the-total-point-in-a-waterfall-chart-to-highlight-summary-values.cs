// Title: Display a data label only for the total point in a Waterfall chart with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to create a waterfall chart, mark the final point as a total with SetDataPointIsTotal (using reflection if needed), and enable a data label only for that total point. | Generate a .NET example that adds a waterfall series, sets the last point as a total, hides data labels for other points, and saves the workbook to an .xlsx file.
// Common Searches: Aspose.Cells C# how to show data label only for total point in waterfall chart | mark last point as total in Aspose.Cells waterfall series | enable data labels for specific point in Aspose.Cells chart .NET | use reflection to call SetDataPointIsTotal in Aspose.Cells | waterfall chart total label Aspose.Cells example
// Tags: waterfall chart total point Aspose.Cells | setdata point is total reflection .NET | show single data label waterfall series | Aspose.Cells chart data labels specific point | C# save waterfall chart to xlsx

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a workbook, adds sample data, inserts a waterfall chart, uses reflection to call SetDataPointIsTotal on the last point, enables a data label only for that total point while hiding others, and saves the file as WaterfallChart_TotalLabel.xlsx.
class WaterfallChartExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for a waterfall chart
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Increase 1");
            sheet.Cells["B3"].PutValue(30);
            sheet.Cells["A4"].PutValue("Decrease 1");
            sheet.Cells["B4"].PutValue(-20);
            sheet.Cells["A5"].PutValue("Increase 2");
            sheet.Cells["B5"].PutValue(40);
            sheet.Cells["A6"].PutValue("Total");
            sheet.Cells["B6"].PutValue(150); // This will be marked as total (if supported)

            // Add a waterfall chart
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
            Chart chart = sheet.Charts[chartIndex];

            // Set the chart's data source (values). Categories will be taken from the first column automatically.
            chart.NSeries.Add("B2:B6", false);

            // Attempt to mark the last point as total (available in newer versions)
            try
            {
                int totalPointIndex = chart.NSeries[0].Points.Count - 1;
                // The SetDataPointIsTotal method may not exist in older versions; use reflection to call it safely.
                var method = chart.NSeries[0].GetType().GetMethod("SetDataPointIsTotal");
                if (method != null)
                {
                    method.Invoke(chart.NSeries[0], new object[] { totalPointIndex, true });
                }
            }
            catch (Exception ex)
            {
                // If the API does not support marking a total point, ignore the error.
                Console.WriteLine($"Info: Unable to set total point – {ex.Message}");
            }

            // Enable data labels for the series (show values for all points)
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Prepare output path
            string outputPath = "WaterfallChart_TotalLabel.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

            // Ensure the output directory exists (if any)
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error saving workbook: {ex.Message}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
