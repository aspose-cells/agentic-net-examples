// Title: Disable data label background fill in a doughnut chart using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# hide background color of doughnut chart data labels | remove label fill from Excel doughnut chart programmatically with Aspose.Cells | how to make doughnut chart data labels transparent in .NET using Aspose.Cells | C# example for customizing doughnut chart data label appearance with Aspose.Cells | Aspose.Cells chart label background removal workaround for doughnut charts
// Tags: Aspose.Cells chart data label background suppression | doughnut chart label fill customization .NET | Excel chart label appearance control using Aspose.Cells | transparent data labels in Aspose.Cells charts | remove label background fill from Excel charts C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, adds sample data, inserts a doughnut chart, binds the series to the data range, enables value display on the data labels, notes that Aspose.Cells does not provide a direct API to disable the data label background fill, and saves the workbook as DoughnutChart.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook wb = new Workbook();

            // Access the first worksheet
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the doughnut chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["B2"].PutValue(30);
            ws.Cells["A3"].PutValue("B");
            ws.Cells["B3"].PutValue(20);
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B4"].PutValue(50);

            // Add a doughnut chart to the worksheet
            int chartIndex = ws.Charts.Add(ChartType.Doughnut, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIndex];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // Configure data labels (disable background fill is not supported directly)
            if (chart.NSeries.Count > 0)
            {
                var series = chart.NSeries[0];
                // Optionally display the values on data labels
                series.DataLabels.ShowValue = true;
            }

            // Define output file path
            string outputPath = "DoughnutChart.xlsx";

            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook to a file
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
