// Title: Create an Excel column chart template with right‑aligned legend and value data labels using Aspose.Cells for .NET
// AI Prompts: Generate C# code with Aspose.Cells that adds a column chart, positions the legend on the right, and enables value data labels for the first series. | Adjust the chart template to move the legend to the top and configure data labels to display both values and percentages for each series.
// Common Searches: Aspose.Cells C# set chart legend position to right | Enable value data labels for a column chart series using Aspose.Cells | Create reusable Excel chart template with predefined legend and data label settings in .NET | How to change legend placement and data label format in Aspose.Cells charts | Aspose.Cells column chart example with sample data and customized legend
// Tags: Aspose.Cells chart legend right alignment | Aspose.Cells series data label values | Aspose.Cells column chart template | Aspose.Cells reusable chart formatting .NET | Aspose.Cells chart with sample data

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example creates a new workbook, fills it with category and value data, adds a column chart, sets the legend to the right side, enables value data labels for the first series, and saves the file as 'WorkbookWithChart.xlsx'.
class Program
{
    static void Main()
    {
        try
        {
            // ---------- Create a workbook and add sample data ----------
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);

            // ---------- Add a chart ----------
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Set the data range for the chart
            chart.NSeries.Add("B2:B4", true);
            chart.NSeries.CategoryData = "A2:A4";

            // ---------- Predefined legend position ----------
            chart.Legend.Position = LegendPositionType.Right; // Options: Top, Bottom, Left, Right, etc.

            // ---------- Data label settings ----------
            // Enable data labels for the first series (show values)
            chart.NSeries[0].DataLabels.ShowValue = true;
            // Optional: set position of data labels if needed
            // chart.NSeries[0].DataLabels.Position = DataLabelPositionType.Center;

            // ---------- Save the workbook ----------
            string outputPath = "WorkbookWithChart.xlsx";

            // Ensure the directory exists before saving
            string outputDir = Path.GetDirectoryName(outputPath);
            if (string.IsNullOrEmpty(outputDir))
            {
                outputDir = Directory.GetCurrentDirectory();
            }
            if (!Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{Path.GetFullPath(outputPath)}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
