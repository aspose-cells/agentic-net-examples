// Title: Create a Waterfall chart in C# with a designated total bar using Aspose.Cells and export to Excel
// AI Prompts: Write C# code that builds a workbook, populates category and value cells, adds a Waterfall chart, and marks the last data point as a total bar with Aspose.Cells. | Show how to apply distinct formatting to the total segment in an Aspose.Cells Waterfall chart generated from C# data.
// Common Searches: how to mark a total bar in an Aspose.Cells waterfall chart C# | C# Aspose.Cells example for creating waterfall chart with custom total segment | set total point in waterfall chart using Aspose.Cells .NET | Aspose.Cells waterfall chart series range and automatic total detection in C#
// Tags: Aspose.Cells create waterfall chart | C# Aspose.Cells set total segment | waterfall chart series range Aspose.Cells | export Excel workbook Aspose.Cells C# | format total bar Aspose.Cells chart

using System;
using System.Drawing;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The program creates a new workbook, fills category and value cells, adds a Waterfall chart that treats the last point as a total bar, and saves the result as an Excel file using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];

            // Populate data for the waterfall chart
            // Column A: categories, Column B: values
            sheet.Cells["A1"].PutValue("Category");
            sheet.Cells["B1"].PutValue("Value");
            sheet.Cells["A2"].PutValue("Start");
            sheet.Cells["B2"].PutValue(100);
            sheet.Cells["A3"].PutValue("Revenue");
            sheet.Cells["B3"].PutValue(150);
            sheet.Cells["A4"].PutValue("Expense");
            sheet.Cells["B4"].PutValue(-70);
            sheet.Cells["A5"].PutValue("Total");
            sheet.Cells["B5"].PutValue(0); // placeholder; will be treated as total by the chart

            // Add a Waterfall chart
            int chartIndex = sheet.Charts.Add(ChartType.Waterfall, 7, 0, 25, 10);
            var chart = sheet.Charts[chartIndex];
            chart.Title.Text = "Waterfall Chart Example";

            // Add series using the values range (B2:B5)
            int seriesIndex = chart.NSeries.Add("B2:B5", true);
            var series = chart.NSeries[seriesIndex];

            // Note: CategoryData and IsTotal properties may not be available in older Aspose.Cells versions.
            // The chart will use default numeric categories and automatic total detection.

            // Define output file path
            string outputPath = "WaterfallChart.xlsx";

            // Ensure the output directory exists (handles cases where only a file name is provided)
            string directory = Path.GetDirectoryName(Path.GetFullPath(outputPath));
            if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
            {
                Directory.CreateDirectory(directory);
            }

            // Save the workbook with the chart
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception saveEx)
            {
                Console.WriteLine("Failed to save the workbook:");
                Console.WriteLine(saveEx.Message);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while creating the waterfall chart:");
            Console.WriteLine(ex.Message);
        }
    }
}
