// Title: Enable data labels that display point values for the first series of a column chart using Aspose.Cells for .NET
// AI Prompts: Write C# code that creates a new workbook, adds sample data, inserts a column chart, and activates ShowValue on the first series to display each point's value. | Modify an existing Aspose.Cells chart in C# so that the first NSeries shows its values as data labels without altering other series. | Generate an Excel file with a column chart where the chart's first series automatically shows data labels containing the underlying cell values, using the Aspose.Cells API.
// Common Searches: Aspose.Cells C# enable data labels for first series of column chart | how to turn on ShowValue for a chart series using Aspose.Cells .NET | display point values as data labels in Excel column chart with Aspose.Cells | C# code example for adding data labels to the first series of an Aspose chart
// Tags: Aspose.Cells chart series ShowValue | C# column chart data labels Aspose | Enable data labels Excel chart Aspose.Cells | Aspose.Cells NSeries data label configuration | Excel workbook column chart Aspose.NET

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// The example creates a workbook, fills cells A1:B4 with sample data, adds a column chart, defines the first series range (B2:B4), enables data labels to show point values for that series, and saves the file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];

            // Populate sample data for the chart
            ws.Cells["A1"].PutValue("Category");
            ws.Cells["B1"].PutValue("Value");
            ws.Cells["A2"].PutValue("A");
            ws.Cells["A3"].PutValue("B");
            ws.Cells["A4"].PutValue("C");
            ws.Cells["B2"].PutValue(10);
            ws.Cells["B3"].PutValue(20);
            ws.Cells["B4"].PutValue(30);

            // Add a column chart to the worksheet
            int chartIdx = ws.Charts.Add(ChartType.Column, 5, 0, 20, 10);
            Chart chart = ws.Charts[chartIdx];

            // Set the data range for the series (categories will be taken from column A automatically)
            chart.NSeries.Add("B2:B4", true);

            // Enable data labels to display point values
            chart.NSeries[0].DataLabels.ShowValue = true;

            // Determine output path and ensure its directory exists
            string outputPath = "output.xlsx";
            string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

            // Save the workbook
            wb.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
