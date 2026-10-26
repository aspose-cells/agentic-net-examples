// Title: Export an Excel worksheet chart to PNG using Aspose.Cells for .NET (C# example)
// Common Searches: Aspose.Cells C# export chart from workbook to PNG image | How to save the first chart in an Excel sheet as a PNG file with Aspose.Cells | C# code to convert Excel chart to PNG using Aspose.Cells Chart.ToImage | Export Excel chart to image folder using Aspose.Cells .NET library | Save chart as PNG in C# without opening Excel
// Tags: Aspose.Cells chart export to PNG | C# Aspose.Cells Chart.ToImage usage | export Excel chart as image .NET | save worksheet chart as PNG file | Aspose.Cells image conversion for charts

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// This C# sample loads an Excel workbook, locates the first chart on the first worksheet, and uses Aspose.Cells' Chart.ToImage method to export the chart as a PNG file named "ChartExport.png" into a specified output folder, handling missing files and exceptions.
class ExportChartAsPng
{
    static void Main()
    {
        // Define input workbook path and output folder
        string workbookPath = @"C:\Input\Sample.xlsx";
        string outputFolder = @"C:\Output";

        try
        {
            // Verify that the input workbook exists
            if (!File.Exists(workbookPath))
            {
                Console.WriteLine($"Input workbook not found: {workbookPath}");
                return;
            }

            // Ensure the output folder exists
            if (!Directory.Exists(outputFolder))
            {
                Directory.CreateDirectory(outputFolder);
            }

            // Load the workbook
            Workbook workbook = new Workbook(workbookPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Retrieve the first chart on the worksheet (adjust index if needed)
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            Chart chart = sheet.Charts[0];

            // Define the output PNG file path
            string pngPath = Path.Combine(outputFolder, "ChartExport.png");

            // Export the chart as a PNG image (default format is PNG)
            chart.ToImage(pngPath);

            Console.WriteLine($"Chart exported successfully to: {pngPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
