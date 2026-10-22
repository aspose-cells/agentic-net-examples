// Title: Disable data label text wrapping in an Excel chart with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx workbook, finds the first chart's first series, sets its DataLabels.IsTextWrapped property to false, and saves the file. | Show a concise example that disables text wrapping for chart data labels using Aspose.Cells for .NET, including basic checks for file existence and chart presence.
// Common Searches: aspnet cells set chart data label no wrap c# | c# aspose.cells disable text wrapping for Excel chart labels | how to turn off data label wrapping in an Excel chart using Aspose.Cells | example of DataLabels.IsTextWrapped false in Aspose.Cells C#
// Tags: Aspose.Cells chart data label no wrap | DataLabels.IsTextWrapped false C# | Excel chart label wrapping Aspose.Cells | modify chart series data labels .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// Loads an existing Excel workbook, accesses the first worksheet and its first chart, disables text wrapping for the data labels of the first series via DataLabels.IsTextWrapped = false, and saves the modified workbook.
class Program
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        try
        {
            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                return;
            }

            // Load the existing workbook that contains a chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet has at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the first worksheet.");
                return;
            }

            // Access the first chart on the worksheet (adjust index if needed)
            Chart chart = worksheet.Charts[0];

            // Ensure the chart has at least one series to access data labels
            if (chart.NSeries.Count == 0)
            {
                Console.WriteLine("The chart does not contain any series.");
                return;
            }

            // Disable text wrapping for the first series' data labels
            chart.NSeries[0].DataLabels.IsTextWrapped = false;

            // Save the modified workbook
            workbook.Save(outputPath, SaveFormat.Xlsx);
            Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            // Catch any unexpected errors and display a friendly message
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
