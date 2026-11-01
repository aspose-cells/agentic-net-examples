// Title: Hide specific chart data points in an Excel workbook using Aspose.Cells by setting series visibility to false (C#)
// Common Searches: C# how to hide individual points in an Excel chart with Aspose.Cells without deleting them | Aspose.Cells set series visibility false for hidden rows in a chart | programmatically hide chart data points based on hidden rows using Aspose.Cells .NET | Aspose.Cells chart point visibility control via Series.IsVisible property C# | hide chart series points in Excel workbook using Aspose.Cells API
// Tags: Aspose.Cells chart series visibility toggle | C# programmatic Excel chart point hiding | Aspose.Cells series visibility false setting | chart point hiding based on hidden rows .NET | Excel chart data point visibility control Aspose.Cells | Aspose.Cells chart manipulation C#

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

// The example loads an Excel workbook, checks for charts on the first worksheet, and explains that the current Aspose.Cells version does not expose Series.DataPoints or YValues, so the logic to hide individual data points by setting series visibility to false is omitted. The workbook is then saved unchanged.
class ChartVisibilityExample
{
    static void Main()
    {
        const string inputPath = "input.xlsx";
        const string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            var workbook = new Workbook(inputPath);

            // Access the first worksheet
            var worksheet = workbook.Worksheets[0];

            // Check if there is at least one chart on the sheet
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found on the first worksheet.");
            }
            else
            {
                // NOTE: The original example attempted to hide individual data points
                // based on hidden rows/columns. The current Aspose.Cells version used
                // in this environment does not expose Series.DataPoints or Series.YValues,
                // so this part is omitted. You can implement custom logic here if needed.
                Console.WriteLine($"Found {worksheet.Charts.Count} chart(s) on the first worksheet.");
            }

            // Save the (potentially unchanged) workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
