// Title: How to remove a data series from an Excel chart by index using Aspose.Cells for .NET (C#)
// Common Searches: c# aspnet remove chart series at specific index using Aspose.Cells | delete second series from Excel chart programmatically with Aspose.Cells | Aspose.Cells chart NSeries.RemoveAt example in C# | remove unwanted series from Excel chart workbook using Aspose.Cells .NET | how to hide a data series in Excel chart via Aspose.Cells code
// Tags: Aspose.Cells chart series removal | C# NSeries.RemoveAt method | Excel chart data series deletion Aspose.Cells | modify chart series Aspose.Cells .NET | remove chart series by index Excel C#

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Loads an existing workbook, verifies a chart exists on the first worksheet, and removes the data series at a given zero‑based index from the chart's NSeries collection using Aspose.Cells before saving the updated file.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index or name as needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet (adjust index or name as needed)
            Chart chart = sheet.Charts[0];

            // Index of the series to remove (0‑based). Example: remove the second series.
            int seriesIndex = 1;

            // Verify that the index is valid before attempting removal
            if (seriesIndex >= 0 && seriesIndex < chart.NSeries.Count)
            {
                chart.NSeries.RemoveAt(seriesIndex);
                Console.WriteLine($"Series at index {seriesIndex} removed.");
            }
            else
            {
                Console.WriteLine("Series index out of range.");
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
