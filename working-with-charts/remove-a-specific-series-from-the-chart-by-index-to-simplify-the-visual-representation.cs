// Title: How to remove a specific data series from an Excel chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that loads an .xlsx workbook, verifies the first worksheet contains a chart, and removes the series at index 1 from that chart. | Generate a robust C# example that checks for the input file, ensures the chart exists, uses NSeries.RemoveAt to delete a data series by zero‑based index, and saves the updated workbook. | Create a C# snippet that handles FileNotFound and out‑of‑range series index exceptions while programmatically dropping a specified series from an Excel chart with Aspose.Cells.
// Common Searches: aspnet remove chart series by index Aspose.Cells example | C# Aspose.Cells delete second series from chart in existing workbook | how to programmatically drop a data series from an Excel chart using Aspose.Cells | remove unwanted series from Excel chart with Aspose.Cells .NET
// Tags: Aspose.Cells remove chart series C# | NSeries.RemoveAt Excel chart | delete data series from Excel chart .NET | chart series index removal Aspose.Cells | handle missing chart Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Loads an existing .xlsx file, checks for a chart in the first worksheet, removes the data series at the specified zero‑based index using NSeries.RemoveAt, and saves the workbook with the updated chart.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
                throw new FileNotFoundException($"The input file '{inputPath}' was not found.");

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Verify that the worksheet contains at least one chart
            if (sheet.Charts.Count == 0)
                throw new InvalidOperationException("No charts found in the first worksheet.");

            // Access the target chart (here we use the first chart in the sheet)
            int chartIndex = 0;
            Chart chart = sheet.Charts[chartIndex];

            // Index of the series to remove (0‑based). Change as required.
            int seriesIndexToRemove = 1; // example: remove the second series

            // Verify the index is valid before removal
            if (seriesIndexToRemove >= 0 && seriesIndexToRemove < chart.NSeries.Count)
            {
                // Remove the specified series from the chart
                chart.NSeries.RemoveAt(seriesIndexToRemove);
            }
            else
            {
                Console.WriteLine($"Series index {seriesIndexToRemove} is out of range. No series removed.");
            }

            // Save the workbook with the updated chart
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
