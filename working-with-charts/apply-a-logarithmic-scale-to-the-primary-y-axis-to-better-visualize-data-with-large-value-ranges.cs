// Title: How to set a logarithmic scale on the primary Y‑axis of an Excel chart using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that opens an existing .xlsx workbook, locates the first chart, and sets its primary Y‑axis to a logarithmic scale with base 10 using Aspose.Cells. | Show a snippet that verifies a worksheet contains charts before applying IsLogarithmic = true to the chart's value axis. | Demonstrate how to save the workbook after changing the chart axis to logarithmic scaling in a .NET console application.
// Common Searches: Aspose.Cells C# set chart primary Y axis to logarithmic scale | Enable log base 10 on Excel chart axis using Aspose.Cells .NET | Change value axis to log scale in existing workbook with Aspose.Cells | C# example for applying logarithmic scaling to chart axis in Excel file
// Tags: Aspose.Cells chart logarithmic Y axis | set primary value axis log scale .NET | logarithmic axis base 10 Aspose.Cells | modify Excel chart axis C# Aspose | apply log scaling to chart using Aspose.Cells

using Aspose.Cells;
using Aspose.Cells.Charts;
using System;
using System.IO;

// Loads an existing Excel workbook, accesses the first chart on the first worksheet, enables logarithmic scaling on its primary Y‑axis (value axis) with a base of 10, and saves the modified workbook to a new file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook that contains the chart
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet worksheet = workbook.Worksheets[0];

            // Ensure the worksheet contains at least one chart
            if (worksheet.Charts.Count == 0)
            {
                Console.WriteLine("No charts found in the worksheet.");
                return;
            }

            // Get the first chart on the worksheet (adjust index if needed)
            Chart chart = worksheet.Charts[0];

            // Access the primary Y axis (value axis) of the chart
            Axis primaryYAxis = chart.ValueAxis;

            // Enable logarithmic scaling for the Y axis
            primaryYAxis.IsLogarithmic = true;

            // Optionally set the logarithmic base (default is 10)
            primaryYAxis.LogBase = 10;

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
