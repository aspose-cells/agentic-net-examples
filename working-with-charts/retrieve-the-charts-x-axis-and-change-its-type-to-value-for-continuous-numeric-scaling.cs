// Title: Change an Excel chart's X‑axis to a value axis by converting it to a scatter chart with Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells change chart X axis to numeric value axis | How to convert an existing Excel chart to scatter type using Aspose.Cells .NET | Aspose.Cells set chart X axis scaling to continuous numbers in C# | Retrieve and modify chart axis type in a workbook with Aspose.Cells | Programmatically change Excel chart axis type to value axis using Aspose.Cells
// Tags: Aspose.Cells scatter chart conversion C# | Excel chart numeric X axis Aspose.Cells | change chart type to scatter Aspose.Cells | adjust chart axis scaling .NET | continuous numeric axis Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, accesses its first worksheet and chart, changes the chart type to Scatter (which converts the X‑axis to a value axis for continuous numeric scaling), and saves the modified workbook to a new file.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input workbook exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Error: The file \"{inputPath}\" was not found.");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Retrieve the first chart on the worksheet (adjust index if needed)
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                Chart chart = worksheet.Charts[0];

                // Change the chart type to a scatter chart so the X axis becomes a value axis
                chart.Type = ChartType.Scatter;

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to \"{outputPath}\".");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
