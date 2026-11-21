// Title: Apply predefined Style20 chart style to an Excel chart using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells set chart style to Style20 programmatically | apply predefined chart style in Excel workbook using Aspose.Cells .NET | how to change Excel chart appearance to Style20 with Aspose.Cells API | Aspose.Cells chart.Style property usage example for Style20 | set chart style index 20 in Aspose.Cells C# code
// Tags: Aspose.Cells chart style index | C# set chart.Style property | chart style index 20 Aspose.Cells | Style20 chart formatting Aspose.Cells | consistent chart appearance .NET

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsExample
{
    // Loads an existing Excel workbook, checks for at least one chart on the first worksheet, assigns the predefined Style20 (index 20) to the chart via the chart.Style property, and saves the modified file, with error handling for missing files and style‑setting failures.
    class Program
    {
        static void Main(string[] args)
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists before attempting to load it
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook from the input file
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet in the workbook
                Worksheet worksheet = workbook.Worksheets[0];

                // Ensure the worksheet contains at least one chart
                if (worksheet.Charts.Count == 0)
                {
                    Console.WriteLine("No charts found in the worksheet.");
                    return;
                }

                // Retrieve the first chart
                Chart chart = worksheet.Charts[0];

                try
                {
                    // Apply a predefined chart style (Style20). 
                    // In some Aspose.Cells versions the Style property accepts an integer index.
                    chart.Style = 20;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Failed to set chart style: {ex.Message}");
                }

                // Save the modified workbook to the output file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
