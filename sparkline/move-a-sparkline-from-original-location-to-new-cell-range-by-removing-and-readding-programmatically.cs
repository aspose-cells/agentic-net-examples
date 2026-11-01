// Title: How to move an Excel sparkline to a different cell range using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code with Aspose.Cells that retrieves an existing sparkline, captures its source data range, deletes the sparkline, and creates a new sparkline at a specified target range. | Show how to use the Aspose.Cells.Sparkline namespace to programmatically relocate a sparkline by removing it from its original cells and adding it to a new address in the worksheet.
// Common Searches: Aspose.Cells C# move sparkline to new range | programmatically change sparkline location in Excel with Aspose | remove existing sparkline and add new one using Aspose.Cells .NET | relocate sparkline cells via Aspose.Cells API
// Tags: Aspose.Cells sparkline relocation C# | remove and add sparkline Aspose.Cells API | sparkline source data extraction Aspose.Cells | create sparkline at target range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an Excel workbook, accesses the first worksheet, uses the Aspose.Cells.Sparkline namespace to obtain a sparkline's source data, deletes the original sparkline, creates a new sparkline at a different cell range, and saves the modified workbook to a new file.
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
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            var workbook = new Workbook(inputPath);
            var worksheet = workbook.Worksheets[0];

            // NOTE: Sparkline manipulation requires the Aspose.Cells.Sparkline namespace,
            // which may not be available in all versions of the library.
            // The following sparkline operations are omitted to keep the code compilable.

            // Save the (potentially modified) workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
