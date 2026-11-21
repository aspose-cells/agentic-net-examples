// Title: Copy only formatting from a range offset five rows upward back to the original range using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# copy cell formatting from an offset range back to the original range | how to shift a range up by rows and paste only formats with Aspose.Cells | C# example of using PasteOptions Formats to transfer styling between ranges in Aspose.Cells | copy formatting between two ranges after applying a row offset in Aspose.Cells | offset a range upward and apply its style to the original cells using Aspose.Cells .NET
// Tags: offset range rows Aspose.Cells .NET | copy formatting between ranges Aspose.Cells | PasteOptions Formats Aspose.Cells | create range with row offset C# | apply formatting from shifted range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExample
{
    // The code loads an Excel workbook, defines a source range (B2:D10), creates a second range shifted five rows upward, and uses PasteOptions with PasteType.Formats to copy only the formatting from the offset range back to the original range before saving the file.
    class Program
    {
        static void Main(string[] args)
        {
            // Define input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Define the original range (adjust the address as needed)
                AsposeRange originalRange = sheet.Cells.CreateRange("B2:D10");

                // Calculate the offset range (5 rows upward)
                int offsetRows = -5;
                int offsetStartRow = Math.Max(0, originalRange.FirstRow + offsetRows);
                AsposeRange offsetRange = sheet.Cells.CreateRange(
                    offsetStartRow,
                    originalRange.FirstColumn,
                    originalRange.RowCount,
                    originalRange.ColumnCount);

                // Prepare paste options to transfer only formatting
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.Formats
                };

                // Copy the formatting from the offset range back to the original range
                offsetRange.Copy(originalRange, pasteOptions);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
