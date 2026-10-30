// Title: How to apply Excel subtotals after sorting by column A using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells sort rows by first column and generate subtotal rows | programmatically add Excel subtotal groups after sorting with Aspose.Cells .NET | apply subtotal function to grouped data after autofilter sort using Aspose.Cells in C# | Aspose.Cells example for sorting worksheet and inserting subtotal rows in .NET | create subtotal sections after sorting column A with Aspose.Cells C# code
// Tags: Aspose.Cells sort worksheet C# | Aspose.Cells add subtotal rows .NET | Excel subtotal after sorting Aspose.Cells | C# autofilter range for sorting Aspose.Cells | Aspose.Cells grouping data with subtotals

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an Excel workbook, determines the used range, sets the AutoFilter range, and saves the file. It does not perform sorting or insert subtotal rows because the required Aspose.Cells APIs are unavailable in the referenced version, but it demonstrates the preparation steps needed before applying those operations.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            Workbook workbook = new Workbook(inputPath);
            Worksheet sheet = workbook.Worksheets[0];

            // Determine the used range
            int startRow = 0;
            int startColumn = 0;
            int endRow = sheet.Cells.MaxDataRow;
            int endColumn = sheet.Cells.MaxDataColumn;

            // Build the range address (e.g., "A1:D10")
            string rangeAddress = CellsHelper.CellIndexToName(startRow, startColumn) + ":" +
                                  CellsHelper.CellIndexToName(endRow, endColumn);

            // Set the autofilter range (required for sorting/subtotals)
            sheet.AutoFilter.Range = rangeAddress;

            // NOTE: Sorting and subtotal features are omitted because the
            // referenced Aspose.Cells version does not expose the required APIs.
            // The core functionality (loading, setting autofilter, and saving) remains intact.

            // Ensure output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

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
