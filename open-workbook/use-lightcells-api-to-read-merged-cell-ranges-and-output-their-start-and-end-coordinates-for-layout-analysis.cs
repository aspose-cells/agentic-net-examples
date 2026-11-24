// Title: Extract merged cell ranges and their start/end addresses from an Excel workbook using Aspose.Cells in C#
// Common Searches: c# get merged cell ranges from an xlsx file using Aspose.Cells | how to list start and end addresses of merged cells with Aspose.Cells | enumerate merged areas in each worksheet Aspose.Cells C# example | Aspose.Cells read merged cell coordinates for layout analysis | retrieve merged cell ranges and convert to A1 notation in C#
// Tags: Aspose.Cells get merged cell areas | C# enumerate merged ranges in Excel workbook | convert merged cell indices to A1 address Aspose | Excel merged cell extraction using Aspose.Cells | output merged range start end coordinates C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The program loads an Excel workbook with Aspose.Cells, iterates each worksheet, obtains all merged CellArea objects via Cells.GetMergedAreas(), reads their start and end row/column indices, converts those indices to A1‑style addresses using CellsHelper.CellIndexToName, and prints the sheet name along with each merged range (e.g., A1:C3) to the console.
    class Program
    {
        static void Main()
        {
            // Path to the workbook to be analyzed
            string inputPath = "input.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"File not found: {inputPath}");
                return;
            }

            try
            {
                // Load the workbook (standard API, suitable for reading merged cells)
                Workbook workbook = new Workbook(inputPath);

                // Iterate through each worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // Get all merged cell areas in the current worksheet
                    CellArea[] mergedAreas = sheet.Cells.GetMergedAreas();

                    // Iterate through each merged area
                    foreach (CellArea area in mergedAreas)
                    {
                        // Start coordinates (zero‑based indices)
                        int startRow = area.StartRow;
                        int startColumn = area.StartColumn;

                        // End coordinates (zero‑based indices)
                        int endRow = area.EndRow;
                        int endColumn = area.EndColumn;

                        // Convert indices to A1 style addresses for readability
                        string startAddress = CellsHelper.CellIndexToName(startRow, startColumn);
                        string endAddress = CellsHelper.CellIndexToName(endRow, endColumn);

                        // Output the merged range with sheet name
                        Console.WriteLine($"Sheet: {sheet.Name} - Merged Range: {startAddress}:{endAddress}");
                    }
                }
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error processing workbook: {ex.Message}");
            }
        }
    }
}
