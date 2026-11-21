// Title: Delete cells in range F1:F20 and shift remaining rows upward with Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells delete cells F1 to F20 and move rows up | Remove a column range and shift rows upward using Aspose.Cells .NET | How to delete rows 1‑20 in Excel workbook with Aspose.Cells and preserve data | Aspose.Cells delete specific cell range and shift remaining cells up in C# | Shift cells up after deleting range F1:F20 with Aspose.Cells API
// Tags: Aspose.Cells DeleteRows shift up | C# delete Excel column F range | remove first twenty rows Aspose.Cells | shift remaining cells upward Aspose.Cells | Excel range deletion using Aspose.Cells .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an Excel workbook, creates a range for cells F1:F20, deletes those rows using worksheet.Cells.DeleteRows (which shifts subsequent rows upward), and saves the updated file.
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
                var workbook = new Workbook(inputPath);
                var worksheet = workbook.Worksheets[0];

                // Define the range F1:F20 (column F, rows 1‑20)
                var range = worksheet.Cells.CreateRange("F1", "F20");

                // Delete the cells in the range and shift remaining cells upward.
                // Aspose.Cells does not provide DeleteRows on Range, so we delete the rows
                // directly on the worksheet's Cells collection. This removes rows 1‑20
                // and shifts the rest of the sheet up.
                worksheet.Cells.DeleteRows(0, range.RowCount, true);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any unexpected errors
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
