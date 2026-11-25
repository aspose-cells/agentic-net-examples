// Title: Reorder Excel worksheets alphabetically and freeze the header row using Aspose.Cells for .NET (C#)
// Common Searches: C# code to sort workbook worksheets alphabetically with Aspose.Cells | How to apply FreezePanes to the first row of every sheet in Aspose.Cells .NET | Programmatically move Excel worksheets to a specific order using Aspose.Cells | Freeze top header row across all worksheets in a .xlsx file with Aspose.Cells C# | Aspose.Cells example for reordering sheets and freezing rows in a workbook
// Tags: alphabetical worksheet reordering Aspose.Cells .NET | apply FreezePanes to multiple worksheets | move worksheets to specific index C# | sort Excel sheets by name programmatically | process workbook sheet order and freeze rows

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;

// The sample loads an Excel workbook with Aspose.Cells, sorts all worksheets alphabetically by name using MoveTo, then freezes the first row of each sheet with FreezePanes(1,0,1,0), and finally saves the updated file.
class Program
{
    static void Main(string[] args)
    {
        // Input and output file paths
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify input file exists
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Error: Input file \"{inputPath}\" not found.");
            return;
        }

        try
        {
            // Load the workbook
            using (Workbook workbook = new Workbook(inputPath))
            {
                // -------------------------------------------------
                // 1. Reorder worksheets alphabetically by name
                // -------------------------------------------------
                // Get all worksheets into a list
                List<Worksheet> allSheets = workbook.Worksheets.Cast<Worksheet>().ToList();

                // Sort the list by worksheet name
                List<Worksheet> sortedSheets = allSheets
                    .OrderBy(ws => ws.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();

                // Move each worksheet to its new index according to the sorted order
                for (int targetIndex = 0; targetIndex < sortedSheets.Count; targetIndex++)
                {
                    Worksheet ws = sortedSheets[targetIndex];
                    // If the worksheet is already at the correct position, skip
                    if (ws.Index != targetIndex)
                    {
                        ws.MoveTo(targetIndex);
                        // After moving, the indices of other worksheets shift.
                        // Refresh the sorted list to reflect new indices for subsequent moves.
                        sortedSheets = workbook.Worksheets.Cast<Worksheet>()
                            .OrderBy(w => w.Name, StringComparer.OrdinalIgnoreCase)
                            .ToList();
                    }
                }

                // -------------------------------------------------
                // 2. Freeze the top row of each worksheet
                // -------------------------------------------------
                foreach (Worksheet ws in workbook.Worksheets)
                {
                    // Freeze the first row (row index 1 corresponds to the second row, i.e., split just below row 1)
                    // Parameters: row, column, freezedRows, freezedColumns
                    // Freeze 1 row, 0 columns.
                    ws.FreezePanes(1, 0, 1, 0);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook processed successfully. Saved as \"{outputPath}\".");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred while processing the workbook:");
            Console.WriteLine(ex.Message);
        }
    }
}
