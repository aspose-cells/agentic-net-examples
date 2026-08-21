// Title: C# Example: Delete Rows Above Frozen Pane and Reset FreezePanes with Aspose.Cells
// Description: Demonstrates how to load a workbook, detect existing frozen panes, delete all rows above the frozen row, recalculate the freeze coordinates, reapply FreezePanes, and save the updated file using Aspose.Cells for .NET.
// Keywords: Aspose.Cells C# | FreezePanes | delete rows above frozen pane | adjust frozen row after deletion | GetFreezedPanes | worksheet freeze settings | Aspose.Cells .NET example | preserve frozen header | row deletion API
// Common Searches: Aspose.Cells delete rows above frozen pane | How to adjust FreezePanes after deleting rows in C# | Get frozen pane indices Aspose.Cells .NET | Reapply FreezePanes after row removal | C# code to reset freeze panes after row deletion
// Developer Intent: Remove rows that precede the frozen area and update the FreezePanes parameters so the frozen region stays correctly positioned.
// Use Cases: Clean up a worksheet by removing header rows while keeping the frozen header row for scrolling. | Prepare a template where rows above the frozen pane must be stripped before exporting to another format. | Update a generated report that contains frozen panes, ensuring the freeze position remains accurate after row deletions.
// AI Prompts: Write C# code with Aspose.Cells to delete all rows above the current frozen row and update the FreezePanes settings. | Explain how to retrieve frozen pane parameters, delete rows, and reapply FreezePanes in Aspose.Cells for .NET. | Provide robust error handling best practices when adjusting freeze panes after row deletions using Aspose.Cells.

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to load a workbook, detect existing frozen panes, delete all rows above the frozen row, recalculate the freeze coordinates, reapply FreezePanes, and save the updated file using Aspose.Cells for .NET.
    public class AdjustFreezeAfterRowDeletion
    {
        public static void Run()
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            Workbook workbook = null;

            try
            {
                // Load existing workbook or create a new one if the file is missing
                if (File.Exists(inputPath))
                {
                    workbook = new Workbook(inputPath);
                }
                else
                {
                    Console.WriteLine($"Input file '{inputPath}' not found. Creating a new workbook.");
                    workbook = new Workbook();
                }

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Retrieve current freeze pane settings
                int frozenRow, frozenColumn, frozenRowsCount, frozenColumnsCount;
                bool hasFreeze = sheet.GetFreezedPanes(out frozenRow, out frozenColumn, out frozenRowsCount, out frozenColumnsCount);

                if (!hasFreeze)
                {
                    Console.WriteLine("No frozen panes found. Saving workbook without changes.");
                    workbook.Save(outputPath);
                    return;
                }

                // Delete rows above the frozen row
                int rowsToDelete = frozenRow; // rows 0 to frozenRow-1
                if (rowsToDelete > 0)
                {
                    sheet.Cells.DeleteRows(0, rowsToDelete);

                    // Adjust frozen row index after deletion
                    int newFrozenRow = frozenRow - rowsToDelete;

                    // Reapply freeze panes with the adjusted row index
                    sheet.FreezePanes(newFrozenRow, frozenColumn, frozenRowsCount, frozenColumnsCount);
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            AdjustFreezeAfterRowDeletion.Run();
        }
    }
}
