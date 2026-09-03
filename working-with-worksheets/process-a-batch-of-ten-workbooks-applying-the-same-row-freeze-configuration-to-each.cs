// Title: Freeze the first row in all worksheets of multiple Excel workbooks using Aspose.Cells for .NET
// AI Prompts: Write C# code that loads a collection of .xlsx file paths, iterates through each worksheet, and applies FreezePanes to lock the top row with Aspose.Cells. | Create a reusable function that accepts workbook paths and customizable frozen‑row/column counts, then applies the same FreezePanes settings to every sheet in each workbook. | Add robust logging and error handling to a batch‑processing script so it skips missing files, continues with the remaining workbooks, and saves changes in place.
// Common Searches: Aspose.Cells C# batch freeze first row across multiple Excel files | How to apply FreezePanes to every worksheet in a set of workbooks using .NET | C# loop through list of .xlsx files and set identical freeze pane configuration with Aspose.Cells | Programmatically lock top row in all sheets of several Excel workbooks in C# | Error handling for missing Excel files when using Aspose.Cells FreezePanes in a batch script
// Tags: batch freeze panes Aspose.Cells C# | freeze top row multiple workbooks .NET | apply FreezePanes to all worksheets programmatically | iterate workbook collection Aspose.Cells | skip missing Excel files error handling

using System;
using System.Collections.Generic;
using System.IO;
using Aspose.Cells;

// The C# program iterates over a list of Excel workbook paths, loads each file with Aspose.Cells, applies FreezePanes to lock the first row on every worksheet, saves the changes back to the original files, and logs success while gracefully handling missing files.
class BatchFreezeRows
{
    static void Main()
    {
        // List of workbook file paths to process (adjust paths as needed)
        List<string> workbookPaths = new List<string>
        {
            @"C:\Workbooks\Book1.xlsx",
            @"C:\Workbooks\Book2.xlsx",
            @"C:\Workbooks\Book3.xlsx",
            @"C:\Workbooks\Book4.xlsx",
            @"C:\Workbooks\Book5.xlsx",
            @"C:\Workbooks\Book6.xlsx",
            @"C:\Workbooks\Book7.xlsx",
            @"C:\Workbooks\Book8.xlsx",
            @"C:\Workbooks\Book9.xlsx",
            @"C:\Workbooks\Book10.xlsx"
        };

        // Define the row and column indices to start freezing (0‑based).
        // To freeze the first row, set freezeRow = 1 and freezeColumn = 0.
        int freezeRow = 1;
        int freezeColumn = 0;
        // Number of rows and columns to keep frozen (1 row, 0 columns in this case)
        int frozenRows = 1;
        int frozenColumns = 0;

        foreach (string path in workbookPaths)
        {
            try
            {
                if (!File.Exists(path))
                {
                    Console.WriteLine($"File not found: {path}");
                    continue;
                }

                // Load the workbook from file
                Workbook workbook = new Workbook(path);

                // Apply the same freeze configuration to every worksheet in the workbook
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    // FreezePanes(row, column, totalRows, totalColumns)
                    sheet.FreezePanes(freezeRow, freezeColumn, frozenRows, frozenColumns);
                }

                // Save the workbook back to the same file (overwrites original)
                workbook.Save(path, SaveFormat.Xlsx);
                Console.WriteLine($"Processed: {path}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error processing '{path}': {ex.Message}");
            }
        }
    }
}
