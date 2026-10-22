// Title: C# example: unfreeze existing panes, insert rows, and reapply FreezePanes with adjusted indices using Aspose.Cells
// AI Prompts: Generate C# code that removes any frozen panes, inserts a specific number of rows at a given position, and then sets new FreezePanes coordinates with Aspose.Cells. | Show how to recalculate FreezePanes row and column offsets after inserting rows in an Excel worksheet using Aspose.Cells for .NET. | Provide a step‑by‑step C# snippet that unfreezes panes, inserts rows, and re‑freezes the top N rows and left M columns in a workbook.
// Common Searches: Aspose.Cells C# refresh frozen pane after inserting rows | How to adjust FreezePanes row index after adding rows with Aspose.Cells | Unfreeze then reapply FreezePanes in .NET Excel workbook | Insert rows and keep first two rows frozen using Aspose.Cells | C# Aspose.Cells FreezePanes parameters after row insertion
// Tags: Aspose.Cells unfreeze panes C# | Aspose.Cells insert rows worksheet | Aspose.Cells FreezePanes after row insertion | Aspose.Cells update frozen pane indices | Aspose.Cells save workbook Excel

using System;
using System.IO;
using Aspose.Cells;

// The sample loads an existing Excel file, clears any frozen panes, inserts three rows after a specified index, then freezes the first two rows and columns using updated FreezePanes coordinates, and finally saves the modified workbook.
class Program
{
    static void Main()
    {
        try
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Ensure the input file exists before loading
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the workbook
            var workbook = new Workbook(inputPath);
            var worksheet = workbook.Worksheets[0];

            // Unfreeze any existing frozen panes
            worksheet.FreezePanes(0, 0, 0, 0);

            // Insert rows (example: insert 3 rows after row index 5)
            int insertRowIndex = 5;   // zero‑based index where rows will be inserted
            int rowsToInsert = 3;
            worksheet.Cells.InsertRows(insertRowIndex, rowsToInsert);

            // Reapply FreezePanes with updated indices
            // Example: freeze the first 2 rows and first 2 columns after insertion
            int frozenRows = 2;
            int frozenColumns = 2;
            // Parameters: row, column, topRow, leftColumn
            worksheet.FreezePanes(frozenRows, frozenColumns, frozenRows, frozenColumns);

            // Save the workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
