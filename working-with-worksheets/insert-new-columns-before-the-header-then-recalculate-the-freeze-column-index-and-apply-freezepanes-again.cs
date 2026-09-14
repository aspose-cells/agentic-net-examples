// Title: Insert columns before the header row and refresh freeze panes in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Insert two new columns at the very left of a worksheet and adjust the FreezePanes call to freeze the newly added columns with Aspose.Cells in C#. | Recalculate the freeze column index after inserting columns and apply the four‑parameter FreezePanes method while preserving any existing frozen rows. | Programmatically add columns before the header and update the frozen pane settings in an existing .xlsx file using Aspose.Cells for .NET.
// Common Searches: Aspose.Cells insert columns at index 0 and keep freeze panes | C# update frozen columns after inserting new columns in Excel workbook | How to recalculate freeze pane column index with Aspose.Cells after column insertion | Preserve freeze panes when adding columns before header using Aspose.Cells for .NET
// Tags: insert columns before header Aspose.Cells C# | reapply FreezePanes after column insertion Aspose.Cells | adjust freeze pane column index .NET Excel | modify worksheet layout programmatically Aspose.Cells | update frozen columns in .xlsx using Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, inserts two columns at the beginning of the first worksheet, recalculates the freeze pane to include the new columns, and saves the updated file.
class Program
{
    static void Main()
    {
        try
        {
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Number of new columns to insert before the header column
            int columnsToInsert = 2;

            // Insert the columns at index 0 (before the existing header column)
            sheet.Cells.InsertColumns(0, columnsToInsert);

            // Reapply freeze panes: keep any existing frozen rows (assumed 0) and freeze the new columns
            // Use the overload with four parameters: row, column, totalRows, totalColumns
            sheet.FreezePanes(0, columnsToInsert, 0, columnsToInsert);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
