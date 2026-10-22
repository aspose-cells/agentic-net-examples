// Title: How to insert rows before a header row and reset FreezePanes in an Excel worksheet using Aspose.Cells for .NET (C#)
// AI Prompts: Insert two blank rows above row 5 and then update the FreezePanes settings so the header stays frozen using Aspose.Cells in C#. | After adding rows, compute the new freeze row number and call Worksheet.FreezePanes to keep the top rows locked. | Save the workbook to a new file after adjusting the frozen rows with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# insert rows before header and keep freeze panes | how to update FreezePanes after inserting rows in an Excel file using .NET | recalculate frozen row count when adding rows with Aspose.Cells | C# program to add rows and reset freeze panes in Excel workbook | example of Worksheet.FreezePanes after row insertion Aspose.Cells
// Tags: insert rows before header Aspose.Cells | recalculate freeze pane index C# | apply FreezePanes after row insertion Aspose.Cells | save modified workbook Aspose.Cells .NET | adjust frozen rows programmatically Excel

using System;
using System.IO;
using Aspose.Cells;

// The example loads an existing workbook, inserts two rows before the header at row index 4, recalculates the freeze row index, reapplies FreezePanes to keep the header frozen, and saves the updated file as output.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Input and output file paths
            string inputPath = "input.xlsx";
            string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            // Load the existing workbook
            Workbook workbook = new Workbook(inputPath);

            // Get the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // -------------------------------------------------
            // 1. Insert new rows before the header row
            // -------------------------------------------------
            // Assume the header is currently at row index 4 (i.e., Excel row 5)
            int headerRowIndex = 4;          // zero‑based index of the header row
            int rowsToInsert = 2;            // number of rows you want to add

            // Insert the rows; existing rows shift down
            sheet.Cells.InsertRows(headerRowIndex, rowsToInsert);

            // -------------------------------------------------
            // 2. Apply FreezePanes after inserting rows
            // -------------------------------------------------
            // Freeze rows above the header (including any previously frozen rows)
            int newFreezeRow = headerRowIndex + rowsToInsert; // rows to freeze (zero‑based count)
            int freezeColumn = 0; // no columns frozen

            // FreezePanes(row, column, totalRows, totalColumns)
            // row/column specify the first unfrozen cell; totalRows/totalColumns specify how many rows/columns to freeze
            sheet.FreezePanes(newFreezeRow, freezeColumn, newFreezeRow, freezeColumn);

            // -------------------------------------------------
            // 3. Save the modified workbook
            // -------------------------------------------------
            // Ensure the output directory exists
            string outputDir = Path.GetDirectoryName(outputPath);
            if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
            {
                Directory.CreateDirectory(outputDir);
            }

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
