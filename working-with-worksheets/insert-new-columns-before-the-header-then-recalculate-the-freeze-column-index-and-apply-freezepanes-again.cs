// Title: Insert columns before the header row and recalculate the freeze pane column index using Aspose.Cells for .NET
// AI Prompts: Add a specified number of columns at index 0 of the first worksheet and then call FreezePanes to lock those columns with Aspose.Cells. | Determine the correct column index for FreezePanes after the column insertion and apply the freeze pane based on that calculated index. | Write the modified workbook to a new file after updating the column layout and reapplying the freeze pane.
// Common Searches: how to add columns at the beginning of a worksheet and preserve freeze panes with Aspose.Cells C# | recalculate freeze pane column index after inserting columns in an Excel file using Aspose.Cells .NET | Aspose.Cells example for inserting columns before header and setting FreezePanes on new columns
// Tags: insert columns before header Aspose.Cells | reapply FreezePanes after column insertion .NET | calculate freeze column index Aspose.Cells | modify worksheet layout C# Aspose.Cells | freeze newly added columns Excel Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The program loads an existing Excel workbook, inserts two columns at the very left of the first worksheet, recomputes the column index to freeze, reapplies FreezePanes to lock the newly added columns, and saves the result to a new file.
class Program
{
    static void Main()
    {
        // Paths for input and output files
        string inputPath = "input.xlsx";
        string outputPath = "output.xlsx";

        // Verify that the input file exists to avoid FileNotFoundException
        if (!File.Exists(inputPath))
        {
            Console.WriteLine($"Input file not found: {inputPath}");
            return;
        }

        try
        {
            // Load the workbook
            Workbook workbook = new Workbook(inputPath);

            // Access the first worksheet (adjust index if needed)
            Worksheet sheet = workbook.Worksheets[0];

            // Number of new columns to insert before the header
            int insertCount = 2; // adjust as required

            // Insert new columns at the very beginning (column index 0)
            sheet.Cells.InsertColumns(0, insertCount);

            // Reapply freeze panes.
            // Freeze the first 'insertCount' columns (row = 0, column = insertCount).
            // The last two parameters (freezableRows, freezableColumns) are set to 0
            // because we only want to freeze columns, not rows.
            sheet.FreezePanes(0, insertCount, 0, 0);

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Handle any runtime errors (e.g., loading, saving, processing)
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
