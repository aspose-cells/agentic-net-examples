// Title: Freeze the first ten columns in worksheets that exceed ten columns using Aspose.Cells for .NET
// AI Prompts: Create a helper method that takes a Workbook and freezes columns A‑J on each worksheet where the total column count is greater than ten. | Update the foreach loop to invoke FreezePanes only when sheet.Cells.MaxColumn + 1 > 10, using the overload that freezes the first ten columns. | Add logging to record any worksheet that cannot be frozen and ensure the workbook is saved after processing all sheets.
// Common Searches: c# Aspose.Cells conditional FreezePanes based on column count | how to freeze first ten columns only if worksheet has more than ten columns using .NET | programmatically apply freeze panes to wide Excel sheets with Aspose.Cells | apply FreezePanes to all worksheets in a workbook when column count exceeds 10
// Tags: conditional FreezePanes with Aspose.Cells | freeze first ten columns .NET Excel | worksheet column count check Aspose.Cells | apply FreezePanes based on MaxColumn | Aspose.Cells wide worksheet handling

using System;
using System.IO;
using Aspose.Cells;

// The program loads an Excel workbook, iterates through each worksheet, checks if the sheet contains more than ten columns, and freezes the first ten columns using the FreezePanes method when the condition is met, then saves the updated file.
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

            // Process each worksheet in the workbook
            foreach (Worksheet sheet in workbook.Worksheets)
            {
                try
                {
                    // Determine the last column index that contains data (0‑based)
                    int lastColumnIndex = sheet.Cells.MaxColumn;

                    // If the worksheet has more than ten columns, freeze the first ten columns
                    if (lastColumnIndex + 1 > 10)
                    {
                        // Freeze the first ten columns using the 4‑parameter overload
                        // totalRows and totalColumns define the pane size; using the current max values
                        int totalRows = sheet.Cells.MaxRow + 1;
                        int totalColumns = sheet.Cells.MaxColumn + 1;
                        sheet.FreezePanes(totalRows, totalColumns, 0, 10);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing sheet '{sheet.Name}': {ex.Message}");
                }
            }

            // Save the modified workbook
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved successfully to {outputPath}");
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
