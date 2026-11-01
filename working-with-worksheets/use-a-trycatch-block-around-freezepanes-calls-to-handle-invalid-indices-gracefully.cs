// Title: Safely freeze panes in an Aspose.Cells worksheet with C# by handling invalid row or column indices using try‑catch
// AI Prompts: Write C# code that creates an Aspose.Cells workbook, freezes panes at a specified row and column, and wraps the FreezePanes call in a try‑catch block that catches ArgumentException and logs a clear error message. | Generate a reusable C# method `FreezePaneSafe(Worksheet sheet, int row, int column)` that attempts to freeze panes with Aspose.Cells, returns a success flag, and handles any out‑of‑range argument exceptions.
// Common Searches: C# Aspose.Cells how to catch errors when calling FreezePanes with invalid indices | example of using try‑catch around FreezePanes in Aspose.Cells | prevent ArgumentException for out of range row column in Aspose.Cells FreezePanes | handle invalid freeze pane parameters in Aspose.Cells C# code
// Tags: Aspose.Cells FreezePanes exception handling | C# try-catch worksheet pane freezing | invalid freeze pane indices Aspose.Cells | Aspose.Cells workbook pane freeze error management

using System;
using Aspose.Cells;

// Demonstrates creating an Aspose.Cells workbook, freezing panes at a given row and column, and wrapping the FreezePanes call in try‑catch blocks to gracefully handle ArgumentException for invalid indices and other unexpected errors before saving the file.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook (standard creation rule)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Example indices for freezing panes
            int row = 5;    // Freeze rows above this index
            int column = 3; // Freeze columns to the left of this index

            // Freeze panes using the correct overload (row, column, totalRows, totalColumns)
            // Here we set totalRows and totalColumns to the same values to define the scrollable area.
            sheet.FreezePanes(row, column, row, column);
            Console.WriteLine($"Successfully froze panes at row {row}, column {column}.");

            // Save the workbook (standard saving rule)
            string outputPath = "FrozenPanesExample.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (ArgumentException ex)
        {
            // Handle specific argument errors (e.g., negative indices)
            Console.WriteLine($"Invalid freeze pane indices: {ex.Message}");
        }
        catch (Exception ex)
        {
            // Handle any other unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
