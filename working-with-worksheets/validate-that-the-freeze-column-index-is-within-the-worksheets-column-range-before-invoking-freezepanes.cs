// Title: How to validate a freeze column index against a worksheet’s column range before using Worksheet.FreezePanes in Aspose.Cells for .NET
// AI Prompts: Verify that a zero‑based column index is less than the worksheet’s used column count before calling Worksheet.FreezePanes. | Add conditional logic to skip FreezePanes when the specified column index exceeds the maximum column in the sheet. | Implement error handling that logs a message if the freeze column index is out of range and prevents an exception.
// Common Searches: Aspose.Cells C# check column index bounds before FreezePanes | prevent out of range error when freezing panes in Aspose.Cells workbook | how to get used column count in Aspose.Cells and validate freeze column | C# example for validating freeze column index with Worksheet.FreezePanes
// Tags: freeze panes column index validation Aspose.Cells | worksheet used column count C# | prevent out of range FreezePanes exception | dynamic total rows columns Aspose.Cells | error handling for FreezePanes bounds

using System;
using Aspose.Cells;

// // Demonstrates checking that a zero‑based freeze column index lies within the worksheet's used column count before invoking Worksheet.FreezePanes, with fallback handling for out‑of‑range values.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Desired freeze column index (zero‑based)
            int freezeColumnIndex = 5;

            // Determine the number of columns that contain data
            int usedColumnCount = sheet.Cells.MaxColumn + 1; // MaxColumn is zero‑based

            // Validate the index is within the worksheet's column range
            if (freezeColumnIndex >= 0 && freezeColumnIndex < usedColumnCount)
            {
                // Freeze panes at row 0 and the specified column.
                // The overload requires total rows and columns for the scrollable area.
                int totalRows = sheet.Cells.MaxRow + 1;
                int totalColumns = sheet.Cells.MaxColumn + 1;
                sheet.FreezePanes(0, freezeColumnIndex, totalRows, totalColumns);
            }
            else
            {
                // Index out of range – handle as needed
                Console.WriteLine("Freeze column index is out of the worksheet's column range.");
            }

            // Save the workbook
            workbook.Save("output.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
