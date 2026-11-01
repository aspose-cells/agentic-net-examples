// Title: Copy only cell values from a source range to a destination range in another worksheet with Aspose.Cells for .NET (C#)
// Common Searches: how to copy only the values of a range to another worksheet using Aspose.Cells in C# | Aspose.Cells copy range data without formatting between workbooks | C# code to transfer Excel cells A1:C10 to E1 without styles using Aspose.Cells | skip cell formatting when moving data between worksheets with Aspose.Cells .NET | programmatically copy Excel range values only with Aspose.Cells API
// Tags: copy cell values Aspose.Cells C# | range value transfer without formatting | Aspose.Cells copy values between worksheets | Excel range data migration C# Aspose.Cells | PutValue method Aspose.Cells example | transfer worksheet data programmatically

using Aspose.Cells;
using System;
using System.IO;

// The example loads a source workbook, defines a source range (A1:C10), creates or opens a destination workbook, and iterates through the range copying each cell's value with PutValue to a target range starting at E1. No formatting, formulas, or comments are transferred, and the result is saved as a new Excel file.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string destinationPath = "destination.xlsx";

            // Verify source file exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file \"{sourcePath}\" not found.");
                return;
            }

            // Load the source workbook containing the data to copy
            Workbook srcWorkbook = new Workbook(sourcePath);

            // Get the first worksheet (adjust index if needed)
            Worksheet srcWorksheet = srcWorkbook.Worksheets[0];

            // Define the source range to copy (e.g., cells A1:C10)
            Aspose.Cells.Range srcRange = srcWorksheet.Cells.CreateRange("A1", "C10");

            // Create a new workbook for the destination (or load an existing one)
            Workbook destWorkbook = new Workbook(); // starts with a default worksheet
            Worksheet destWorksheet = destWorkbook.Worksheets[0];

            // Define the top‑left cell of the destination range (e.g., cell E1)
            int destStartRow = 0;      // Row index for E1 (zero‑based)
            int destStartColumn = 4;   // Column index for E (A=0, B=1, ...)

            // Transfer only the cell values from source to destination
            for (int i = 0; i < srcRange.RowCount; i++)
            {
                for (int j = 0; j < srcRange.ColumnCount; j++)
                {
                    // Get source cell
                    Cell srcCell = srcRange[i, j];

                    // Get corresponding destination cell
                    Cell destCell = destWorksheet.Cells[destStartRow + i, destStartColumn + j];

                    // Copy only the value (no formatting, formulas, comments, etc.)
                    destCell.PutValue(srcCell.Value);
                }
            }

            // Save the destination workbook
            destWorkbook.Save(destinationPath);
            Console.WriteLine($"Data copied successfully to \"{destinationPath}\".");
        }
        catch (Exception ex)
        {
            // Log unexpected errors
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
