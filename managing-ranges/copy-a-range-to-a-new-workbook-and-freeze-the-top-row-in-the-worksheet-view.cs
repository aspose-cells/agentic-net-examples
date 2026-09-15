// Title: Copy a cell range from an existing Excel workbook to a new workbook and freeze the first row with Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code using Aspose.Cells to load source.xlsx, copy the A1:D10 range into a freshly created workbook, and then apply FreezePanes to lock the header row. | Generate a program that creates a destination workbook, pastes a copied range from a source worksheet, and freezes the top row of the sheet with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells copy specific range to another workbook C# | how to freeze the first row after copying data with Aspose.Cells | C# example for copying A1:D10 from one Excel file to a new file using Aspose | using FreezePanes to lock header row in Aspose.Cells worksheet | copy range and preserve formatting to new workbook Aspose.Cells .NET
// Tags: copy range to new workbook Aspose.Cells | freeze first row worksheet Aspose.Cells | Aspose.Cells CreateRange and Copy usage | Aspose.Cells FreezePanes C# example | C# load workbook and save separate file Aspose | preserve formatting when copying Excel range Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program loads source.xlsx, copies the A1:D10 range into a newly created workbook, freezes the top row of the destination worksheet, and saves the result as destination.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        const string sourcePath = "source.xlsx";
        const string destinationPath = "destination.xlsx";

        // Verify source file exists
        if (!File.Exists(sourcePath))
        {
            Console.WriteLine($"Error: Source file \"{sourcePath}\" not found.");
            return;
        }

        try
        {
            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Define the range to copy (example: A1:D10)
            AsposeRange sourceRange = sourceSheet.Cells.CreateRange("A1:D10");

            // Create a new workbook (contains one default worksheet)
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Create a destination range with the same size and copy the source range into it
            AsposeRange destRange = destinationSheet.Cells.CreateRange(
                0, 0, sourceRange.RowCount, sourceRange.ColumnCount);
            destRange.Copy(sourceRange);

            // Freeze the top row in the destination worksheet view
            // FreezePanes(row, column, scrollRow, scrollColumn)
            destinationSheet.FreezePanes(1, 0, 1, 0);

            // Save the new workbook
            destinationWorkbook.Save(destinationPath);
            Console.WriteLine($"Workbook saved successfully to \"{destinationPath}\".");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
