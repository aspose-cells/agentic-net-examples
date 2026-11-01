// Title: Copy columns A‑D from a source worksheet to a new workbook with Aspose.Cells range copy in C#
// Common Searches: c# Aspose.Cells copy columns A to D from one Excel file to another | how to transfer first four columns preserving formulas using Aspose.Cells | efficiently copy a column range between workbooks in C# | Aspose.Cells example for copying A:D range to a new workbook | C# code to copy selected columns and keep formatting with Aspose.Cells
// Tags: Aspose.Cells range copy columns | copy A-D columns between workbooks C# | create new workbook transfer column range Aspose | C# Excel column range copy with formatting | Aspose.Cells copy selected columns efficiently

using System;
using System.IO;
using Aspose.Cells;

// The program loads a source workbook, creates an empty destination workbook, determines the last used row, defines a range covering columns A through D for all used rows, copies the range (including values, formulas, and formatting) to the destination worksheet, and saves the new workbook.
class Program
{
    static void Main()
    {
        try
        {
            string sourcePath = "SourceWorkbook.xlsx";
            string destPath = "DestinationWorkbook.xlsx";

            // Ensure the source file exists
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the source workbook
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new (empty) destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Get the first worksheet from each workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Determine the last used row in the source sheet
            int lastRow = sourceSheet.Cells.MaxDataRow; // zero‑based index

            // If there is no data, nothing to copy
            if (lastRow < 0)
            {
                Console.WriteLine("Source worksheet contains no data to copy.");
                destinationWorkbook.Save(destPath);
                return;
            }

            // Define the source range: columns A (0) through D (3) and all used rows
            // +1 because CreateRange uses count, not end index
            Aspose.Cells.Range sourceRange = sourceSheet.Cells.CreateRange(0, 0, lastRow + 1, 4);

            // Define the destination range starting at A1 with the same size
            Aspose.Cells.Range destRange = destinationSheet.Cells.CreateRange(0, 0, lastRow + 1, 4);

            // Copy values, formulas, formats, etc.
            sourceRange.Copy(destRange);

            // Save the new workbook
            destinationWorkbook.Save(destPath);
            Console.WriteLine($"Data copied successfully to {destPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
