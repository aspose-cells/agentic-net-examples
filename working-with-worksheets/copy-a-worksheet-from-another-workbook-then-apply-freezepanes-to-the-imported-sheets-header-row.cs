// Title: Copy a worksheet from one Excel workbook to another and freeze the header row using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy worksheet from one workbook to another | how to freeze the first row after copying a sheet with Aspose.Cells | using AddCopy to duplicate a sheet and apply FreezePanes in .NET | remove default worksheet before copying sheet in Aspose.Cells | set FreezePanes to A2 on copied worksheet using Aspose.Cells for .NET
// Tags: Aspose.Cells AddCopy worksheet duplication | freeze panes after sheet copy Aspose.Cells | remove default sheet Aspose.Cells .NET | set frozen header row Aspose.Cells | copy Excel sheet between workbooks C#

using Aspose.Cells;
using System;
using System.IO;

// The example loads a source Excel file, creates a new workbook, clears its default sheet, copies the first worksheet using the AddCopy method, applies FreezePanes to lock the first row (making A2 the top‑left scrollable cell), and saves the destination file, with error handling for missing files.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "Source.xlsx";
            const string destinationPath = "Destination.xlsx";

            // Verify that the source workbook exists to avoid FileNotFoundException
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the source workbook containing the worksheet to copy
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new (empty) destination workbook
            Workbook destinationWorkbook = new Workbook();

            // Remove the default sheet that Aspose.Cells creates in a new workbook
            destinationWorkbook.Worksheets.Clear();

            // Copy the first worksheet from the source workbook to the destination workbook
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            // AddCopy expects the sheet name, returns the index of the newly added sheet
            int copiedIndex = destinationWorkbook.Worksheets.AddCopy(sourceSheet.Name);
            Worksheet copiedSheet = destinationWorkbook.Worksheets[copiedIndex];

            // Freeze the first row (header). FreezePanes(row, column, frozenRows, frozenColumns)
            // Top‑left cell of the scrollable area is A2 (row 2, column 1)
            copiedSheet.FreezePanes(2, 1, 1, 0);

            // Save the destination workbook with the copied and frozen sheet
            destinationWorkbook.Save(destinationPath);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
