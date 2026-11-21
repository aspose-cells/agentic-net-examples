// Title: Copy a worksheet to a new workbook and freeze its header row using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# copy worksheet from source workbook to new workbook and freeze first row | C# example using Worksheets.AddCopy and FreezePanes together in Aspose.Cells | How to freeze header row after duplicating a sheet with Aspose.Cells for .NET | Copy first sheet of source.xlsx to result.xlsx and lock header using Aspose.Cells | Apply FreezePanes to imported worksheet in Aspose.Cells C# code
// Tags: Worksheets.AddCopy method Aspose.Cells | FreezePanes API C# | copy sheet between workbooks Aspose.Cells | freeze header row Excel .NET | load and save workbook Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

// The sample loads a source Excel file, creates an empty destination workbook, copies the first worksheet using Worksheets.AddCopy, applies FreezePanes to lock the first row as a header, and saves the result, with error handling for missing files.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string resultPath = "result.xlsx";

            // Verify that the source file exists before attempting to load it
            if (!File.Exists(sourcePath))
                throw new FileNotFoundException($"Source file not found: {sourcePath}");

            // Load the source workbook that contains the worksheet to copy
            Workbook sourceWorkbook = new Workbook(sourcePath);

            // Create a new workbook (empty destination workbook)
            Workbook destinationWorkbook = new Workbook();

            // Copy the first worksheet from the source workbook into the destination workbook
            // AddCopy expects the name of the source worksheet
            int copiedIndex = destinationWorkbook.Worksheets.AddCopy(sourceWorkbook.Worksheets[0].Name);

            // Get a reference to the newly copied worksheet
            Worksheet copiedSheet = destinationWorkbook.Worksheets[copiedIndex];

            // Freeze the first row (header). Use the 4‑parameter overload:
            // row = 2 (first row below the frozen area), column = 0 (no column freeze),
            // totalRows = 1 (freeze one row), totalColumns = 0 (no column freeze)
            copiedSheet.FreezePanes(2, 0, 1, 0);

            // Save the destination workbook with the copied and modified worksheet
            destinationWorkbook.Save(resultPath);
            Console.WriteLine($"Workbook saved successfully to '{resultPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
