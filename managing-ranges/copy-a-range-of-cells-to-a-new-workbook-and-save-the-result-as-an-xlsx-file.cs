// Title: Copy a cell range from an existing Excel workbook to a new workbook and save as XLSX using Aspose.Cells for .NET (C#)
// AI Prompts: Load source.xlsx, copy the range A1:C5 to a fresh workbook, and save it as result.xlsx using Aspose.Cells in C#. | Create a new workbook, transfer a defined cell block from another workbook's worksheet, then export the new file as XLSX with Aspose.Cells for .NET.
// Common Searches: Aspose.Cells C# copy specific range to new workbook and save as XLSX | How to duplicate a block of cells from one Excel file to another using Aspose.Cells .NET | C# example for moving a cell range into a separate workbook with Aspose.Cells | Exporting a copied range as its own XLSX file using Aspose.Cells for .NET
// Tags: range copy Aspose.Cells C# | create new workbook from range Aspose.Cells | save workbook as XLSX Aspose.Cells | copy range to separate file Aspose.Cells | Aspose.Cells copy range example

using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program loads source.xlsx, copies a defined cell range from the first worksheet into a newly created workbook, and saves the new workbook as result.xlsx in XLSX format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            const string sourcePath = "source.xlsx";
            const string resultPath = "result.xlsx";

            // Verify that the source file exists before loading
            if (!File.Exists(sourcePath))
            {
                Console.WriteLine($"Source file not found: {sourcePath}");
                return;
            }

            // Load the source workbook containing the data to copy
            Workbook sourceWorkbook = new Workbook(sourcePath);
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];

            // Define the range to copy (A1:C5) using zero‑based indices
            int startRow = 0;
            int startColumn = 0;
            int endRow = 4;
            int endColumn = 2;
            int totalRows = endRow - startRow + 1;
            int totalColumns = endColumn - startColumn + 1;

            // Create a new workbook that will receive the copied range
            Workbook destinationWorkbook = new Workbook();
            Worksheet destinationSheet = destinationWorkbook.Worksheets[0];

            // Create range objects for source and destination using the Aspose alias
            AsposeRange srcRange = sourceSheet.Cells.CreateRange(startRow, startColumn, totalRows, totalColumns);
            AsposeRange destRange = destinationSheet.Cells.CreateRange(0, 0, totalRows, totalColumns);

            // Copy the defined range from the source sheet to the destination sheet
            srcRange.Copy(destRange);

            // Save the new workbook as an XLSX file
            destinationWorkbook.Save(resultPath, SaveFormat.Xlsx);
            Console.WriteLine($"Result workbook saved to {resultPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
