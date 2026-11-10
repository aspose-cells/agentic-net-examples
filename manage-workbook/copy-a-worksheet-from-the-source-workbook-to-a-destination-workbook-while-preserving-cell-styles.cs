// Title: Copy a worksheet from one Excel workbook to another while preserving cell styles using Aspose.Cells for .NET
// Common Searches: how to copy a worksheet from one workbook to another using Aspose.Cells C# | preserve cell formatting when duplicating sheets with Aspose.Cells .NET | Aspose.Cells AddCopy method example for moving worksheets between workbooks | copy first worksheet to an existing workbook without losing styles in C# | duplicate an Excel sheet to another file while keeping formatting using Aspose.Cells
// Tags: copy worksheet between workbooks Aspose.Cells | preserve cell styles during sheet copy | AddCopy method Excel duplication .NET | transfer worksheet with formatting C# Aspose.Cells | merge workbooks retaining formatting Aspose.Cells | duplicate Excel sheet preserving styles

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a source workbook, opens or creates a destination workbook, uses the AddCopy method to copy the first worksheet (including its cell styles) into the destination, and saves the updated file.
    class Program
    {
        static void Main()
        {
            try
            {
                string sourcePath = "source.xlsx";
                string destinationPath = "destination.xlsx";

                // Ensure the source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Load existing destination workbook or create a new one
                Workbook destinationWorkbook = File.Exists(destinationPath)
                    ? new Workbook(destinationPath)
                    : new Workbook();

                // Get the first worksheet from the source workbook
                Worksheet sourceWorksheet = sourceWorkbook.Worksheets[0];

                // Copy the source worksheet into the destination workbook
                destinationWorkbook.Worksheets.AddCopy(sourceWorksheet.Name);

                // Save the destination workbook with the copied worksheet
                destinationWorkbook.Save(destinationPath);
                Console.WriteLine($"Worksheet copied successfully to {destinationPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
