// Title: Copy a worksheet from one Excel workbook to another while preserving cell styles using Aspose.Cells for .NET (C#)
// Common Searches: how to copy a worksheet from one workbook to another with formatting using Aspose.Cells C# | Aspose.Cells copy worksheet preserving cell styles to a new workbook | C# duplicate Excel sheet and keep cell formatting with Aspose.Cells | using Worksheets.AddCopy to transfer a sheet while retaining styles in .NET | example of moving first worksheet to another file with formatting using Aspose.Cells
// Tags: Worksheets.AddCopy method preserving formatting | copy worksheet between workbooks Aspose.Cells | duplicate Excel sheet with styles C# | transfer worksheet content and formatting .NET | preserve cell styles during workbook copy Aspose

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a source Excel file, creates an empty destination workbook, and uses Worksheets.AddCopy to copy the first worksheet—including all cell content and styles—to the new workbook, then saves it.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string sourcePath = "SourceWorkbook.xlsx";
                string destinationPath = "DestinationWorkbook.xlsx";

                // Verify source file exists
                if (!File.Exists(sourcePath))
                {
                    Console.WriteLine($"Source file not found: {sourcePath}");
                    return;
                }

                // Load the source workbook
                Workbook sourceWorkbook = new Workbook(sourcePath);

                // Create a new (empty) destination workbook
                Workbook destinationWorkbook = new Workbook();

                // Get the first worksheet from the source workbook
                Worksheet sourceWorksheet = sourceWorkbook.Worksheets[0];

                // Copy the worksheet to the destination workbook preserving all content and formatting
                destinationWorkbook.Worksheets.AddCopy(
                    new Worksheet[] { sourceWorksheet },
                    new string[] { sourceWorksheet.Name });

                // Save the destination workbook
                destinationWorkbook.Save(destinationPath);
                Console.WriteLine($"Destination workbook saved to: {destinationPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
