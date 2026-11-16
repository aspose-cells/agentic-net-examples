// Title: How to freeze the first row and first column together using Aspose.Cells FreezePanes in C#
// Common Searches: Aspose.Cells C# freeze top row and left column in the same worksheet | C# code to freeze row 1 and column A simultaneously with Aspose.Cells | Using FreezePanes method to lock first row and first column in an Excel file with Aspose.Cells | Freeze panes for first row and column in .NET Excel workbook using Aspose.Cells | Set freeze panes at cell B2 with Aspose.Cells C#
// Tags: Aspose.Cells FreezePanes method example | C# Excel workbook freeze panes configuration | freeze panes starting at B2 Aspose.Cells | worksheet freeze top row and left column .NET | Excel file freeze panes API C#

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // This C# example creates a new Workbook, accesses the first worksheet, and calls Worksheet.FreezePanes(1, 1, 1, 1) to freeze the top row and leftmost column simultaneously, then saves the workbook as output.xlsx with basic error handling.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Get the first worksheet
                Worksheet sheet = workbook.Worksheets[0];

                // Freeze the first row and first column simultaneously.
                // Parameters: row, column (zero‑based index of the cell after the frozen panes),
                // totalRows, totalColumns (number of rows/columns to freeze).
                sheet.FreezePanes(1, 1, 1, 1);

                // Define output file path
                string outputPath = "output.xlsx";

                // Ensure the directory for the output file exists (if any)
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook to a file
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Log the exception details
                Console.Error.WriteLine($"An error occurred: {ex.Message}");
                Console.Error.WriteLine(ex.StackTrace);
            }
        }
    }
}
