// Title: Load an Excel workbook from a UNC network share using Aspose.Cells in C#
// Common Searches: C# Aspose.Cells load Excel workbook from a UNC network share | How to check if an Excel file exists before opening with Aspose.Cells in .NET | Reading cell A1 after opening a workbook from a shared folder using Aspose.Cells | Handling UNC paths when creating a Workbook object with Aspose.Cells in C# | Exception handling for loading a workbook from a network share with Aspose.Cells
// Tags: load workbook from UNC path Aspose.Cells | verify file existence C# Aspose.Cells | read cell A1 Aspose.Cells | network share path handling Aspose.Cells | exception handling workbook load Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example demonstrates how to specify a UNC network share path, verify the Excel file's existence, load it into an Aspose.Cells Workbook via the file‑path constructor, read the value of cell A1 from the first worksheet, and gracefully handle any loading errors.
    class Program
    {
        static void Main(string[] args)
        {
            // UNC network share path to the Excel file.
            string uncPath = @"\\ServerName\ShareName\Folder\MyWorkbook.xlsx";

            try
            {
                // Verify that the file exists before attempting to load.
                if (!File.Exists(uncPath))
                {
                    Console.WriteLine($"File not found: {uncPath}");
                    return;
                }

                // Load the workbook.
                Workbook workbook = new Workbook(uncPath);

                // Read the value of the first cell (A1) in the first worksheet.
                Worksheet sheet = workbook.Worksheets[0];
                string firstCellValue = sheet.Cells["A1"].StringValue;
                Console.WriteLine($"Value of A1: {firstCellValue}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error loading workbook: {ex.Message}");
            }
        }
    }
}
