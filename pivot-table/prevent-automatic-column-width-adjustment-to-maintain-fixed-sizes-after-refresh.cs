// Title: How to set fixed column widths in an Excel workbook using Aspose.Cells for .NET so they stay unchanged after a pivot table refresh
// Common Searches: Aspose.Cells C# set column width and prevent auto‑fit after pivot refresh | keep Excel column width constant when refreshing pivot table with Aspose.Cells | C# code to lock column width in workbook using Aspose.Cells | disable automatic column auto‑fit in Aspose.Cells after data refresh | set specific column width in Aspose.Cells .NET without auto‑adjust
// Tags: Aspose.Cells column width configuration .NET | avoid Excel column auto‑fit with Aspose.Cells | maintain fixed column dimensions after pivot refresh Aspose.Cells | C# lock column size in workbook using Aspose.Cells | preserve column width on workbook refresh Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing Excel file with Aspose.Cells, sets explicit widths for columns A and B, and saves the workbook, ensuring the columns retain their fixed sizes and are not auto‑adjusted during subsequent pivot table refreshes.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            try
            {
                // Verify that the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Set a fixed width for column A (index 0)
                Column columnA = worksheet.Cells.Columns[0];
                columnA.Width = 15; // Desired width in characters

                // Set a fixed width for column B (index 1)
                Column columnB = worksheet.Cells.Columns[1];
                columnB.Width = 20; // Desired width in characters

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                // Catch any runtime exceptions and display a friendly message
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
