// Title: Reposition a worksheet to the third tab and freeze its first column using Aspose.Cells for .NET (C#)
// Common Searches: c# aspose.cells move worksheet to third tab | how to freeze the first column in an Excel file with Aspose.Cells | aspose.cells reorder worksheets and apply FreezePanes programmatically | move worksheet to specific index and freeze column in .net workbook
// Tags: worksheet MoveTo method Aspose.Cells | first column FreezePanes Aspose.Cells | reorder worksheet tabs .NET Excel | Aspose.Cells worksheet repositioning C# | column freeze pane Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace WorksheetManipulation
{
    // The example loads an existing workbook, locates the worksheet named 'DataSheet', moves it to the third tab (zero‑based index 2) using the MoveTo method, freezes only the first column with FreezePanes(0, 1, 0, 1), and saves the modified file as output.xlsx.
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
                    Console.WriteLine($"Input file not found: {Path.GetFullPath(inputPath)}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Get the worksheet to reposition (by name)
                Worksheet sheet = workbook.Worksheets["DataSheet"];
                if (sheet == null)
                {
                    Console.WriteLine("Worksheet 'DataSheet' not found.");
                    return;
                }

                // Move the worksheet to the third tab (zero‑based index 2)
                sheet.MoveTo(2);

                // Freeze the first column (no row freeze, freeze 1 column)
                // FreezePanes(row, column, totalRows, totalColumns)
                sheet.FreezePanes(0, 1, 0, 1);

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Log any unexpected errors
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
