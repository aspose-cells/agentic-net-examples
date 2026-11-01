// Title: Replace 'Pending' with 'Completed' in Excel cells D5:D25 using Aspose.Cells for .NET
// Common Searches: Aspose.Cells C# find and replace text in a specific cell range | How to bulk update cells D5 to D25 from 'Pending' to 'Completed' in Excel with Aspose | C# Aspose.Cells FindOptions example for replacing values in column D | Replace case‑insensitive text in Excel worksheet range using Aspose.Cells .NET | Update multiple Excel cells programmatically with Aspose.Cells in C#
// Tags: Aspose.Cells find and replace in range | C# bulk cell value update Excel | Excel range D5:D25 value substitution Aspose | case-insensitive text replacement Aspose.Cells | Aspose.Cells FindOptions usage C# | replace specific column cells Aspose.Cells

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads an existing workbook, iterates through cells D5 to D25 on the first worksheet, and replaces any occurrence of the text "Pending" (case‑insensitive) with "Completed" before saving the modified file, demonstrating a targeted find‑and‑replace operation with Aspose.Cells for .NET.
    class Program
    {
        static void Main(string[] args)
        {
            const string inputPath = "input.xlsx";
            const string outputPath = "output.xlsx";

            // Verify that the input file exists to avoid FileNotFoundException
            if (!File.Exists(inputPath))
            {
                Console.WriteLine($"Input file not found: {inputPath}");
                return;
            }

            try
            {
                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (index 0)
                Worksheet worksheet = workbook.Worksheets[0];

                // Iterate through cells D5:D25 (column index 3, rows 4‑24)
                for (int row = 4; row <= 24; row++)
                {
                    Cell cell = worksheet.Cells[row, 3];
                    if (cell.Value != null &&
                        string.Equals(cell.StringValue, "Pending", StringComparison.OrdinalIgnoreCase))
                    {
                        cell.PutValue("Completed");
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to {outputPath}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
