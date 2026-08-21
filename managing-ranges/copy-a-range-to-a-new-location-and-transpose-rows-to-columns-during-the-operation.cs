// Title: Copy a Range and Transpose Rows to Columns with Aspose.Cells for .NET (C#)
// Description: Demonstrates how to create a workbook, fill a horizontal range (A1:D1), define source and destination ranges, enable PasteOptions.Transpose, copy the data vertically, and save the result as CopyTransposeDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells copy range transpose C# | PasteOptions.Transpose Aspose.Cells | copy range to new location Aspose.Cells | transpose rows to columns .NET | Aspose.Cells range copy example
// Common Searches: Aspose.Cells copy range with transpose | C# Aspose.Cells transpose rows to columns | how to use PasteOptions.Transpose in Aspose.Cells | copy horizontal range to vertical range Aspose.Cells | Aspose.Cells range copy and paste example
// Developer Intent: Copy a source range to a different location while converting rows into columns in a .NET application.
// Use Cases: Turn a header row into a vertical list for data validation or dropdowns. | Reformat a matrix so that reporting templates expecting columnar data receive the correct layout. | Convert horizontally collected sensor readings into a column format for charting or analysis.
// AI Prompts: Generate C# code that copies a range and transposes it with Aspose.Cells, keeping formatting and formulas intact. | Show an example of copying a vertical range and pasting it horizontally using PasteOptions.Transpose in Aspose.Cells. | Explain how to copy multiple non‑contiguous ranges and apply transposition to each with Aspose.Cells for .NET.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    // Demonstrates how to create a workbook, fill a horizontal range (A1:D1), define source and destination ranges, enable PasteOptions.Transpose, copy the data vertically, and save the result as CopyTransposeDemo.xlsx using Aspose.Cells for .NET.
    public class CopyAndTransposeDemo
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Populate a horizontal source range (A1:D1) with sample data
                for (int col = 0; col < 4; col++)
                {
                    cells[0, col].PutValue($"Data {col + 1}");
                }

                // Define the source range (row 0, column 0, 1 row, 4 columns)
                AsposeRange sourceRange = cells.CreateRange(0, 0, 1, 4);

                // Define the destination range where the transposed data will be placed
                // It should be 4 rows by 1 column starting at A2 (row 1, column 0)
                AsposeRange destRange = cells.CreateRange(1, 0, 4, 1);

                // Set up paste options to enable transposition
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All, // copy all content (values, formats, etc.)
                    Transpose = true           // transpose rows ↔ columns during paste
                };

                // Perform the copy with transpose
                destRange.Copy(sourceRange, pasteOptions);

                // Save the workbook to a file
                string outputPath = "CopyTransposeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }

    // Entry point for the application
    public class Program
    {
        public static void Main(string[] args)
        {
            CopyAndTransposeDemo.Run();
        }
    }
}
