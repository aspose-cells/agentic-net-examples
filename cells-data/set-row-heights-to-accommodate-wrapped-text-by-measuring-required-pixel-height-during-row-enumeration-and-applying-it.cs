// Title: Auto‑fit Excel row heights for wrapped text using Aspose.Cells for .NET
// Common Searches: C# Aspose.Cells auto fit row height when cells have wrap text | How to adjust Excel row height based on wrapped content with Aspose.Cells | Programmatically set row height for wrapped cells in a .NET workbook | Measure required row pixel height for wrapped text using Aspose.Cells | Iterate rows and auto‑fit height for wrapped text in Aspose.Cells C#
// Tags: auto-fit rows Aspose.Cells C# | detect wrapped text cell style | row height calculation Aspose.Cells | worksheet.AutoFitRow usage | excel row height optimization .NET

using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsExample
{
    // The example loads a workbook, scans each used row for cells with the Wrap Text style, and calls Worksheet.AutoFitRow to automatically set the row height so wrapped text is fully visible, then saves the updated file.
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                string inputPath = "input.xlsx";
                string outputPath = "output.xlsx";

                // Ensure the input file exists to avoid FileNotFoundException
                if (!File.Exists(inputPath))
                {
                    Console.WriteLine($"Input file not found: {inputPath}");
                    return;
                }

                // Load the existing workbook
                Workbook workbook = new Workbook(inputPath);

                // Access the first worksheet (adjust index if needed)
                Worksheet worksheet = workbook.Worksheets[0];

                // Determine the used range
                int maxRow = worksheet.Cells.MaxDataRow;
                int maxCol = worksheet.Cells.MaxDataColumn;

                // Iterate through each row in the used range
                for (int rowIndex = 0; rowIndex <= maxRow; rowIndex++)
                {
                    bool rowHasWrappedText = false;

                    // Check each cell in the current row for the Wrap Text flag
                    for (int colIndex = 0; colIndex <= maxCol; colIndex++)
                    {
                        Cell cell = worksheet.Cells[rowIndex, colIndex];
                        Style style = cell.GetStyle();

                        if (style.IsTextWrapped)
                        {
                            rowHasWrappedText = true;
                            break; // No need to check further cells in this row
                        }
                    }

                    // AutoFit the row height if any cell has wrapped text
                    if (rowHasWrappedText)
                    {
                        // AutoFitRow with a single argument adjusts the row height.
                        worksheet.AutoFitRow(rowIndex);
                    }
                }

                // Save the modified workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
