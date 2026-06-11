using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsExamples
{
    public class CountNonEmptyCellsInRange
    {
        public static void Run()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate some sample data in the worksheet
                cells["A1"].PutValue("Header");
                cells["B2"].PutValue(123);
                cells["C3"].PutValue("Text");
                cells["D4"].PutValue("");          // Empty string (treated as empty)
                cells["E5"].PutValue(null);        // Explicit null (treated as empty)

                // Define the range to examine (e.g., B2:D5)
                string rangeAddress = "B2:D5";
                AsposeRange range = cells.CreateRange(rangeAddress);

                // Count non‑empty cells in the specified range
                int nonEmptyCount = 0;
                foreach (Cell cell in range)
                {
                    // A cell is considered non‑empty if its Value is not null
                    // and its string representation is not an empty string.
                    if (cell.Value != null && !string.IsNullOrEmpty(cell.Value.ToString()))
                    {
                        nonEmptyCount++;
                    }
                }

                Console.WriteLine($"Total non‑empty cells in range {rangeAddress}: {nonEmptyCount}");

                // Save the workbook (optional, demonstrates the required save rule)
                string outputPath = "CountNonEmptyCellsInRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to: {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Entry point for testing
        public static void Main()
        {
            Run();
        }
    }
}