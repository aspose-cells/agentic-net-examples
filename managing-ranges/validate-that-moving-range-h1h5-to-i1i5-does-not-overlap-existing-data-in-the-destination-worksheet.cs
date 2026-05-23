using System;
using System.IO;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeMoveValidation
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate source range H1:H5 with sample data
                for (int i = 0; i < 5; i++)
                {
                    cells[i, 7].PutValue($"Source {i + 1}"); // Column 7 = H
                }

                // (Optional) Populate destination range I1:I5 with data to test overlap detection
                // Uncomment the following line to simulate existing data in the destination
                // cells[2, 8].PutValue("Existing Data"); // Row 2 (zero‑based) = I3

                // Define source and destination ranges
                AsposeRange sourceRange = cells.CreateRange("H1:H5");
                AsposeRange destRange = cells.CreateRange("I1:I5");

                // Check geometric intersection (should be false for non‑overlapping columns)
                bool isIntersect = sourceRange.IsIntersect(destRange);

                // Determine if any cell in the destination range already contains data
                bool destHasData = false;
                int destLastRow = destRange.FirstRow + destRange.RowCount - 1;
                for (int row = destRange.FirstRow; row <= destLastRow; row++)
                {
                    // A cell is considered non‑empty if its Value is not null
                    if (cells[row, destRange.FirstColumn].Value != null)
                    {
                        destHasData = true;
                        break;
                    }
                }

                // Validate before moving
                if (!isIntersect && !destHasData)
                {
                    // Move the source range to the destination start cell (I1)
                    sourceRange.MoveTo(destRange.FirstRow, destRange.FirstColumn);
                    Console.WriteLine("Range moved successfully.");
                }
                else
                {
                    Console.WriteLine("Cannot move the range: it either intersects or the destination contains data.");
                }

                // Prepare output path
                string outputPath = "RangeMoveValidationOutput.xlsx";
                string outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));

                // Ensure the output directory exists (handle possible null)
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                // Save the workbook
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}