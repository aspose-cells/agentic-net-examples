// Title: Aspose.Cells for .NET: Create a UnionRange of non‑contiguous blocks and clear their contents (C#)
// Description: This C# example shows how to build a Workbook, define three separate ranges (A1:B2, D4:E5, G7:H8), merge them with the UnionRanges method into a UnionRange, clear the data in all blocks, and save the file. The code demonstrates the most concise way to reset multiple non‑adjacent areas in a worksheet.
// Keywords: Aspose.Cells | UnionRange | C# | .NET | clear contents | non‑contiguous ranges | range union | Excel automation | sample code | GitHub example
// Common Searches: Aspose.Cells clear multiple ranges C# | UnionRange example for .NET | how to clear non‑adjacent cells with Aspose.Cells | combine A1:B2 D4:E5 G7:H8 into one range | Aspose.Cells UnionRanges method usage
// Developer Intent: Combine several distinct cell blocks into a UnionRange and remove all values from those blocks in a single operation.
// Use Cases: Reset temporary calculation zones before generating a final report. | Erase user‑entered data from specific template sections while keeping formatting intact. | Delete test data from multiple report areas after automated validation.
// AI Prompts: Write C# code that creates a UnionRange from three given ranges and clears their contents using Aspose.Cells. | Suggest a method to clear all cells in a UnionRange without looping through each range individually. | Explain how the UnionRanges API works and how to apply it to manipulate non‑contiguous blocks in a worksheet.

using System;
using Aspose.Cells;

namespace AsposeCellsUnionRangeClearDemo
{
    // This C# example shows how to build a Workbook, define three separate ranges (A1:B2, D4:E5, G7:H8), merge them with the UnionRanges method into a UnionRange, clear the data in all blocks, and save the file. The code demonstrates the most concise way to reset multiple non‑adjacent areas in a worksheet.
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

                // Define three separate ranges
                Aspose.Cells.Range range1 = cells.CreateRange("A1:B2");
                Aspose.Cells.Range range2 = cells.CreateRange("D4:E5");
                Aspose.Cells.Range range3 = cells.CreateRange("G7:H8");

                // Populate the ranges with sample data (optional, just for demonstration)
                range1[0, 0].PutValue("R1C1");
                range1[0, 1].PutValue("R1C2");
                range1[1, 0].PutValue("R2C1");
                range1[1, 1].PutValue("R2C2");

                range2[0, 0].PutValue("R1C3");
                range2[0, 1].PutValue("R1C4");
                range2[1, 0].PutValue("R2C3");
                range2[1, 1].PutValue("R2C4");

                range3[0, 0].PutValue("R1C5");
                range3[0, 1].PutValue("R1C6");
                range3[1, 0].PutValue("R2C5");
                range3[1, 1].PutValue("R2C6");

                // Create a UnionRange that combines the three blocks
                UnionRange unionRange = range1.UnionRanges(new Aspose.Cells.Range[] { range2, range3 });

                // Clear contents of all blocks in the union range
                foreach (Aspose.Cells.Range r in unionRange.Ranges)
                {
                    r.ClearContents();
                }

                // Save the workbook
                string outputPath = "UnionRangeClearContentsDemo.xlsx";
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
