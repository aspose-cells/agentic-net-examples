// Title: Clear multiple non‑adjacent ranges with a UnionRange in Aspose.Cells for .NET (C#)
// Description: Shows how to create three separate ranges (A1:B2, D1:E2, G1:H2), merge them into a UnionRange, and clear their contents using Aspose.Cells for .NET. The workbook is saved to confirm the cells are emptied.
// Keywords: Aspose.Cells UnionRange | C# UnionRange clear contents | non‑adjacent range clear | Aspose.Cells ClearContents | multiple ranges Aspose.Cells | Aspose.Cells .NET example | UnionRange C#
// Common Searches: Aspose.Cells clear non adjacent cells | How to use UnionRange in C# Aspose.Cells | Clear multiple ranges with one call Aspose.Cells | UnionRange ClearContents example | Combine separate ranges Aspose.Cells
// Developer Intent: Combine several distinct cell blocks into a UnionRange and remove their values in a single workflow.
// Use Cases: Refresh scattered input sections of a report before inserting new data. | Delete intermediate calculation results stored in different worksheet areas. | Batch clear placeholder or validation text across non‑contiguous template ranges.
// AI Prompts: Write C# code that builds a UnionRange from an array of address strings and clears both contents and formatting in one step. | Show how to clear all cells in a UnionRange using a single method call instead of iterating each Range. | Explain how to reuse a UnionRange object to clear cells on multiple worksheets within the same workbook.

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsUnionRangeClearDemo
{
    // Shows how to create three separate ranges (A1:B2, D1:E2, G1:H2), merge them into a UnionRange, and clear their contents using Aspose.Cells for .NET. The workbook is saved to confirm the cells are emptied.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Populate three separate blocks with sample data
                // Block 1: A1:B2
                cells["A1"].PutValue("Block1_R1C1");
                cells["B1"].PutValue("Block1_R1C2");
                cells["A2"].PutValue("Block1_R2C1");
                cells["B2"].PutValue("Block1_R2C2");

                // Block 2: D1:E2
                cells["D1"].PutValue("Block2_R1C1");
                cells["E1"].PutValue("Block2_R1C2");
                cells["D2"].PutValue("Block2_R2C1");
                cells["E2"].PutValue("Block2_R2C2");

                // Block 3: G1:H2
                cells["G1"].PutValue("Block3_R1C1");
                cells["H1"].PutValue("Block3_R1C2");
                cells["G2"].PutValue("Block3_R2C1");
                cells["H2"].PutValue("Block3_R2C2");

                // Create Range objects for each block
                AsposeRange range1 = cells.CreateRange("A1:B2");
                AsposeRange range2 = cells.CreateRange("D1:E2");
                AsposeRange range3 = cells.CreateRange("G1:H2");

                // Build a UnionRange from the first range
                UnionRange unionRange = cells.CreateRange(range1.RefersTo).UnionRanges(new AsposeRange[] { range1 });

                // Add the second and third ranges using the Union(string) overload
                unionRange = unionRange.Union("D1:E2");
                unionRange = unionRange.Union("G1:H2");

                // Clear contents of all ranges within the UnionRange
                foreach (AsposeRange r in unionRange.Ranges)
                {
                    r.ClearContents();
                }

                // Save the workbook to verify that the three blocks are cleared
                workbook.Save("UnionRangeClearContentsDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
