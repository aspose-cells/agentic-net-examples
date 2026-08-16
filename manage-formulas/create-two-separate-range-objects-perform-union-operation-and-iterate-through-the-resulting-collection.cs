// Title: Combine Two Ranges with UnionRange and Iterate Cells – Aspose.Cells for .NET (C#)
// Description: This example shows how to create a workbook, define two separate ranges (A1:B2 and C1:D2), merge them into a UnionRange using both the UnionRanges and Union(string) overloads, loop through every cell in the combined range, and save the result as an Excel file.
// Keywords: Aspose.Cells UnionRange C# | combine multiple ranges .NET | iterate cells in UnionRange | create range A1:B2 Aspose.Cells | non‑contiguous range processing | Aspose.Cells workbook save | bulk formatting with UnionRange
// Common Searches: union two ranges Aspose.Cells C# | how to iterate UnionRange cells | create UnionRange from existing ranges .NET | Aspose.Cells combine ranges and save workbook | C# example UnionRange iteration
// Developer Intent: The developer needs to merge two distinct cell blocks into a single UnionRange and process each cell in the merged collection.
// Use Cases: Apply a single style or formula to several non‑adjacent blocks of data. | Export values from multiple separate areas of a sheet in one pass. | Generate a report that consolidates scattered tables before saving the workbook.
// AI Prompts: Write C# code that creates three ranges, merges them into a UnionRange, and sets a yellow background for all cells in the union. | Show how to add a user‑specified range to an existing UnionRange with the Union(string) method, then iterate and log each cell value. | Explain how to obtain the total number of cells in a UnionRange and write that count into cell Z1.

using System;
using Aspose.Cells;
using ARange = Aspose.Cells.Range;

namespace AsposeCellsUnionExample
{
    // This example shows how to create a workbook, define two separate ranges (A1:B2 and C1:D2), merge them into a UnionRange using both the UnionRanges and Union(string) overloads, loop through every cell in the combined range, and save the result as an Excel file.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Create the first range (A1:B2) and put some sample data
                ARange range1 = worksheet.Cells.CreateRange("A1:B2");
                range1[0, 0].PutValue("R1C1");
                range1[0, 1].PutValue("R1C2");
                range1[1, 0].PutValue("R2C1");
                range1[1, 1].PutValue("R2C2");

                // Create the second range (C1:D2) and put some sample data
                ARange range2 = worksheet.Cells.CreateRange("C1:D2");
                range2[0, 0].PutValue("R1C3");
                range2[0, 1].PutValue("R1C4");
                range2[1, 0].PutValue("R2C3");
                range2[1, 1].PutValue("R2C4");

                // Build a UnionRange from the first range
                UnionRange unionRange = worksheet.Cells
                    .CreateRange("A1:B2")               // base range for the UnionRange
                    .UnionRanges(new ARange[] { range1 }); // add the first range

                // Add the second range to the union using the Union(string) overload
                unionRange = unionRange.Union("C1:D2");

                // Iterate through all cells in the resulting UnionRange
                Console.WriteLine("Iterating through cells in the UnionRange:");
                foreach (Cell cell in unionRange)
                {
                    Console.WriteLine($"{cell.Name}: {cell.Value}");
                }

                // Save the workbook (optional, just to persist data)
                workbook.Save("UnionRangeDemo.xlsx");
                Console.WriteLine("Workbook saved as UnionRangeDemo.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
