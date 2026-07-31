// Title: Aspose.Cells for .NET – Union two Range objects and iterate over the UnionRange
// Description: The sample creates a workbook, defines two separate ranges (A1:B2 and C3:D4), fills them with data, merges them into a UnionRange using the UnionRanges/Union methods, iterates through each constituent range and all cells in the combined range, and saves the workbook.
// Keywords: Aspose.Cells | UnionRange | C# .NET | combine ranges | iterate cells | CreateRange | non‑contiguous ranges | Excel automation | range union example | Aspose.Cells API
// Common Searches: how to union two ranges Aspose.Cells .NET | iterate over UnionRange cells C# | Aspose.Cells combine non adjacent ranges | UnionRanges method example | save workbook after unioning ranges Aspose.Cells
// Developer Intent: The developer wants to merge multiple non‑adjacent ranges into a single UnionRange and loop through the resulting ranges and cells programmatically.
// Use Cases: Apply a single formatting style to several scattered blocks by uniting them into a UnionRange. | Calculate totals or aggregates across disjoint worksheet sections by iterating a UnionRange. | Export data from multiple areas as one logical range for charting or reporting.
// AI Prompts: Generate C# code that creates three non‑adjacent ranges, unions them with UnionRange, and writes the sum of each range to a new cell. | Provide an Aspose.Cells .NET snippet that unions ranges and sets a background color for all cells in the resulting UnionRange. | Explain how to retrieve the address of each range inside a UnionRange and use those addresses to create named ranges.

using System;
using Aspose.Cells;

namespace AsposeCellsUnionExample
{
    // Alias to avoid conflict with System.Range introduced in C# 8.0
    using AsposeRange = Aspose.Cells.Range;

    // The sample creates a workbook, defines two separate ranges (A1:B2 and C3:D4), fills them with data, merges them into a UnionRange using the UnionRanges/Union methods, iterates through each constituent range and all cells in the combined range, and saves the workbook.
    public class Program
    {
        public static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];

                // Define two separate ranges
                AsposeRange range1 = worksheet.Cells.CreateRange("A1:B2");
                AsposeRange range2 = worksheet.Cells.CreateRange("C3:D4");

                // Populate the ranges with sample data (optional, just for visualization)
                range1[0, 0].PutValue("R1C1");
                range1[0, 1].PutValue("R1C2");
                range1[1, 0].PutValue("R2C1");
                range1[1, 1].PutValue("R2C2");

                range2[0, 0].PutValue("R3C3");
                range2[0, 1].PutValue("R3C4");
                range2[1, 0].PutValue("R4C3");
                range2[1, 1].PutValue("R4C4");

                // Perform Union operation using the UnionRange API
                UnionRange unionRange = worksheet.Cells.CreateRange(range1.RefersTo).UnionRanges(new AsposeRange[] { range1 });
                unionRange = unionRange.Union(new AsposeRange[] { range2 });

                // Iterate through the collection of ranges that constitute the UnionRange
                Console.WriteLine("Iterating through the ranges in the UnionRange:");
                for (int i = 0; i < unionRange.RangeCount; i++)
                {
                    AsposeRange r = unionRange.Ranges[i];
                    Console.WriteLine($"Range {i + 1}: {r.Address}");
                }

                // Optionally, iterate through all cells in the UnionRange
                Console.WriteLine("\nIterating through all cells in the UnionRange:");
                foreach (Cell cell in unionRange)
                {
                    Console.WriteLine($"{cell.Name}: {cell.Value}");
                }

                // Save the workbook (optional, demonstrates that data persists)
                string outputPath = "UnionRangeDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"\nWorkbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
