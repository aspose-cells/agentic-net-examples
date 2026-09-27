// Title: Merge two non‑contiguous cell ranges with UnionRanges and loop through the resulting UnionRange using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that defines two separate Range objects (A1:B2 and C3:D3) on a worksheet and merges them using the UnionRanges API. | Create a foreach loop that iterates over the UnionRange returned by UnionRanges and outputs each cell's address and value. | Demonstrate how to persist the workbook after performing the range union, including specifying the output file name.
// Common Searches: Aspose.Cells C# example for unioning non‑contiguous ranges | How to combine A1:B2 and C3:D3 into a single range with UnionRanges | Iterate over cells in a UnionRange object in Aspose.Cells for .NET | Saving workbook after using UnionRanges method in Aspose.Cells | UnionRanges usage for merging separate ranges in C#
// Tags: unionranges method Aspose.Cells | merge noncontiguous ranges C# | enumerate cells in UnionRange | save workbook after range union Aspose.Cells | create multiple Range objects Aspose.Cells

using Aspose.Cells;
using System;

// The sample creates a workbook, defines two separate ranges (A1:B2 and C3:D3), merges them with the UnionRanges method to obtain a UnionRange, iterates through each cell printing its address and value, and saves the file as UnionResult.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);
            sheet.Cells["A2"].PutValue(30);
            sheet.Cells["B2"].PutValue(40);
            sheet.Cells["C3"].PutValue(50);
            sheet.Cells["D3"].PutValue(60);

            // Create two separate Range objects (use fully qualified name to avoid ambiguity with System.Range)
            Aspose.Cells.Range range1 = sheet.Cells.CreateRange("A1:B2"); // 2x2 block
            Aspose.Cells.Range range2 = sheet.Cells.CreateRange("C3:D3"); // single row block

            // Combine the ranges using the UnionRanges method (returns UnionRange)
            UnionRange unionRange = range1.UnionRanges(new Aspose.Cells.Range[] { range2 });

            // Iterate through the cells in the resulting union range
            foreach (Cell cell in unionRange)
            {
                Console.WriteLine($"Cell {cell.Name} = {cell.Value}");
            }

            // Save the workbook (optional)
            string outputPath = "UnionResult.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
