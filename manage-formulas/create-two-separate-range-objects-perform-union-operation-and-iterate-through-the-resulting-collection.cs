// Title: How to union two non‑contiguous ranges and iterate over their cells using Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# union non‑contiguous ranges and read cell values | C# iterate over multiple ranges after UnionRanges in Aspose.Cells | example of using UnionRanges to combine A1:B2 and C3:D4 in Aspose.Cells | how to loop through cells in a unioned range with Aspose.Cells for .NET | Aspose.Cells create range objects and union them programmatically
// Tags: unionranges method Aspose.Cells | create range objects worksheet Aspose.Cells | iterate cells in unioned range C# | combine noncontiguous ranges Aspose.Cells | loop through unioned ranges .NET

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The sample creates a workbook, fills cells A1 through D4, defines two separate Range objects (A1:B2 and C3:D4), merges them with the UnionRanges API, and then iterates over each range and its cells, printing each cell's address and value.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];

            // Populate sample data
            worksheet.Cells["A1"].PutValue(10);
            worksheet.Cells["A2"].PutValue(20);
            worksheet.Cells["B1"].PutValue(30);
            worksheet.Cells["B2"].PutValue(40);
            worksheet.Cells["C3"].PutValue(50);
            worksheet.Cells["C4"].PutValue(60);
            worksheet.Cells["D3"].PutValue(70);
            worksheet.Cells["D4"].PutValue(80);

            // Create two separate Range objects
            AsposeRange range1 = worksheet.Cells.CreateRange("A1:B2");
            AsposeRange range2 = worksheet.Cells.CreateRange("C3:D4");

            // Perform Union operation using the recommended API
            UnionRange unionRange = range1.UnionRanges(new AsposeRange[] { range2 });

            // Iterate through the cells in all unioned ranges
            foreach (AsposeRange r in unionRange)
            {
                foreach (Cell cell in r)
                {
                    Console.WriteLine($"Cell {cell.Name} = {cell.Value}");
                }
            }

            // Optional: Save the workbook if needed
            // workbook.Save("Result.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
