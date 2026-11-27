// Title: Create two Aspose.Cells Range objects, union them, and iterate over each cell in C#
// Common Searches: Aspose.Cells C# combine two ranges with Union and loop through cells | example of iterating cells from multiple Aspose.Cells ranges in .NET | C# Aspose.Cells create range A1:B2 and C1:C2 then enumerate union | using Aspose.Cells to union non‑contiguous ranges and read cell values | Aspose.Cells range collection enumeration after union operation
// Tags: Aspose.Cells range union enumeration | C# Aspose.Cells combine multiple ranges | Aspose.Cells iterate cells in unioned ranges | Aspose.Cells create specific range A1:B2 | Aspose.Cells workbook save after range union

using Aspose.Cells;
using System;
using System.Collections;
using System.IO;

// Alias to avoid conflict with System.Range introduced in C# 8.0
using AsposeRange = Aspose.Cells.Range;

// The sample program creates a workbook, fills cells A1 through C2, defines two ranges (A1:B2 and C1:C2) with Aspose.Cells, adds them to a collection to form a union, iterates each cell in the combined ranges to display its address and value, and saves the workbook as UnionRangeExample.xlsx.
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
            sheet.Cells["A1"].PutValue(1);
            sheet.Cells["A2"].PutValue(2);
            sheet.Cells["B1"].PutValue(3);
            sheet.Cells["B2"].PutValue(4);
            sheet.Cells["C1"].PutValue(5);
            sheet.Cells["C2"].PutValue(6);

            // Create the first range (A1:B2)
            AsposeRange range1 = sheet.Cells.CreateRange("A1:B2");

            // Create the second range (C1:C2)
            AsposeRange range2 = sheet.Cells.CreateRange("C1:C2");

            // Combine both ranges into a collection (union)
            ArrayList unionRanges = new ArrayList { range1, range2 };

            // Iterate through each cell in the combined ranges
            foreach (AsposeRange rng in unionRanges)
            {
                foreach (Cell cell in rng)
                {
                    Console.WriteLine($"Cell {cell.Name} = {cell.Value}");
                }
            }

            // Save the workbook (optional)
            string outputPath = "UnionRangeExample.xlsx";
            try
            {
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception saveEx)
            {
                Console.Error.WriteLine($"Failed to save workbook: {saveEx.Message}");
            }
        }
        catch (Exception ex)
        {
            // Log any unexpected errors
            Console.Error.WriteLine($"Error: {ex.Message}");
        }
    }
}
