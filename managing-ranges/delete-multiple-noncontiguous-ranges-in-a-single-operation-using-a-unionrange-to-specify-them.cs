// Title: Delete rows for multiple non‑contiguous ranges in a worksheet using Aspose.Cells for .NET
// AI Prompts: Generate C# code that removes all rows intersecting the ranges A1:B2 and D4:E5 in a worksheet with Aspose.Cells, ensuring row indices remain correct. | Rewrite the example so the two separate ranges are merged into a single operation for row deletion using Aspose.Cells.
// Common Searches: C# Aspose.Cells how to delete rows from multiple separate ranges | Aspose.Cells remove non‑adjacent cell blocks without manual index handling | Delete rows for A1:B2 and D4:E5 in one call using Aspose.Cells .NET | Aspose.Cells example of deleting rows defined by several ranges
// Tags: Aspose.Cells UnionRange row deletion | Aspose.Cells delete noncontiguous ranges .NET | Aspose.Cells prevent index shift when deleting rows | Aspose.Cells multiple range removal example | Aspose.Cells workbook save after row deletion

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The program creates a new workbook, populates two separate blocks (A1:B2 and D4:E5), builds Range objects for each block, collects all row indices covered by both ranges, deletes those rows in descending order to avoid index shifting, and saves the workbook as DeletedRanges.xlsx.
class DeleteNonContiguousRanges
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data in the ranges that will be deleted
            cells["A1"].PutValue("Data A1");
            cells["A2"].PutValue("Data A2");
            cells["B1"].PutValue("Data B1");
            cells["B2"].PutValue("Data B2");
            cells["D4"].PutValue("Data D4");
            cells["D5"].PutValue("Data D5");
            cells["E4"].PutValue("Data E4");
            cells["E5"].PutValue("Data E5");

            // Create two separate ranges
            AsposeRange range1 = cells.CreateRange("A1:B2");
            AsposeRange range2 = cells.CreateRange("D4:E5");

            // Collect all row indices that intersect the two ranges
            var rowsToDelete = new List<int>();

            for (int r = range1.FirstRow; r < range1.FirstRow + range1.RowCount; r++)
                rowsToDelete.Add(r);

            for (int r = range2.FirstRow; r < range2.FirstRow + range2.RowCount; r++)
                rowsToDelete.Add(r);

            // Delete rows from bottom to top to avoid index shifting
            foreach (int rowIndex in rowsToDelete.Distinct().OrderByDescending(i => i))
                cells.DeleteRow(rowIndex);

            // Save the workbook
            string outputPath = "DeletedRanges.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
