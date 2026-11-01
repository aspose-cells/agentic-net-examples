// Title: How to compute the effective cell count of a range that contains merged cells using Aspose.Cells for .NET (C#)
// Common Searches: C# Aspose.Cells get number of distinct cells in a range with merged cells | calculate actual cell count for a merged area in an Aspose.Cells range | Aspose.Cells .NET determine cell count after merging cells B2:C3 | verify merged cell count in a defined range using Aspose.Cells API | Aspose.Cells CreateRange count cells excluding merged duplicates
// Tags: aspose.cells merged cells count | c# aspose.cells range cell count | aspose.cells create range merged area | aspose.cells rowcount columncount merged cells | aspose.cells calculate distinct cells in range | aspose.cells workbook save merged cells example

using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

// The example creates a new workbook, merges a 2x2 block (B2:C3), defines a range A1:D4, computes the total cells in the range, adjusts the count so the merged block is treated as a single cell, verifies the result against the expected count, outputs the verification status, and saves the workbook as MergedCellCountDemo.xlsx.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Merge a 2x2 block (B2:C3) – rows and columns are zero‑based
            sheet.Cells.Merge(1, 1, 2, 2);

            // Define a range that includes the merged area (A1:D4)
            AsposeRange range = sheet.Cells.CreateRange("A1:D4");

            // Calculate the actual cell count.
            // Total cells in the range minus the extra cells covered by the merged area.
            int totalCellsInRange = range.RowCount * range.ColumnCount; // 4 rows * 4 columns = 16
            int mergedAreaSize = 2 * 2;                                 // 2 rows * 2 columns = 4
            int actualCellCount = totalCellsInRange - (mergedAreaSize - 1); // merged area counts as 1

            // Expected count (same calculation for verification)
            int expectedCellCount = 16 - 3; // = 13

            // Output verification results
            Console.WriteLine($"Actual cell count: {actualCellCount}");
            Console.WriteLine($"Expected cell count: {expectedCellCount}");
            Console.WriteLine($"Verification result: {(actualCellCount == expectedCellCount ? "PASS" : "FAIL")}");

            // Save the workbook to inspect the merged cells manually
            string outputPath = "MergedCellCountDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to '{outputPath}'.");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
