// Title: C# – Unmerge Cells D4:F4 with Aspose.Cells and Restore Original Boundaries
// Description: Demonstrates how to create or load a workbook, define the merged range D4:F4, call UnMerge to split the cells back to their original boundaries, verify the merge state with IsMerged, and save the file as UnmergedRangeDemo.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells unmerge range C# | unmerge merged cells Aspose.Cells | D4 F4 unmerge Aspose.Cells | restore cell boundaries .NET | UnMerge method Aspose.Cells
// Common Searches: how to unmerge a range in Aspose.Cells C# | unmerge D4:F4 cells Aspose.Cells | restore original cell layout after merging Aspose.Cells | Aspose.Cells UnMerge example | check IsMerged after unmerge Aspose.Cells
// Developer Intent: Split the previously merged range D4:F4 into individual cells and confirm the cells are no longer merged.
// Use Cases: Remove a merged header before applying column‑specific formatting. | Prepare imported data for per‑cell calculations after a bulk merge. | Validate that a range is unmerged before iterating row‑wise.
// AI Prompts: Generate C# code using Aspose.Cells to unmerge the range D4:F4 and save the workbook. | Show how to use the IsMerged property to confirm a cell is no longer merged after UnMerge. | Explain step‑by‑step how to programmatically unmerge a cell range and verify its state with Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to create or load a workbook, define the merged range D4:F4, call UnMerge to split the cells back to their original boundaries, verify the merge state with IsMerged, and save the file as UnmergedRangeDemo.xlsx using Aspose.Cells for .NET.
    public class UnmergeRangeDemo
    {
        public static void Main(string[] args)
        {
            try
            {
                Run();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        public static void Run()
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create a Range object that represents the previously merged cells D4:F4
            Aspose.Cells.Range mergedRange = worksheet.Cells.CreateRange("D4", "F4");

            // Unmerge the range, restoring the original individual cells
            mergedRange.UnMerge();

            // Verify that the cells are no longer merged
            Console.WriteLine("Is D4 merged after UnMerge? " + worksheet.Cells["D4"].IsMerged);

            // Save the workbook
            workbook.Save("UnmergedRangeDemo.xlsx");
            Console.WriteLine("Workbook saved as UnmergedRangeDemo.xlsx");
        }
    }
}
