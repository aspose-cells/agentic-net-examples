// Title: C# Example: Unmerge cells D4‑F4 and restore original boundaries using Aspose.Cells
// Description: Shows how to load or create a workbook, select the merged range D4:F4, call Range.UnMerge() to split the cells back to their original size, optionally write distinct values, and save the workbook with Aspose.Cells for .NET.
// Keywords: Aspose.Cells unmerge cells C# | Range.UnMerge example | split merged range D4:F4 | restore original cell boundaries Aspose | C# Aspose.Cells unmerge range
// Common Searches: C# Aspose.Cells unmerge specific range | how to split merged cells D4 to F4 | restore original cells after merging Aspose.Cells | unmerge merged header row Aspose.Cells .NET
// Developer Intent: Remove the merge on range D4:F4 so each cell (D4, E4, F4) becomes independent again.
// Use Cases: Revert a temporary merged header before applying column‑level formatting. | Split merged cells after importing data to enable individual cell edits. | Restore cell boundaries when a report layout requires separate cell values.
// AI Prompts: Generate C# code that uses Aspose.Cells to unmerge the range D4:F4 and writes a unique value to each cell. | Explain how to programmatically unmerge a range and then apply different formatting to D4, E4, and F4 with Aspose.Cells. | Show how to check if a range is merged before calling UnMerge in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

// Shows how to load or create a workbook, select the merged range D4:F4, call Range.UnMerge() to split the cells back to their original size, optionally write distinct values, and save the workbook with Aspose.Cells for .NET.
class UnmergeDemo
{
    static void Main()
    {
        try
        {
            // Create a new workbook (or load an existing one)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];

            // Create the range that was previously merged (D4:F4)
            // Use fully qualified type to avoid conflict with System.Range
            Aspose.Cells.Range mergedRange = worksheet.Cells.CreateRange("D4", "F4");

            // Unmerge the range to restore original cell boundaries
            mergedRange.UnMerge();

            // Optional: add values to the now separate cells for verification
            worksheet.Cells["D4"].PutValue("D4");
            worksheet.Cells["E4"].PutValue("E4");
            worksheet.Cells["F4"].PutValue("F4");

            // Save the workbook
            workbook.Save("UnmergedDemo.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"An error occurred: {ex.Message}");
        }
    }
}
