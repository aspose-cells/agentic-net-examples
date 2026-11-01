// Title: How to merge a range containing hidden rows while preserving hidden state with Aspose.Cells for .NET (C#)
// Common Searches: Aspose.Cells C# merge range that includes hidden rows without unhiding them | keep hidden rows after merging cells using Aspose.Cells .NET | merge A1:C3 in Aspose.Cells while preserving row IsHidden property | C# Aspose.Cells merge cells across hidden rows example | retain hidden row state when merging cells in a workbook with Aspose.Cells
// Tags: merge range with hidden rows Aspose.Cells | retain row hidden state after merge C# | Aspose.Cells create and merge cell range | preserve IsHidden property during cell merge | C# workbook merge cells A1:C3

using System;
using System.IO;
using Aspose.Cells;

// The example creates a workbook, hides rows 2 and 3, merges the range A1:C3, confirms the rows stay hidden, and saves the file as MergedWithHiddenRows.xlsx.
class MergeHiddenRowsExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some data for visual reference (optional)
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue("Row 2");
            sheet.Cells["A3"].PutValue("Row 3");

            // Hide rows 2 and 3 (zero‑based index: rows 1 and 2)
            sheet.Cells.Rows[1].IsHidden = true;
            sheet.Cells.Rows[2].IsHidden = true;

            // Merge a range that includes the hidden rows (A1:C3)
            // This will merge cells A1, B1, C1, A2, B2, C2, A3, B3, C3
            sheet.Cells.CreateRange("A1:C3").Merge();

            // Verify that the hidden rows remain hidden after merging
            bool row2Hidden = sheet.Cells.Rows[1].IsHidden;
            bool row3Hidden = sheet.Cells.Rows[2].IsHidden;

            Console.WriteLine("Row 2 hidden after merge: " + row2Hidden);
            Console.WriteLine("Row 3 hidden after merge: " + row3Hidden);

            // Save the workbook to verify the result manually if needed
            string outputPath = "MergedWithHiddenRows.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
