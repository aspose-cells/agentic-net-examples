// Title: C# – Unmerge cells in row 10 of an Excel workbook with Aspose.Cells while preserving values
// Description: Load an existing .xlsx file, detect every merged range that touches row 10, unmerge those ranges, and keep the original top‑left cell content using Aspose.Cells for .NET.
// Keywords: Aspose.Cells unmerge row 10 | C# detect merged cells Excel | preserve data when unmerging | unmerge specific row Aspose | Excel merged area handling C#
// Common Searches: how to unmerge cells in a specific row using Aspose.Cells | C# code to remove merged cells from row 10 in Excel | keep cell value after unmerging with Aspose.Cells | detect merged ranges in a worksheet C# | split merged header rows in Excel programmatically
// Developer Intent: Remove any merged cells that span row 10 in an existing workbook while retaining the original cell value.
// Use Cases: Cleaning imported reports where header cells are merged across row 10 before data processing. | Preparing a sheet for row‑based formulas or data validation that requires each column to have an individual cell in row 10. | Automating template normalization so that row 10 can be split into separate cells without losing the header text.
// AI Prompts: Write C# code with Aspose.Cells that unmerges all ranges intersecting row 15 and keeps the top‑left value. | Show a snippet that lists merged areas in a worksheet and unmerges only those covering a given row index. | Explain how Aspose.Cells' UnMerge method preserves cell data when breaking merged cells.

using System;
using Aspose.Cells;

// Load an existing .xlsx file, detect every merged range that touches row 10, unmerge those ranges, and keep the original top‑left cell content using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Retrieve all merged areas in the worksheet
        CellArea[] mergedAreas = worksheet.Cells.GetMergedAreas();

        // Row index for row 10 (zero‑based indexing)
        int targetRow = 9;

        // Iterate through each merged area and unmerge those that intersect row 10
        foreach (CellArea area in mergedAreas)
        {
            // Does the merged area include the target row?
            if (targetRow >= area.StartRow && targetRow <= area.EndRow)
            {
                int totalRows = area.EndRow - area.StartRow + 1;
                int totalCols = area.EndColumn - area.StartColumn + 1;

                // Unmerge the range; the value in the top‑left cell is preserved automatically
                worksheet.Cells.UnMerge(area.StartRow, area.StartColumn, totalRows, totalCols);
            }
        }

        // Save the modified workbook
        workbook.Save("output.xlsx");
    }
}
