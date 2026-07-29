// Title: Copy Adjacent Columns with Merged Cells Preserved Using Aspose.Cells for .NET
// Description: Shows how to duplicate columns A‑C to columns E‑G in a workbook while keeping merged regions (e.g., header A1:C1) intact. The code captures existing merged areas, applies CopyColumns with PasteOptions.All, shifts the merge coordinates, and re‑creates the merged blocks in the target columns before saving.
// Keywords: Aspose.Cells | CopyColumns | merged cells | preserve merges | duplicate columns | C# | .NET | PasteOptions.All | Excel automation | worksheet column copy | US developers | European developers
// Common Searches: Aspose.Cells copy columns keep merged cells | How to duplicate adjacent columns in .NET Excel library | Preserve merged headers when copying columns Aspose | CopyColumns with merged areas example C# | Excel column copy without losing merges using Aspose
// Developer Intent: Copy a set of neighboring columns to another location in the same worksheet while ensuring any merged cells remain unchanged.
// Use Cases: Replicate a multi‑column header for a new table section without breaking the merge. | Clone a formatted data block, including merged titles, for a printable report. | Create a side‑by‑side comparison view by copying columns and preserving their layout.
// AI Prompts: Write C# code that copies columns A‑C to E‑G with Aspose.Cells and automatically re‑applies any merged cells from the source range. | Explain the steps to capture merged areas before a column copy and restore them after using CopyColumns in Aspose.Cells. | Provide a concise tutorial for duplicating adjacent columns while keeping merged headers intact, including PasteOptions configuration.

using System;
using Aspose.Cells;

namespace AsposeCellsColumnCopyWithMergedCells
{
    // Shows how to duplicate columns A‑C to columns E‑G in a workbook while keeping merged regions (e.g., header A1:C1) intact. The code captures existing merged areas, applies CopyColumns with PasteOptions.All, shifts the merge coordinates, and re‑creates the merged blocks in the target columns before saving.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // ------------------------------------------------------------
                // Sample data: create some merged cells spanning columns A-C
                // ------------------------------------------------------------
                // Fill values in the source columns (A, B, C)
                cells["A1"].PutValue("Header");
                cells["A2"].PutValue(10);
                cells["B2"].PutValue(20);
                cells["C2"].PutValue(30);
                cells["A3"].PutValue(40);
                cells["B3"].PutValue(50);
                cells["C3"].PutValue(60);

                // Merge cells A1:C1 (first row header)
                cells.Merge(0, 0, 1, 3); // rows are zero‑based, columns are zero‑based

                // ------------------------------------------------------------
                // Define copy parameters
                // ------------------------------------------------------------
                int sourceColumnIndex = 0;          // Column A (0‑based)
                int columnNumber = 3;               // Number of columns to copy (A‑C)
                int destinationColumnIndex = 4;     // Column E (0‑based) – leave a gap for clarity

                // ------------------------------------------------------------
                // Preserve merged areas that intersect the source range
                // ------------------------------------------------------------
                // Get all merged areas before copying
                CellArea[] mergedAreas = cells.GetMergedAreas();

                // ------------------------------------------------------------
                // Perform the column copy with all data and formatting
                // ------------------------------------------------------------
                PasteOptions pasteOptions = new PasteOptions
                {
                    PasteType = PasteType.All   // copy values, formats, merged cells, etc.
                };
                cells.CopyColumns(cells, sourceColumnIndex, destinationColumnIndex, columnNumber, pasteOptions);

                // ------------------------------------------------------------
                // Re‑create merged cells in the destination columns
                // ------------------------------------------------------------
                int columnShift = destinationColumnIndex - sourceColumnIndex;
                foreach (CellArea area in mergedAreas)
                {
                    // Verify the merged area lies completely within the source columns
                    if (area.StartColumn >= sourceColumnIndex &&
                        area.EndColumn < sourceColumnIndex + columnNumber)
                    {
                        // Calculate new merged area coordinates by shifting columns
                        int newFirstColumn = area.StartColumn + columnShift;
                        int newFirstRow = area.StartRow;
                        int rowCount = area.EndRow - area.StartRow + 1;
                        int columnCount = area.EndColumn - area.StartColumn + 1;

                        // Apply the merged region at the new location
                        cells.Merge(newFirstRow, newFirstColumn, rowCount, columnCount);
                    }
                }

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                workbook.Save("ColumnCopyWithMergedCells.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
