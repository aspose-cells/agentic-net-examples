// Title: Duplicate Adjacent Columns While Preserving Merged Cells – Aspose.Cells for .NET (C#)
// Description: This example shows how to copy a block of adjacent columns with Aspose.Cells' CopyColumns method, capture any merged cells in the source range, and re‑apply those merges to the destination columns so the layout remains unchanged. The workbook is saved as DuplicatedColumnsWithMergedCells.xlsx.
// Keywords: Aspose.Cells CopyColumns | duplicate columns merged cells | preserve merged cells Aspose | C# Aspose.Cells copy columns | copy columns with merged cells .NET | Aspose.Cells merge handling | duplicate adjacent columns | CopyColumns merged area | Aspose.Cells example C#
// Common Searches: copy columns keep merged cells Aspose.Cells C# | duplicate adjacent columns Aspose.Cells .NET | how to preserve merges when using CopyColumns | Aspose.Cells copy columns merged range example | C# copy columns with merged cells Aspose
// Developer Intent: Copy a set of neighboring columns and retain all merged‑cell formatting in the copied area.
// Use Cases: Replicate a header that spans multiple columns for a side‑by‑side report layout. | Duplicate a data section with a merged title to create comparison tables. | Generate a template where the same column group appears in several worksheet regions without losing merge definitions.
// AI Prompts: Write C# code using Aspose.Cells to copy columns A‑B to D‑E and automatically preserve any merged cells. | Provide a method that extracts merged areas from a source column range, copies the columns with CopyColumns, and re‑creates the merges in the target range. | Explain how to calculate the column offset after CopyColumns and re‑merge cells to keep the original layout intact.

using System;
using System.Collections.Generic;
using Aspose.Cells;

// This example shows how to copy a block of adjacent columns with Aspose.Cells' CopyColumns method, capture any merged cells in the source range, and re‑apply those merges to the destination columns so the layout remains unchanged. The workbook is saved as DuplicatedColumnsWithMergedCells.xlsx.
class DuplicateColumnsWithMergedCells
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            cells["A1"].PutValue("Header1");
            cells["B1"].PutValue("Header2");
            cells["C1"].PutValue("Header3");
            cells["A2"].PutValue(100);
            cells["B2"].PutValue(200);
            cells["C2"].PutValue(300);

            // Merge cells A1:B1 (spanning two columns)
            cells.Merge(0, 0, 1, 2); // firstRow, firstColumn, totalRows, totalColumns

            // Define source and destination columns (duplicate columns A and B to D and E)
            int sourceColumnIndex = 0;      // Column A (zero‑based)
            int columnNumber = 2;           // Number of columns to copy (A and B)
            int destinationColumnIndex = 3; // Column D (zero‑based)

            // Capture merged areas that intersect the source columns
            List<CellArea> sourceMerges = new List<CellArea>();
            foreach (CellArea area in cells.GetMergedAreas())
            {
                if (area.StartColumn >= sourceColumnIndex && area.StartColumn < sourceColumnIndex + columnNumber)
                {
                    sourceMerges.Add(area);
                }
            }

            // Copy the columns
            cells.CopyColumns(cells, sourceColumnIndex, destinationColumnIndex, columnNumber);

            // Re‑apply the merged cells in the destination range
            int columnOffset = destinationColumnIndex - sourceColumnIndex;
            foreach (CellArea srcArea in sourceMerges)
            {
                CellArea destArea = srcArea;
                destArea.StartColumn += columnOffset;
                destArea.EndColumn += columnOffset;

                int rowCount = destArea.EndRow - destArea.StartRow + 1;
                int columnCount = destArea.EndColumn - destArea.StartColumn + 1;

                cells.Merge(destArea.StartRow, destArea.StartColumn, rowCount, columnCount);
            }

            // Save the workbook
            workbook.Save("DuplicatedColumnsWithMergedCells.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
    }
}
