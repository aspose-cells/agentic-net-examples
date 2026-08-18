// Title: CopyColumns to Duplicate Adjacent Columns and Preserve Merged Cells – Aspose.Cells for .NET (C#)
// Description: This example creates a workbook, merges cells A1:B1, then uses Cells.CopyColumns to copy columns A‑B to C‑D. After the copy it enumerates the original merged areas, shifts their coordinates, and re‑applies the merges so the duplicated columns retain the same merged layout. The workbook is saved as DuplicatedColumnsWithMergedCells.xlsx.
// Keywords: Aspose.Cells | CopyColumns | duplicate columns | preserve merged cells | C# | .NET | Excel automation | merged regions | cells.Merge | copy columns with formats
// Common Searches: Aspose.Cells copy columns keep merged cells | CopyColumns method example C# | duplicate adjacent columns Aspose.Cells | how to copy merged cells in Excel using Aspose | preserve merged areas after copying columns .NET
// Developer Intent: Copy a set of adjacent columns to another location and ensure any merged cells are duplicated correctly.
// Use Cases: Replicate a header row that spans multiple columns across a report. | Create a copy of a data block with a merged title for a new worksheet. | Generate monthly sheets by cloning formatted columns, including merged cells. | Build a template where column groups are repeated with the same layout.
// AI Prompts: Write C# code that uses Aspose.Cells to copy columns A‑B to C‑D and duplicate any merged cells. | Explain how to adjust CellArea coordinates after using Cells.CopyColumns to keep merged regions. | Provide a step‑by‑step tutorial for duplicating adjacent columns while preserving merges in an Excel workbook with Aspose.Cells. | Show how to retrieve original merged areas and re‑apply them after copying columns in .NET.

using System;
using Aspose.Cells;

namespace DuplicateColumnsWithMergedCells
{
    // This example creates a workbook, merges cells A1:B1, then uses Cells.CopyColumns to copy columns A‑B to C‑D. After the copy it enumerates the original merged areas, shifts their coordinates, and re‑applies the merges so the duplicated columns retain the same merged layout. The workbook is saved as DuplicatedColumnsWithMergedCells.xlsx.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // -------------------------------------------------
                // Sample data and a merged region spanning columns A and B
                // -------------------------------------------------
                cells["A1"].PutValue("Header1");
                cells["B1"].PutValue("Header2");
                cells["A2"].PutValue(100);
                cells["B2"].PutValue(200);

                // Merge cells A1:B1 (first row, first two columns)
                cells.Merge(0, 0, 1, 2); // Row 0, Column 0, 1 row, 2 columns

                // -------------------------------------------------
                // Duplicate the adjacent columns (A and B) to columns C and D
                // -------------------------------------------------
                int sourceColumnIndex = 0;          // Column A (zero‑based)
                int columnCount = 2;                // Number of columns to copy (A and B)
                int destinationColumnIndex = 2;     // Column C (where the copy will start)

                // Copy the columns together with data, formats, and merged information
                cells.CopyColumns(cells, sourceColumnIndex, destinationColumnIndex, columnCount);

                // -------------------------------------------------
                // Preserve merged cells in the newly copied columns
                // -------------------------------------------------
                // Get all merged areas that existed before the copy
                CellArea[] originalMergedAreas = cells.GetMergedAreas();

                foreach (CellArea area in originalMergedAreas)
                {
                    // Determine if the merged area lies completely within the source columns
                    if (area.StartColumn >= sourceColumnIndex &&
                        area.EndColumn < sourceColumnIndex + columnCount)
                    {
                        // Calculate the column offset to the destination location
                        int columnOffset = destinationColumnIndex - sourceColumnIndex;

                        // Shift the merged area to the destination columns
                        int newFirstRow = area.StartRow;
                        int newFirstColumn = area.StartColumn + columnOffset;
                        int rowCount = area.EndRow - area.StartRow + 1;
                        int colCount = area.EndColumn - area.StartColumn + 1;

                        // Apply the merged region to the destination columns
                        cells.Merge(newFirstRow, newFirstColumn, rowCount, colCount);
                    }
                }

                // -------------------------------------------------
                // Save the workbook
                // -------------------------------------------------
                string outputPath = "DuplicatedColumnsWithMergedCells.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved successfully to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred: {ex.Message}");
            }
        }
    }
}
