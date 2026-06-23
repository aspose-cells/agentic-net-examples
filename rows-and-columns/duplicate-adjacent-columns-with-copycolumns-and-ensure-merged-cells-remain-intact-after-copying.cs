using System;
using System.IO;
using Aspose.Cells;

namespace AsposeCellsCopyColumnsWithMergedCells
{
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
                // Sample data: create two source columns (A and B) with a merged area
                // ------------------------------------------------------------
                cells["A1"].PutValue("Header1");
                cells["B1"].PutValue("Header2");
                cells["A2"].PutValue(100);
                cells["B2"].PutValue(200);
                cells["A3"].PutValue(300);
                cells["B3"].PutValue(400);

                // Merge cells A1:B2 (first two rows across the two columns)
                cells.Merge(0, 0, 2, 2); // firstRow=0, firstColumn=0, totalRows=2, totalColumns=2

                // ------------------------------------------------------------
                // Remember the merged areas that belong to the source columns
                // ------------------------------------------------------------
                CellArea[] sourceMergedAreas = cells.GetMergedAreas();

                // ------------------------------------------------------------
                // Insert empty columns where the duplicated columns will be placed
                // ------------------------------------------------------------
                // Insert two new columns starting at index 2 (C).
                cells.InsertColumns(2, 2);

                // ------------------------------------------------------------
                // Copy the two source columns (0 and 1) to the new location (starting at index 2)
                // ------------------------------------------------------------
                // Use the overload that copies data and formats.
                cells.CopyColumns(cells, 0, 2, 2);

                // ------------------------------------------------------------
                // Replicate the merged cells for the copied columns
                // ------------------------------------------------------------
                foreach (CellArea area in sourceMergedAreas)
                {
                    // Check if the merged area is within the source column range (0‑1)
                    if (area.StartColumn >= 0 && area.EndColumn <= 1)
                    {
                        // Offset for the destination columns
                        int columnOffset = 2; // destination start column index

                        // Calculate row/column span
                        int rowCount = area.EndRow - area.StartRow + 1;
                        int columnCount = area.EndColumn - area.StartColumn + 1;

                        // Apply the merge in the destination location
                        cells.Merge(area.StartRow, area.StartColumn + columnOffset, rowCount, columnCount);
                    }
                }

                // ------------------------------------------------------------
                // Save the workbook
                // ------------------------------------------------------------
                string outputPath = "CopyColumnsWithMergedCells.xlsx";

                // Ensure the output directory exists (handle possible null)
                string? outputDir = Path.GetDirectoryName(Path.GetFullPath(outputPath));
                if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
                {
                    Directory.CreateDirectory(outputDir);
                }

                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to '{outputPath}'.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}