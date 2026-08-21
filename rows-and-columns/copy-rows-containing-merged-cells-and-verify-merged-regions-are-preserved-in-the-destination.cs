// Title: Copy rows with merged cells while preserving merged regions using Aspose.Cells for .NET
// Description: Demonstrates how to copy all rows from a source worksheet that contains a merged range (A1:C2) to a destination worksheet using Aspose.Cells' Cells.CopyRows method. After copying, the example verifies that the merged region is retained by checking the IsMerged flag and enumerating merged areas with GetMergedAreas, then saves both workbooks.
// Keywords: Aspose.Cells | CopyRows | merged cells | preserve merged regions | .NET Excel | Cells.CopyRows example | GetMergedAreas | IsMerged property | copy rows between workbooks
// Common Searches: Aspose.Cells copy rows with merged cells | preserve merged ranges when copying rows .NET | verify merged areas after CopyRows | how to copy rows between workbooks Aspose.Cells | retain merged cells during row copy
// Developer Intent: Copy rows from one worksheet to another and keep any merged cell ranges intact.
// Use Cases: Clone a template sheet that has a merged header into a new workbook without losing formatting. | Move data rows from a report worksheet to a summary workbook while maintaining column-spanning merged cells. | Generate multiple Excel files with identical layout, including merged cells, by programmatically copying rows.
// AI Prompts: Write C# code with Aspose.Cells to copy rows from a source sheet to a destination sheet and ensure merged cells are preserved. | Show how to validate merged regions after using Cells.CopyRows, including logging each merged area's coordinates. | Explain the behavior of Cells.CopyRows regarding merged cells and any extra steps needed for verification.

using System;
using Aspose.Cells;

namespace AsposeCellsCopyRowsWithMergedCells
{
    // Demonstrates how to copy all rows from a source worksheet that contains a merged range (A1:C2) to a destination worksheet using Aspose.Cells' Cells.CopyRows method. After copying, the example verifies that the merged region is retained by checking the IsMerged flag and enumerating merged areas with GetMergedAreas, then saves both workbooks.
    class Program
    {
        static void Main()
        {
            // ---------- Create source workbook ----------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cells sourceCells = sourceSheet.Cells;

            // Create merged cells spanning rows 0-1 and columns 0-2 (A1:C2)
            sourceCells.Merge(0, 0, 2, 3);
            sourceCells["A1"].PutValue("Merged Header");

            // Add additional data in rows below the merged area
            sourceCells["A3"].PutValue("Row 3 Data");
            sourceCells["B3"].PutValue(123);
            sourceCells["A4"].PutValue("Row 4 Data");
            sourceCells["B4"].PutValue(456);

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            Cells destCells = destSheet.Cells;

            // Copy all rows from source to destination, preserving merged regions
            int totalRowsToCopy = sourceCells.MaxDisplayRange.RowCount;
            destCells.CopyRows(sourceCells, 0, 0, totalRowsToCopy);

            // ---------- Verify merged regions in destination ----------
            // Check if the top-left cell of the merged area is still merged
            bool isMerged = destCells["A1"].IsMerged;
            Console.WriteLine($"Destination cell A1 IsMerged: {isMerged}");

            // Retrieve merged areas and display their coordinates
            CellArea[] mergedAreas = destSheet.Cells.GetMergedAreas();
            Console.WriteLine($"Number of merged areas in destination: {mergedAreas.Length}");
            foreach (CellArea area in mergedAreas)
            {
                Console.WriteLine($"Merged area: StartRow={area.StartRow}, StartColumn={area.StartColumn}, " +
                                  $"EndRow={area.EndRow}, EndColumn={area.EndColumn}");
            }

            // ---------- Save workbooks ----------
            sourceWorkbook.Save("SourceWorkbook.xlsx");
            destWorkbook.Save("DestinationWorkbook.xlsx");
        }
    }
}
