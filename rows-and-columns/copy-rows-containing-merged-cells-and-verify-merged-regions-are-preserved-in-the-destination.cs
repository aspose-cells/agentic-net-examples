// Title: Copy rows with merged cells while preserving merged regions – Aspose.Cells for .NET (C#)
// Description: This C# example shows how to create a source workbook, merge a range (A1:B2), copy the first three rows to a new workbook using Cells.CopyRows, retrieve and display the merged areas in the destination sheet, verify the IsMerged flag, and save the file as "CopiedRowsWithMergedCells.xlsx".
// Keywords: Aspose.Cells | CopyRows | merged cells | preserve merged regions | C# Excel automation | Excel row copy | merged area verification | CellArea | IsMerged property | Aspose.Cells .NET
// Common Searches: Aspose.Cells copy rows merged cells | CopyRows preserve merged ranges C# | How to copy rows with merged cells in Aspose.Cells | Verify merged cells after copying rows Aspose | C# copy rows between workbooks Aspose.Cells
// Developer Intent: Copy rows that include merged cells and confirm that the merged ranges remain intact in the target worksheet.
// Use Cases: Replicate a multi‑row header block across multiple sheets while keeping the merge. | Migrate a formatted table with merged title rows into a template workbook. | Automate validation of merged cell structures after row transfer in Excel reports. | Generate a new workbook that reuses existing merged layouts without manual recreation.
// AI Prompts: Generate C# code using Aspose.Cells to copy rows containing merged cells and list the merged areas in the destination sheet. | Explain how Cells.CopyRows treats merged ranges and how to check them with GetMergedAreas and IsMerged. | Show how to copy rows with merged cells while also preserving styles, formulas, and comments. | Provide a step‑by‑step guide to validate merged regions after copying rows in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // This C# example shows how to create a source workbook, merge a range (A1:B2), copy the first three rows to a new workbook using Cells.CopyRows, retrieve and display the merged areas in the destination sheet, verify the IsMerged flag, and save the file as "CopiedRowsWithMergedCells.xlsx".
    public class CopyRowsWithMergedCellsDemo
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
            // ---------- Create source workbook ----------
            Workbook sourceWorkbook = new Workbook();
            Worksheet sourceSheet = sourceWorkbook.Worksheets[0];
            Cells srcCells = sourceSheet.Cells;

            // Fill some data
            srcCells["A1"].PutValue("Header");
            srcCells["A2"].PutValue("Data 1");
            srcCells["A3"].PutValue("Data 2");
            srcCells["B1"].PutValue("SubHeader");
            srcCells["B2"].PutValue("Info 1");
            srcCells["B3"].PutValue("Info 2");

            // Merge cells that span multiple rows (A1:B2)
            srcCells.Merge(0, 0, 2, 2); // Merge A1:B2

            // ---------- Create destination workbook ----------
            Workbook destWorkbook = new Workbook();
            Worksheet destSheet = destWorkbook.Worksheets[0];
            Cells destCells = destSheet.Cells;

            // Copy rows 0-2 (3 rows) from source to destination starting at row 0
            destCells.CopyRows(srcCells, 0, 0, 3);

            // ---------- Verify merged regions in destination ----------
            CellArea[] mergedAreas = destSheet.Cells.GetMergedAreas();

            Console.WriteLine($"Number of merged areas in destination: {mergedAreas.Length}");
            foreach (CellArea area in mergedAreas)
            {
                // Display the address of the top‑left and bottom‑right cells of each merged area
                string startAddress = CellsHelper.CellIndexToName(area.StartRow, area.StartColumn);
                string endAddress = CellsHelper.CellIndexToName(area.EndRow, area.EndColumn);
                Console.WriteLine($"Merged area: {startAddress}:{endAddress}");
            }

            // Additionally, check a specific cell inside the merged region
            bool isMerged = destSheet.Cells["A1"].IsMerged;
            Console.WriteLine($"Cell A1 IsMerged: {isMerged}");

            // ---------- Save the result ----------
            destWorkbook.Save("CopiedRowsWithMergedCells.xlsx");
        }
    }
}
