// Title: Copy Rows with Merged Cells while Preserving Merged Regions – Aspose.Cells for .NET
// Description: Demonstrates how to use Aspose.Cells' Cells.CopyRows to copy rows that contain merged cells, retain the merged areas with GetMergedAreas, verify with IsMerged, and save the workbooks.
// Keywords: Aspose.Cells CopyRows merged cells | preserve merged regions .NET | Cells.CopyRows example | GetMergedAreas Aspose | IsMerged property | C# Aspose.Cells copy rows | merged cell handling
// Common Searches: Aspose.Cells copy rows with merged cells | how to keep merged cells when copying rows .NET | CopyRows merged region preservation | GetMergedAreas after copying rows | check IsMerged after CopyRows
// Developer Intent: Copy rows that contain merged cells from one worksheet to another (or within the same sheet) and keep the original merged areas intact.
// Use Cases: Duplicate a multi‑column header row to a later position in a report sheet. | Transfer a formatted template section with merged titles into a generated workbook while preserving layout. | Automated validation that merged regions were copied correctly by enumerating merged areas. | Create a printable version of a sheet by copying rows without losing merged cell formatting.
// AI Prompts: Generate C# code using Aspose.Cells to copy rows 0‑2 from a source worksheet to row 5 in a destination worksheet, ensuring any merged cells are preserved and then list the merged area coordinates. | Explain how Cells.CopyRows treats merged cells in Aspose.Cells and show code that confirms the merged regions after the copy using GetMergedAreas and IsMerged. | Provide a step‑by‑step tutorial for copying rows with merged cells between workbooks, logging each merged area's start and end positions, and saving both files.

using System;
using Aspose.Cells;

// Demonstrates how to use Aspose.Cells' Cells.CopyRows to copy rows that contain merged cells, retain the merged areas with GetMergedAreas, verify with IsMerged, and save the workbooks.
class CopyRowsWithMergedCells
{
    static void Main()
    {
        // Create source workbook and add merged cells
        Workbook sourceWb = new Workbook();
        Worksheet sourceWs = sourceWb.Worksheets[0];
        Cells sourceCells = sourceWs.Cells;

        // Merge A1:B2 (rows 0-1, columns 0-1) and set a value
        sourceCells.Merge(0, 0, 2, 2);
        sourceCells["A1"].PutValue("Merged Header");

        // Add some regular data in the same rows
        sourceCells["C1"].PutValue("Data1");
        sourceCells["C2"].PutValue("Data2");

        // Create destination workbook
        Workbook destWb = new Workbook();
        Worksheet destWs = destWb.Worksheets[0];
        Cells destCells = destWs.Cells;

        // Copy the first two rows (including the merged region) to row index 4 in the destination sheet
        destCells.CopyRows(sourceCells, 0, 4, 2);

        // Verify that merged areas were copied
        CellArea[] mergedAreas = destWs.Cells.GetMergedAreas();
        Console.WriteLine("Merged areas count in destination: " + mergedAreas.Length);
        foreach (CellArea area in mergedAreas)
        {
            Console.WriteLine($"Merged area: StartRow={area.StartRow}, StartColumn={area.StartColumn}, EndRow={area.EndRow}, EndColumn={area.EndColumn}");
        }

        // Check a specific cell within the copied merged region
        bool isMerged = destWs.Cells[4, 0].IsMerged; // Cell A5 (row index 4)
        Console.WriteLine($"Cell A5 IsMerged: {isMerged}");

        // Save the workbooks (optional)
        sourceWb.Save("Source.xlsx");
        destWb.Save("Destination.xlsx");
    }
}
