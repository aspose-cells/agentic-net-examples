// Title: Create an Excel workbook with merged cells, save it, reload, and verify the merged range using Aspose.Cells for .NET (C#)
// AI Prompts: Write C# code that uses Aspose.Cells to merge cells A1:C1, save the workbook as XLSX, reopen it, and programmatically confirm the merged area still exists. | Provide a .NET example that demonstrates how to check the MergedCells collection after loading a saved Excel file to ensure the merged header is retained.
// Common Searches: Aspose.Cells C# how to check if merged cells are retained after saving the workbook | verify merged cell range A1:C1 after reloading Excel file with Aspose.Cells .NET | C# example for preserving merged cells when exporting to XLSX using Aspose.Cells
// Tags: Aspose.Cells merge cells and validate after save | C# Aspose.Cells preserve merged range | Excel merged cells verification with Aspose.Cells | Aspose.Cells load workbook merged cells check

using System;
using Aspose.Cells;

// The program creates a new workbook, merges cells A1:C1, saves it as an XLSX file, reloads the workbook, iterates through the MergedCells collection to confirm the specific merged range is still present, and outputs whether the merged layout is preserved.
class Program
{
    static void Main()
    {
        // Create a new workbook
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet worksheet = workbook.Worksheets[0];

        // Merge cells A1:C1 (row 0, column 0, 1 row, 3 columns)
        worksheet.Cells.Merge(0, 0, 1, 3);

        // Put a value into the merged cell
        worksheet.Cells[0, 0].PutValue("Merged Header");

        // Save the workbook
        string filePath = "MergedCellsDemo.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        // Load the workbook back
        Workbook loadedWorkbook = new Workbook(filePath);
        Worksheet loadedWorksheet = loadedWorkbook.Worksheets[0];

        // Verify that the merged layout is preserved
        bool mergedPreserved = false;
        foreach (CellArea area in loadedWorksheet.Cells.MergedCells)
        {
            // Check for the specific merged range A1:C1
            if (area.StartRow == 0 && area.StartColumn == 0 && area.EndRow == 0 && area.EndColumn == 2)
            {
                mergedPreserved = true;
                break;
            }
        }

        Console.WriteLine(mergedPreserved
            ? "Merged layout is preserved."
            : "Merged layout is NOT preserved.");
    }
}
