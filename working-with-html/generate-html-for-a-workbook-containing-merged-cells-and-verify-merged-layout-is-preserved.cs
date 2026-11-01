// Title: Generate an Excel workbook with merged cells, save as XLSX, reload, and confirm merged range using Aspose.Cells for .NET
// AI Prompts: Write a C# console program that creates a new workbook, merges cells A1:C2 on the first worksheet, saves the file as XLSX, loads it back, and checks the MergedCells collection to verify the merge exists. | Adapt the program to accept a custom merge range (e.g., startRow, startColumn, rowCount, columnCount), perform the merge, save and reload the workbook, and output whether the merged area is preserved.
// Common Searches: Aspose.Cells .NET how to verify merged cells after saving and reopening an Excel file | C# check if merged range A1:C2 is retained when loading a workbook with Aspose.Cells | preserve merged cells when exporting to XLSX using Aspose.Cells for .NET | retrieve merged cell collection from a loaded worksheet in Aspose.Cells C#
// Tags: Aspose.Cells merge cells A1:C2 | Aspose.Cells verify merged range after save | C# Aspose.Cells load workbook merged cells | Excel XLSX merged cells preservation Aspose.Cells | Aspose.Cells MergedCells collection usage

using System;
using System.IO;
using Aspose.Cells;

// The example creates a new workbook, merges cells A1:C2 on the first worksheet, saves it as an XLSX file, reloads the workbook, iterates through the worksheet's MergedCells collection to locate the original merged range, and prints whether the merged layout was preserved.
class MergedCellsDemo
{
    static void Main()
    {
        // ---------- Create a new workbook ----------
        // (Using the provided create rule)
        Workbook workbook = new Workbook();

        // Access the first worksheet
        Worksheet sheet = workbook.Worksheets[0];

        // Merge cells A1:C2 (range: 0,0 to 1,2)
        // (Using the provided merge rule)
        sheet.Cells.Merge(0, 0, 2, 3);
        // Set a value in the top‑left cell of the merged area
        sheet.Cells["A1"].PutValue("Merged Area");

        // ---------- Save the workbook ----------
        // (Using the provided save rule)
        string filePath = "MergedCellsDemo.xlsx";
        workbook.Save(filePath, SaveFormat.Xlsx);

        // ---------- Load the workbook ----------
        // (Using the provided load rule)
        Workbook loadedWorkbook = new Workbook(filePath);
        Worksheet loadedSheet = loadedWorkbook.Worksheets[0];

        // ---------- Verify merged layout ----------
        // The MergedCells collection contains all merged ranges.
        bool mergePreserved = false;
        foreach (CellArea area in loadedSheet.Cells.MergedCells)
        {
            // Check if the merged area matches the one we created (A1:C2)
            if (area.StartRow == 0 && area.StartColumn == 0 &&
                area.EndRow == 1 && area.EndColumn == 2)
            {
                mergePreserved = true;
                break;
            }
        }

        // Output verification result
        Console.WriteLine("Merged layout preserved: " + mergePreserved);
    }
}
