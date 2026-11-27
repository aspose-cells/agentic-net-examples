// Title: Create and protect a merged cell range (R2:S2) in an Excel workbook using Aspose.Cells for .NET and save as XLSX
// Common Searches: Aspose.Cells C# merge cells R2 S2 and make them read‑only | How to protect only a merged range in an Excel worksheet with Aspose.Cells .NET | Saving a workbook with a protected merged region to XLSX using Aspose.Cells | C# code to lock merged cells while allowing other cells to be edited in Aspose.Cells | Apply worksheet protection after unlocking all cells in Aspose.Cells
// Tags: CellArea merge R2:S2 Aspose.Cells | apply locked style to range Aspose.Cells C# | worksheet protect read‑only merged region Aspose.Cells | workbook save as XLSX Aspose.Cells | unlock all cells before worksheet protection Aspose.Cells

using System;
using Aspose.Cells;

// The example creates a new Workbook, unlocks all cells, merges the range R2:S2, applies a locked style to that merged area, protects the worksheet so only the merged cells are read‑only, and saves the file as MergedProtected.xlsx in XLSX format using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Get the first worksheet
            Worksheet sheet = workbook.Worksheets[0];

            // Unlock all cells so they can be edited after protection
            sheet.Cells.Style.IsLocked = false;

            // Merge cells R2:S2
            CellArea mergedArea = CellArea.CreateCellArea("R2", "S2");
            sheet.Cells.Merge(mergedArea.StartRow, mergedArea.StartColumn,
                              mergedArea.EndRow, mergedArea.EndColumn);

            // Apply a locked style to the merged region (read‑only)
            Style lockedStyle = workbook.CreateStyle();
            lockedStyle.IsLocked = true;
            StyleFlag flag = new StyleFlag();
            flag.Locked = true;
            sheet.Cells.CreateRange("R2:S2").ApplyStyle(lockedStyle, flag);

            // Protect the worksheet – only locked cells (the merged region) become read‑only
            sheet.Protect(ProtectionType.All);

            // Save the workbook as XLSX
            workbook.Save("MergedProtected.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}
