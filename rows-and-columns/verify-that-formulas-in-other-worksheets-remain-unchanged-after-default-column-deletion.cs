// Title: Verify that formulas in other worksheets remain unchanged after deleting column A with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to delete column A from a chosen worksheet, then iterate through all other worksheets and compare each cell's formula with the pre‑deletion values, reporting any additions, modifications, or removals. | Create a method that captures every formula in all sheets except the target sheet, removes the first column, and validates that no formula in the remaining sheets has changed, outputting a summary of the verification.
// Common Searches: Aspose.Cells C# verify formulas are unchanged after deleting column A from a worksheet | how to compare Excel formulas before and after column removal using Aspose.Cells | detect formula changes across multiple sheets after removing a column in .NET | C# Aspose.Cells preserve formula integrity in other worksheets when a column is deleted
// Tags: Aspose.Cells delete column A | Aspose.Cells compare worksheet formulas | C# validate Excel formula integrity | Aspose.Cells track formula changes after column removal | Excel workbook formula verification Aspose.Cells

using System;
using System.Collections.Generic;
using Aspose.Cells;

// The example loads an Excel workbook, records all formulas from every worksheet except the one where column A will be removed, deletes that column using Aspose.Cells, then re‑examines the other sheets to ensure each formula matches its original value, reporting any discrepancies before saving the updated file.
class VerifyFormulasAfterColumnDeletion
{
    static void Main()
    {
        // Load the workbook (replace with your actual file path)
        Workbook workbook = new Workbook("input.xlsx");

        // Choose the worksheet from which the default column will be deleted
        // Here we assume the first worksheet; adjust as needed
        Worksheet targetSheet = workbook.Worksheets[0];

        // Store formulas from all other worksheets before deletion
        var originalFormulas = new Dictionary<string, Dictionary<string, string>>();

        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];
            if (ws == targetSheet) continue; // skip the sheet where column will be deleted

            var sheetFormulas = new Dictionary<string, string>();
            Cells cells = ws.Cells;

            // Iterate through all cells that contain formulas
            foreach (Cell cell in cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    sheetFormulas[cell.Name] = cell.Formula;
                }
            }

            originalFormulas[ws.Name] = sheetFormulas;
        }

        // Delete the default column (column A, index 0) from the target worksheet
        targetSheet.Cells.DeleteColumn(0);

        // Verify that formulas in other worksheets remain unchanged
        bool allUnchanged = true;

        for (int i = 0; i < workbook.Worksheets.Count; i++)
        {
            Worksheet ws = workbook.Worksheets[i];
            if (ws == targetSheet) continue; // skip the modified sheet

            var beforeFormulas = originalFormulas[ws.Name];
            Cells cells = ws.Cells;

            foreach (Cell cell in cells)
            {
                if (!string.IsNullOrEmpty(cell.Formula))
                {
                    string cellName = cell.Name;
                    if (!beforeFormulas.TryGetValue(cellName, out string originalFormula))
                    {
                        Console.WriteLine($"New formula detected in sheet '{ws.Name}' at {cellName}: {cell.Formula}");
                        allUnchanged = false;
                    }
                    else if (!string.Equals(originalFormula, cell.Formula, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"Formula changed in sheet '{ws.Name}' at {cellName}: before='{originalFormula}', after='{cell.Formula}'");
                        allUnchanged = false;
                    }
                }
            }

            // Check for any formulas that disappeared
            foreach (var kvp in beforeFormulas)
            {
                if (cells[kvp.Key].Formula == null)
                {
                    Console.WriteLine($"Formula removed in sheet '{ws.Name}' at {kvp.Key}: was '{kvp.Value}'");
                    allUnchanged = false;
                }
            }
        }

        if (allUnchanged)
        {
            Console.WriteLine("All formulas in other worksheets remain unchanged after column deletion.");
        }
        else
        {
            Console.WriteLine("Some formulas were altered or removed after column deletion.");
        }

        // Optionally, save the modified workbook
        workbook.Save("output.xlsx");
    }
}
