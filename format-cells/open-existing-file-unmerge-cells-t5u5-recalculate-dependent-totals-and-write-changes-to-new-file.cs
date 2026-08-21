// Title: C# – Unmerge Cells T5:U5, Recalculate Formulas, and Save as a New Workbook with Aspose.Cells
// Description: Loads input.xlsx, accesses the first worksheet, removes the merged range T5:U5, forces a full formula recalculation to update any dependent totals, and writes the result to output.xlsx using Aspose.Cells for .NET.
// Keywords: Aspose.Cells unmerge range C# | remove merged cells .NET | recalculate Excel formulas programmatically | Workbook.CalculateFormula example | save modified workbook Aspose.Cells | Excel automation unmerge cells | update dependent totals after unmerge | C# Aspose.Cells tutorial | Excel file processing Aspose
// Common Searches: how to unmerge a specific range with Aspose.Cells C# | Aspose.Cells recalculate formulas after unmerge | save a copy of an Excel file after removing merged cells | C# code to unmerge T5:U5 and refresh totals | Aspose.Cells unmerge cells and recalc all formulas
// Developer Intent: Remove the merged cells in range T5:U5, recalculate every formula so dependent totals are correct, and export the updated workbook as a new file.
// Use Cases: Prepare a financial report for systems that reject merged cells while keeping all calculated values accurate. | Create a clean version of a spreadsheet before feeding it to an automated data‑extraction pipeline that cannot handle merged ranges. | Generate a copy of a template after structural changes (e.g., unmerging headers) without losing any formula‑driven totals.
// AI Prompts: Write C# code that uses Aspose.Cells to unmerge the range T5:U5, run CalculateFormula, and save the workbook as output.xlsx. | Explain step‑by‑step how to remove merged cells in an Excel file and refresh dependent calculations with Aspose.Cells for .NET. | Provide a minimal Aspose.Cells example that loads input.xlsx, unmerges T5:U5, recalculates all formulas, and writes the result to a new file.

using System;
using Aspose.Cells;

// Loads input.xlsx, accesses the first worksheet, removes the merged range T5:U5, forces a full formula recalculation to update any dependent totals, and writes the result to output.xlsx using Aspose.Cells for .NET.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];

        // Unmerge the range T5:U5
        // Row index is zero‑based (5th row -> 4), column T is index 19, total rows = 1, total columns = 2
        worksheet.Cells.UnMerge(4, 19, 1, 2);

        // Recalculate all formulas so dependent totals are updated
        workbook.CalculateFormula();

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
