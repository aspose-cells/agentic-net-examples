// Title: C# – Unmerge cells T5:U5, recalculate formulas, and save as a new workbook with Aspose.Cells
// Description: Loads an existing Excel file, unmerges the range T5:U5 using zero‑based indices, forces a full formula recalculation to refresh dependent totals, and writes the updated workbook to a new file.
// Keywords: Aspose.Cells unmerge range | C# Excel unmerge cells | recalculate formulas Aspose.Cells | save modified workbook .NET | update dependent totals Excel
// Common Searches: Aspose.Cells example to unmerge T5 U5 in C# | how to recalculate formulas after unmerging cells Aspose.Cells | C# code to split merged cells and save new Excel file | unmerge specific range and refresh totals Aspose.Cells
// Developer Intent: Remove the merge on T5:U5, recalculate all formulas, and export the workbook as a new file.
// Use Cases: Prepare a template for systems that reject merged cells by separating headers and updating calculations. | Create a clean copy of a report where merged cells interfere with data aggregation. | Refresh totals after structural changes to ensure formulas reference the correct cell ranges.
// AI Prompts: Write C# code that uses Aspose.Cells to unmerge T5:U5, recalculate all formulas, and save the result as output.xlsx. | Explain why invoking workbook.CalculateFormula is required after unmerging cells in an Aspose.Cells workbook. | Provide a step‑by‑step tutorial for modifying an existing Excel file: unmerge a specific range and update dependent totals with Aspose.Cells for .NET.

using Aspose.Cells;

// Loads an existing Excel file, unmerges the range T5:U5 using zero‑based indices, forces a full formula recalculation to refresh dependent totals, and writes the updated workbook to a new file.
class Program
{
    static void Main()
    {
        // Load the existing workbook
        Workbook workbook = new Workbook("input.xlsx");

        // Access the first worksheet (adjust index if needed)
        Worksheet worksheet = workbook.Worksheets[0];
        Cells cells = worksheet.Cells;

        // Unmerge the range T5:U5
        // Row index for row 5 = 4 (zero‑based)
        // Column index for column T = 19 (zero‑based)
        // 1 row, 2 columns (T and U)
        cells.UnMerge(4, 19, 1, 2);

        // Recalculate all formulas so dependent totals are updated
        workbook.CalculateFormula();

        // Save the modified workbook to a new file
        workbook.Save("output.xlsx");
    }
}
