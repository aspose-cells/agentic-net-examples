// Title: C# – Unmerge C6:E7, Recalculate Formulas, and Save Workbook with Aspose.Cells
// Description: Loads an existing Excel file, accesses the first worksheet, removes the merged range C6:E7 using Cells.UnMerge (zero‑based indices), forces a full formula recalculation, and writes the result to a new file.
// Keywords: Aspose.Cells | C# unmerge range | Cells.UnMerge | recalculate formulas | save workbook | .NET Excel processing | Excel unmerge programmatically
// Common Searches: Aspose.Cells unmerge specific range C6:E7 C# | How to trigger formula calculation after unmerging cells in Aspose.Cells | Save modified workbook after unmerge using Aspose.Cells .NET
// Developer Intent: Programmatically split a merged block, update dependent calculations, and persist the changes.
// Use Cases: Prepare a sheet for CSV export where merged cells are not allowed | Ensure formula results stay correct after structural edits such as unmerging | Automate cleanup of legacy workbooks before further data processing
// AI Prompts: Write C# code that uses Aspose.Cells to unmerge the range C6:E7, recalculate all formulas, and save the file. | Explain the mapping between Excel address C6:E7 and the parameters of Cells.UnMerge. | Add comprehensive error handling for missing files and calculation errors in the unmerge‑and‑recalculate workflow.

using System;
using Aspose.Cells;

namespace AsposeCellsUnmergeAndRecalculate
{
    // Loads an existing Excel file, accesses the first worksheet, removes the merged range C6:E7 using Cells.UnMerge (zero‑based indices), forces a full formula recalculation, and writes the result to a new file.
    class Program
    {
        static void Main()
        {
            // Load an existing workbook from file
            Workbook workbook = new Workbook("input.xlsx");

            // Access the first worksheet (index 0)
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Unmerge the previously merged range C6:E7
            // C -> column index 2, row 6 -> index 5 (zero‑based)
            // Total rows = 2 (rows 6 and 7), total columns = 3 (C, D, E)
            cells.UnMerge(5, 2, 2, 3);

            // Recalculate all formulas in the workbook
            workbook.CalculateFormula();

            // Save the modified workbook to a new file
            workbook.Save("output.xlsx");
        }
    }
}
