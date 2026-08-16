// Title: Recalculate Formulas After Bulk Row Insertion with Aspose.Cells for .NET (C#)
// Description: This example shows how to insert multiple rows in an Aspose.Cells worksheet, update formula references with the updateReference flag, and then force a full recalculation using Workbook.CalculateFormula. It also demonstrates refreshing dynamic‑array formulas via Workbook.RefreshDynamicArrayFormulas so that all dependent calculations remain accurate before saving the file.
// Keywords: Aspose.Cells bulk insert rows | recalculate formulas C# | Workbook.CalculateFormula | RefreshDynamicArrayFormulas | updateReference flag | Excel formula refresh after InsertRows | .NET spreadsheet automation
// Common Searches: how to refresh formulas after inserting rows Aspose.Cells | Aspose.Cells recalculate dependent formulas .NET | update cell references when adding multiple rows | C# insert rows and recalculate workbook | dynamic array formula refresh Aspose.Cells
// Developer Intent: Refresh all workbook formulas after a bulk row insertion to keep dependent calculations correct.
// Use Cases: Insert a block of rows and automatically adjust existing formula references. | Force a complete calculation pass to ensure numeric results reflect new row positions. | Refresh spilled dynamic‑array formulas after bulk insertion before saving the workbook.
// AI Prompts: Generate C# code that inserts rows with Cells.InsertRows, then calls CalculateFormula and RefreshDynamicArrayFormulas to update all formulas in an Aspose.Cells workbook. | Explain the role of the updateReference parameter in Cells.InsertRows and why a subsequent CalculateFormula call is necessary. | Provide a step‑by‑step verification method for checking recalculated cell values after bulk row insertion using Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace RecalculateAfterBulkInsert
{
    // This example shows how to insert multiple rows in an Aspose.Cells worksheet, update formula references with the updateReference flag, and then force a full recalculation using Workbook.CalculateFormula. It also demonstrates refreshing dynamic‑array formulas via Workbook.RefreshDynamicArrayFormulas so that all dependent calculations remain accurate before saving the file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // ------------------------------------------------------------
            // Set up initial data and formulas that depend on each other
            // ------------------------------------------------------------
            // Column A – base values
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);

            // Column B – formula that references column A
            cells["B1"].Formula = "=A1*2";
            cells["B2"].Formula = "=A2*2";
            cells["B3"].Formula = "=A3*2";

            // Column C – formula that sums the values in column B
            cells["C1"].Formula = "=SUM(B1:B3)";

            // ------------------------------------------------------------
            // Perform bulk row insertion (e.g., insert 5 rows after row 2)
            // ------------------------------------------------------------
            // InsertRows(rowIndex, totalRows, updateReference)
            // rowIndex is zero‑based, so 2 means after the existing rows 0 and 1
            cells.InsertRows(2, 5, true);

            // ------------------------------------------------------------
            // Recalculate all formulas to reflect the new row positions
            // ------------------------------------------------------------
            // For regular formulas
            workbook.CalculateFormula();

            // If the workbook contains dynamic array formulas, refresh them as well
            // The 'true' flag also forces calculation of the spilled ranges
            workbook.RefreshDynamicArrayFormulas(true);

            // ------------------------------------------------------------
            // Output some results to verify correct recalculation
            // ------------------------------------------------------------
            Console.WriteLine("After inserting rows and recalculating:");
            Console.WriteLine($"B1 (should be 20) = {cells["B1"].Value}");
            Console.WriteLine($"B2 (should be 40) = {cells["B2"].Value}");
            Console.WriteLine($"B3 (should be 60) = {cells["B3"].Value}");
            Console.WriteLine($"C1 (sum of B1:B3) = {cells["C1"].Value}");

            // ------------------------------------------------------------
            // Save the workbook (lifecycle rule: use provided save logic)
            // ------------------------------------------------------------
            workbook.Save("RecalculatedAfterInsert.xlsx");
        }
    }
}
