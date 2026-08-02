// Title: Recalculate Formulas After Bulk Row Insertion with Aspose.Cells for .NET
// Description: Demonstrates how to insert multiple rows into an Aspose.Cells worksheet, refresh dynamic‑array and shared formulas, recalculate all dependent calculations, and save the updated workbook using C#.
// Keywords: Aspose.Cells bulk row insert | recalculate formulas C# | RefreshDynamicArrayFormulas | CalculateFormula after insert | shared formula update | Excel row insertion programmatic | .NET spreadsheet automation
// Common Searches: Aspose.Cells recalculate formulas after inserting rows | C# insert rows and update formulas in Excel workbook | how to refresh dynamic array formulas Aspose.Cells | shared formula shift after bulk row insert .NET | programmatic Excel row insertion with formula recalculation
// Developer Intent: Update all worksheet formulas automatically after a bulk row insertion so that calculations remain accurate.
// Use Cases: Add several rows to a sheet and keep formulas that reference shifted cells correct. | Refresh dynamic‑array formulas and then run CalculateFormula to obtain up‑to‑date results. | Persist the workbook with recalculated values after structural changes.
// AI Prompts: Show C# code that inserts rows, refreshes dynamic array formulas, and recalculates all formulas with Aspose.Cells. | Explain how RefreshDynamicArrayFormulas differs from CalculateFormula after a bulk row insertion. | Provide a step‑by‑step guide to preserve shared formulas when inserting multiple rows in an Aspose.Cells workbook.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Demonstrates how to insert multiple rows into an Aspose.Cells worksheet, refresh dynamic‑array and shared formulas, recalculate all dependent calculations, and save the updated workbook using C#.
    public class RecalculateAfterBulkInsertDemo
    {
        // Entry point for the application
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
            // Create a new workbook and get the first worksheet
            Workbook wb = new Workbook();
            Worksheet ws = wb.Worksheets[0];
            Cells cells = ws.Cells;

            // -------------------------------------------------
            // 1. Populate initial data
            // -------------------------------------------------
            // Column A: simple numeric values
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue(i + 1); // A1..A5 = 1,2,3,4,5
            }

            // Column B: formula that depends on column A (B = A * 10)
            // Use a shared formula for the range B1:B5
            cells[0, 1].SetSharedFormula("=A1*10", 5, 1, new FormulaParseOptions());

            // Calculate the initial formulas so that dependent values are set
            wb.CalculateFormula();

            // -------------------------------------------------
            // 2. Bulk insert rows
            // -------------------------------------------------
            // Insert 3 rows starting at row index 2 (zero‑based, i.e., before original row 3)
            // The 'true' flag updates references in other worksheets if needed
            cells.InsertRows(2, 3, true);

            // Optionally fill the newly inserted rows with data
            cells[2, 0].PutValue(100); // New A3
            cells[3, 0].PutValue(200); // New A4
            cells[4, 0].PutValue(300); // New A5

            // -------------------------------------------------
            // 3. Recalculate formulas after insertion
            // -------------------------------------------------
            // Refresh dynamic array formulas (if any) and then calculate all formulas
            wb.RefreshDynamicArrayFormulas(true);
            wb.CalculateFormula();

            // -------------------------------------------------
            // 4. Verify results
            // -------------------------------------------------
            Console.WriteLine("After inserting rows and recalculation:");
            for (int i = 0; i <= cells.MaxDataRow; i++)
            {
                Console.WriteLine($"Row {i + 1}: A = {cells[i, 0].Value}, B = {cells[i, 1].Value}");
            }

            // -------------------------------------------------
            // 5. Save the workbook (lifecycle rule)
            // -------------------------------------------------
            wb.Save("RecalculateAfterBulkInsert.xlsx");
        }
    }
}
