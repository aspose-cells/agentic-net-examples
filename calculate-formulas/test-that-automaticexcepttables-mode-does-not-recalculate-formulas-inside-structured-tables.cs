// Title: Aspose.Cells C# – Verify AutomaticExceptTable Mode Leaves Table Formulas Static
// Description: The sample builds a workbook, inserts a ListObject covering A1:B3, sets a structured‑reference formula in the table, changes the calculation setting to CalcModeType.AutomaticExceptTable, performs an initial evaluation, modifies the source value, recalculates, and shows that the table cell keeps its original result while a normal formula updates. The workbook is saved to demonstrate the full creation‑calculation‑persistence flow.
// Keywords: Aspose.Cells | C# | .NET | AutomaticExceptTable | CalcModeType | structured table formula | ListObject | Excel table calculation | formula recalculation control | prevent table formula update | workbook save
// Common Searches: Aspose.Cells AutomaticExceptTable example C# | how to stop table formulas from recalculating in Aspose.Cells | CalcModeType AutomaticExceptTable usage | compare table and regular formula recalculation Aspose.Cells | structured reference formula test with Aspose.Cells
// Developer Intent: Confirm that enabling CalcModeType.AutomaticExceptTable prevents formulas inside a ListObject from being recomputed while ordinary cell formulas continue to update.
// Use Cases: Create a workbook, add a structured table, assign a formula with a structured reference, switch to AutomaticExceptTable mode, change a dependent cell, recalc, and verify the table value stays unchanged. | Place a regular formula alongside the table formula to illustrate that only the non‑table expression reacts to data changes under this setting. | Persist the workbook after the test to showcase end‑to‑end handling of creation, calculation, and file output.
// AI Prompts: Write a C# unit test with Aspose.Cells that asserts table formulas remain constant after source data changes when CalcModeType.AutomaticExceptTable is active. | Explain the internal workflow of Aspose.Cells when AutomaticExceptTable mode is on, focusing on how structured tables and regular cells are treated differently during recalculation. | Modify the example to log before‑and‑after values to a JSON file instead of writing to the console.

using System;
using System.IO;
using Aspose.Cells;
using Aspose.Cells.Tables;

namespace AutomaticExceptTableTest
{
    // The sample builds a workbook, inserts a ListObject covering A1:B3, sets a structured‑reference formula in the table, changes the calculation setting to CalcModeType.AutomaticExceptTable, performs an initial evaluation, modifies the source value, recalculates, and shows that the table cell keeps its original result while a normal formula updates. The workbook is saved to demonstrate the full creation‑calculation‑persistence flow.
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];
                Cells cells = sheet.Cells;

                // Set up a simple table with a header and two data rows
                cells["A1"].PutValue("Value");   // Header for column A
                cells["B1"].PutValue("Result");  // Header for column B (will hold table formula)

                cells["A2"].PutValue(10);        // First data row
                cells["A3"].PutValue(20);        // Second data row

                // Create a ListObject (structured table) covering the range A1:B3
                int tableIndex = sheet.ListObjects.Add("A1", "B3", true);
                ListObject table = sheet.ListObjects[tableIndex];

                // Set a table formula in the first data row of column B using a structured reference
                // This formula will be automatically filled down for the whole column by Excel behavior
                cells["B2"].Formula = "=[@Value]*2";

                // Set calculation mode to AutomaticExceptTable
                workbook.Settings.FormulaSettings.CalculationMode = CalcModeType.AutomaticExceptTable;

                // Initial calculation
                workbook.CalculateFormula();

                // Capture the result of the table formula before any changes
                int beforeChange = cells["B2"].IntValue; // Expected 20 (10*2)

                // Also add a regular (non‑table) formula for comparison
                cells["C1"].Formula = "=A2*2";
                workbook.CalculateFormula();
                int regularBefore = cells["C1"].IntValue; // Expected 20

                // Change the source value that both formulas depend on
                cells["A2"].PutValue(30);

                // Recalculate formulas
                workbook.CalculateFormula();

                // Capture the results after the change
                int afterChange = cells["B2"].IntValue;   // Should remain 20 because table formulas are not recalculated
                int regularAfter = cells["C1"].IntValue; // Should become 60 because regular formulas are recalculated

                // Output the results
                Console.WriteLine($"Table formula before change: {beforeChange}");
                Console.WriteLine($"Table formula after change (should be unchanged): {afterChange}");
                Console.WriteLine($"Regular formula before change: {regularBefore}");
                Console.WriteLine($"Regular formula after change (should be updated): {regularAfter}");

                // Save the workbook (optional, demonstrates lifecycle rule usage)
                string outputPath = "AutomaticExceptTableTest.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}
