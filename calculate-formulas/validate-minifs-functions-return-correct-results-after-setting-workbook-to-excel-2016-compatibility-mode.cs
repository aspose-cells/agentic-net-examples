// Title: Validate MINIFS Formula in Excel 2016 Compatibility Mode with Aspose.Cells for .NET
// Description: This C# example creates a workbook, disables compatibility checking to enable newer functions, populates numeric values and categories, inserts the MINIFS formula (=MINIFS(A2:A6,B2:B6,"X")), calculates the result, saves and reloads the file, recalculates, and verifies that the minimum value for category "X" is 8 both before and after persistence.
// Keywords: Aspose.Cells MINIFS .NET | CheckCompatibility false | Excel 2016 compatibility mode | formula calculation Aspose.Cells | save and reload workbook | validate MINIFS result | C# Excel functions | calculate MINIFS after load
// Common Searches: How to use MINIFS with Aspose.Cells in C# | Set Excel 2016 compatibility mode Aspose.Cells | Validate MINIFS calculation after saving workbook | Recalculate formulas after loading a workbook Aspose.Cells | Disable compatibility checks for new Excel functions
// Developer Intent: Ensure the MINIFS function evaluates correctly when the workbook is configured for Excel 2016 features and remains accurate after the file is saved and reopened.
// Use Cases: Confirm that MINIFS returns the expected minimum for a given criterion before persisting the workbook. | Verify that formula results are preserved and recalculated correctly after a save‑load cycle. | Demonstrate disabling compatibility checks to retain modern Excel functions such as MINIFS during automated calculations.
// AI Prompts: Generate C# code using Aspose.Cells to test the MINIFS function in Excel 2016 compatibility mode and assert the expected result of 8. | Explain how to set Workbook.Settings.CheckCompatibility to false so that newer functions like MINIFS are evaluated. | Show how to recalculate formulas after loading a saved workbook and retrieve the MINIFS result for validation.

using System;
using Aspose.Cells;

namespace AsposeCellsMinifsValidation
{
    // This C# example creates a workbook, disables compatibility checking to enable newer functions, populates numeric values and categories, inserts the MINIFS formula (=MINIFS(A2:A6,B2:B6,"X")), calculates the result, saves and reloads the file, recalculates, and verifies that the minimum value for category "X" is 8 both before and after persistence.
    class Program
    {
        static void Main()
        {
            // -------------------------------------------------
            // 1. Create a new workbook (lifecycle rule: create)
            // -------------------------------------------------
            Workbook workbook = new Workbook();

            // -------------------------------------------------
            // 2. Set workbook to Excel 2016 compatibility mode.
            //    MINIFS is supported from Excel 2016 onward, so we
            //    disable older‑version compatibility checking.
            // -------------------------------------------------
            workbook.Settings.CheckCompatibility = false; // ensures newer functions are retained

            // -------------------------------------------------
            // 3. Populate sample data.
            //    Column A : numeric values
            //    Column B : criteria (text)
            // -------------------------------------------------
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Header row
            cells["A1"].PutValue("Value");
            cells["B1"].PutValue("Category");

            // Data rows
            double[] values = { 10, 5, 8, 3, 12 };
            string[] categories = { "X", "Y", "X", "Y", "X" };

            for (int i = 0; i < values.Length; i++)
            {
                cells[i + 1, 0].PutValue(values[i]);      // A2:A6
                cells[i + 1, 1].PutValue(categories[i]); // B2:B6
            }

            // -------------------------------------------------
            // 4. Insert MINIFS formula.
            //    =MINIFS(A2:A6, B2:B6, "X")
            // -------------------------------------------------
            cells["D1"].Formula = "=MINIFS(A2:A6,B2:B6,\"X\")";

            // -------------------------------------------------
            // 5. Calculate formulas.
            // -------------------------------------------------
            workbook.CalculateFormula();

            // -------------------------------------------------
            // 6. Retrieve and display the result.
            // -------------------------------------------------
            double result = cells["D1"].DoubleValue;
            Console.WriteLine($"MINIFS result (expected 8): {result}");

            // -------------------------------------------------
            // 7. Save the workbook (lifecycle rule: save)
            // -------------------------------------------------
            string filePath = "MinifsTest.xlsx";
            workbook.Save(filePath, SaveFormat.Xlsx);

            // -------------------------------------------------
            // 8. Load the workbook back (lifecycle rule: load)
            // -------------------------------------------------
            Workbook loadedWorkbook = new Workbook(filePath);

            // Ensure the same compatibility setting is applied after load
            loadedWorkbook.Settings.CheckCompatibility = false;

            // Recalculate to guarantee formula evaluation after loading
            loadedWorkbook.CalculateFormula();

            // Retrieve the result from the loaded workbook
            double loadedResult = loadedWorkbook.Worksheets[0].Cells["D1"].DoubleValue;
            Console.WriteLine($"MINIFS result after reload (expected 8): {loadedResult}");

            // -------------------------------------------------
            // 9. Simple validation
            // -------------------------------------------------
            if (Math.Abs(result - 8) < 0.0001 && Math.Abs(loadedResult - 8) < 0.0001)
            {
                Console.WriteLine("MINIFS function validated successfully.");
            }
            else
            {
                Console.WriteLine("MINIFS validation failed.");
            }
        }
    }
}
