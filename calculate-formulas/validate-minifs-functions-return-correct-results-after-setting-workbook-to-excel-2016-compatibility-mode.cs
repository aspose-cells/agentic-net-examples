// Title: Validate MINIFS formula results in Excel 2016 compatibility mode with Aspose.Cells for .NET
// AI Prompts: Write C# code using Aspose.Cells to set the workbook compliance to Excel 2016, insert a MINIFS formula with a numeric criteria range, calculate the sheet, and assert that the result equals the expected minimum value. | Demonstrate how to save a workbook that contains a MINIFS formula, reload it with Aspose.Cells, recalculate, and confirm that the formula result is unchanged.
// Common Searches: how to use MINIFS with Aspose.Cells in C# when workbook is set to Excel 2016 compatibility | Aspose.Cells calculate MINIFS after saving and reopening workbook | C# set OoxmlCompliance to Ecma376_2006 for conditional minimum function | verify MINIFS formula persistence in .NET using Aspose.Cells | testing MINIFS calculation accuracy in Excel 2016 mode with Aspose.Cells
// Tags: Aspose.Cells MINIFS support | Excel 2016 OoxmlCompliance setting | C# formula evaluation after save | conditional minimum verification .NET | recalculate loaded workbook Aspose.Cells

using System;
using Aspose.Cells;

namespace AsposeCellsMinifsValidation
{
    // The example creates a workbook, configures it for Excel 2016 compatibility, populates data, applies a MINIFS formula with a >10 criterion, calculates and checks the result, saves the file, reloads it, recalculates, and confirms that the persisted formula still returns the expected minimum value.
    class Program
    {
        static void Main()
        {
            // ---------- Create a new workbook ----------
            Workbook workbook = new Workbook(); // create

            // Set Excel 2016 compatibility mode (use OOXML compliance as a proxy)
            workbook.Settings.Compliance = OoxmlCompliance.Ecma376_2006; // Excel 2016 compatible

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            // Values to evaluate
            cells["A2"].PutValue(5);
            cells["A3"].PutValue(12);
            cells["A4"].PutValue(8);
            cells["A5"].PutValue(20);

            // Criteria range
            cells["B2"].PutValue(7);
            cells["B3"].PutValue(15);
            cells["B4"].PutValue(9);
            cells["B5"].PutValue(25);

            // MINIFS formula: find minimum in A2:A5 where corresponding B2:B5 > 10
            cells["C2"].Formula = "=MINIFS(A2:A5,B2:B5,\">10\")";

            // Calculate formulas
            workbook.CalculateFormula();

            // Expected result: minimum of {12,20} = 12
            double expected = 12;
            double actual = cells["C2"].DoubleValue;

            Console.WriteLine($"MINIFS result: {actual} (expected: {expected})");
            Console.WriteLine($"Validation {(Math.Abs(actual - expected) < 1e-9 ? "passed" : "failed")}");

            // ---------- Save the workbook ----------
            string filePath = "MinifsValidation.xlsx";
            workbook.Save(filePath); // save

            // ---------- Load the workbook ----------
            Workbook loadedWorkbook = new Workbook(filePath); // load
            Worksheet loadedSheet = loadedWorkbook.Worksheets[0];
            Cells loadedCells = loadedSheet.Cells;

            // Recalculate after load to ensure formula persists correctly
            loadedWorkbook.CalculateFormula();

            double loadedResult = loadedCells["C2"].DoubleValue;
            Console.WriteLine($"Loaded MINIFS result: {loadedResult} (expected: {expected})");
            Console.WriteLine($"Loaded validation {(Math.Abs(loadedResult - expected) < 1e-9 ? "passed" : "failed")}");
        }
    }
}
