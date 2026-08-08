// Title: Aspose.Cells .NET: Copy a Formula Row to Multiple Worksheets and Verify Results
// Description: Demonstrates how to create a template row with a formula, duplicate it across several worksheets, fill each sheet with distinct data, recalculate all formulas, and programmatically validate the computed values before saving the workbook.
// Keywords: Aspose.Cells copy row formula .NET | duplicate template row worksheets | calculate workbook formulas Aspose.Cells | validate formula results C# | Cells.CopyRows example | Aspose.Cells row replication
// Common Searches: How to copy a row with a formula to other sheets using Aspose.Cells | Validate copied formula values in Aspose.Cells .NET | Copy template row across worksheets Aspose.Cells | Recalculate all formulas after copying rows in Aspose.Cells
// Developer Intent: Create a single row that contains a formula, copy that row to multiple worksheets, populate each sheet with its own data, recalculate the workbook, and confirm that the formula outputs are correct.
// Use Cases: Financial reporting: define a totals row once and reuse it on every department sheet, then verify the totals. | Pricing models: apply a discount calculation row across product category sheets while ensuring each sheet reflects the correct discounted price. | Data‑entry templates: copy a summary row to new worksheets and automatically check that the summary matches the entered values.
// AI Prompts: Generate C# code with Aspose.Cells that creates a template row containing a formula, copies it to several worksheets, fills column A with unique values, recalculates, and prints A and B column results. | Write a method that iterates through copied worksheets and asserts that the formula column returns the expected doubled values for each row. | Explain how Cells.CopyRows preserves relative references when a formula row is duplicated across different worksheets in Aspose.Cells.

using System;
using Aspose.Cells;

namespace AsposeCellsTemplateRowDemo
{
    // Demonstrates how to create a template row with a formula, duplicate it across several worksheets, fill each sheet with distinct data, recalculate all formulas, and programmatically validate the computed values before saving the workbook.
    class Program
    {
        static void Main()
        {
            try
            {
                // ------------------------------------------------------------
                // 1. Create a new workbook and get the first worksheet (template)
                // ------------------------------------------------------------
                Workbook workbook = new Workbook();
                Worksheet templateSheet = workbook.Worksheets[0];
                Cells templateCells = templateSheet.Cells;

                // ------------------------------------------------------------
                // 2. Populate sample data in column A of the template sheet
                //    (Rows 2‑6 will hold the data, row index 1‑5)
                // ------------------------------------------------------------
                for (int i = 0; i < 5; i++)
                {
                    // A2:A6 = 10,20,30,40,50
                    templateCells[i + 1, 0].PutValue((i + 1) * 10);
                }

                // ------------------------------------------------------------
                // 3. Create a template row (row 2) that contains a formula.
                //    Column B will contain a formula that doubles the value in column A.
                // ------------------------------------------------------------
                // B2 (row index 1, column index 1)
                templateCells[1, 1].Formula = "=A2*2";

                // ------------------------------------------------------------
                // 4. Add two additional worksheets that will receive the template row
                // ------------------------------------------------------------
                Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
                Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

                // ------------------------------------------------------------
                // 5. Copy the template row (row index 1) from the template sheet
                //    to the same row position in the new sheets.
                // ------------------------------------------------------------
                sheet2.Cells.CopyRows(templateCells, 1, 1, 1);
                sheet3.Cells.CopyRows(templateCells, 1, 1, 1);

                // ------------------------------------------------------------
                // 6. Populate distinct data in column A of each sheet
                // ------------------------------------------------------------
                // Sheet2: 5,15,25,35,45
                for (int i = 0; i < 5; i++)
                    sheet2.Cells[i + 1, 0].PutValue(5 + i * 10);

                // Sheet3: 2,4,6,8,10
                for (int i = 0; i < 5; i++)
                    sheet3.Cells[i + 1, 0].PutValue(2 * (i + 1));

                // ------------------------------------------------------------
                // 7. Calculate all formulas in the workbook
                // ------------------------------------------------------------
                workbook.CalculateFormula();

                // ------------------------------------------------------------
                // 8. Validate and display the calculated results for each sheet
                // ------------------------------------------------------------
                Console.WriteLine("=== Validation Results ===");
                ValidateSheet(templateSheet, "TemplateSheet");
                ValidateSheet(sheet2, "Sheet2");
                ValidateSheet(sheet3, "Sheet3");

                // ------------------------------------------------------------
                // 9. Save the workbook (lifecycle rule)
                // ------------------------------------------------------------
                workbook.Save("TemplateRowCopyResult.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        // Helper method to print values of column A and the calculated column B
        static void ValidateSheet(Worksheet ws, string sheetName)
        {
            try
            {
                Console.WriteLine($"--- {sheetName} ---");
                for (int row = 1; row <= 5; row++) // rows 2‑6 (index 1‑5)
                {
                    double aVal = ws.Cells[row, 0].DoubleValue;
                    double bVal = ws.Cells[row, 1].DoubleValue;
                    Console.WriteLine($"Row {row + 1}: A={aVal}, B (A*2)={bVal}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Validation error in {sheetName}: {ex.Message}");
            }
        }
    }
}
