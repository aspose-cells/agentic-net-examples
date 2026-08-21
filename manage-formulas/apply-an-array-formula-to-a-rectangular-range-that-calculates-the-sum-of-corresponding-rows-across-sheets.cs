// Title: Set a 3‑D array formula to sum matching cells across worksheets with Aspose.Cells for .NET (C#)
// Description: Creates a workbook with three sheets, populates Sheet1 and Sheet2 with sample numbers, and uses SetArrayFormula to place a 5 × 3 array formula "=SUM(Sheet1:Sheet2!A1)" on Sheet3. The formula automatically spills over the target range, CalculateFormula materializes the results, and the workbook is saved as an XLSX file.
// Keywords: Aspose.Cells | C# | .NET | SetArrayFormula | 3‑D reference | array formula across sheets | SUM across worksheets | CalculateFormula | save workbook | Excel automation
// Common Searches: Aspose.Cells set array formula across multiple sheets | C# 3‑D SUM formula with Aspose.Cells | How to apply SetArrayFormula to a rectangular range in Aspose.Cells | Spill array formula over a range using Aspose.Cells .NET | Calculate formulas after setting array formula Aspose.Cells
// Developer Intent: Apply a rectangular 3‑D array formula that sums corresponding cells from Sheet1 and Sheet2 into Sheet3 using Aspose.Cells for .NET.
// Use Cases: Consolidate daily sales figures stored on separate worksheets into a single summary sheet with one array formula. | Combine budget numbers from two department sheets into a merged report without writing individual cell formulas. | Create a rolling financial statement that adds values from multiple period sheets using a single 3‑D array formula.
// AI Prompts: Show how to extend the array formula to include a third worksheet in the 3‑D SUM reference. | Explain how to read values from the spilled range after workbook.CalculateFormula() is called. | Generate code that determines the array size dynamically based on the source sheets' row and column counts.

using System;
using Aspose.Cells;

namespace AsposeCellsArrayFormulaDemo
{
    // Creates a workbook with three sheets, populates Sheet1 and Sheet2 with sample numbers, and uses SetArrayFormula to place a 5 × 3 array formula "=SUM(Sheet1:Sheet2!A1)" on Sheet3. The formula automatically spills over the target range, CalculateFormula materializes the results, and the workbook is saved as an XLSX file.
    class Program
    {
        static void Main()
        {
            // Create a new workbook with three worksheets
            Workbook workbook = new Workbook();
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Populate sample data in Sheet1 (A1:C5)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet1.Cells[row, col].PutValue((row + 1) * 10 + col); // e.g., 11,12,13,...
                }
            }

            // Populate sample data in Sheet2 (A1:C5)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 3; col++)
                {
                    sheet2.Cells[row, col].PutValue((row + 1) * 100 + col); // e.g., 101,102,103,...
                }
            }

            // Set an array formula in Sheet3 that sums the corresponding cells across Sheet1 and Sheet2
            // The 3‑D reference SUM(Sheet1:Sheet2!A1) will adjust for each cell in the spilled range
            Cell targetCell = sheet3.Cells["A1"];
            targetCell.SetArrayFormula("=SUM(Sheet1:Sheet2!A1)", 5, 3);

            // Calculate all formulas so that the array results are materialized
            workbook.CalculateFormula();

            // Optional: display a few results to verify
            Console.WriteLine("Sheet3!A1 = " + sheet3.Cells["A1"].Value); // should be Sheet1!A1 + Sheet2!A1
            Console.WriteLine("Sheet3!C5 = " + sheet3.Cells["C5"].Value); // should be Sheet1!C5 + Sheet2!C5

            // Save the workbook
            workbook.Save("ArrayFormulaDemo.xlsx");
        }
    }
}
