// Title: Copy rows and recalculate formulas with Aspose.Cells for .NET
// Description: Shows how to create a workbook, add numeric data and a SUM formula, duplicate rows using Cells.CopyRows, run Workbook.CalculateFormula to update dependent formulas, and save the result.
// Keywords: Aspose.Cells | CopyRows | CalculateFormula | C# | .NET | recalculate formulas | duplicate rows | Excel automation | update formulas after copy
// Common Searches: Aspose.Cells copy rows C# example | recalculate formulas after copying rows Aspose.Cells | Workbook.CalculateFormula usage | how to duplicate rows with formulas in Aspose.Cells | CopyRows method Aspose.Cells tutorial
// Developer Intent: Refresh all formulas after rows are duplicated so calculations remain accurate.
// Use Cases: Copy a summary row containing formulas to another section of the sheet and automatically update totals. | Programmatically replicate a data table with embedded calculations and ensure the new copy reflects correct results. | Build a template that inserts repeated row blocks and uses CalculateFormula to keep aggregate values consistent.
// AI Prompts: Provide C# code that copies rows with Aspose.Cells and then calls CalculateFormula to refresh dependent formulas. | Show an example of using Cells.CopyRows followed by Workbook.CalculateFormula to keep SUM formulas correct after duplication. | Explain the steps required to ensure formulas recalculate automatically after copying rows in Aspose.Cells for .NET.

using System;
using Aspose.Cells;

namespace AsposeCellsExamples
{
    // Shows how to create a workbook, add numeric data and a SUM formula, duplicate rows using Cells.CopyRows, run Workbook.CalculateFormula to update dependent formulas, and save the result.
    public class CopyRowsAndRecalculateDemo
    {
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
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data: A1:A3 numbers, B1 formula summing A1:A3
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["A3"].PutValue(30);
            cells["B1"].Formula = "=SUM(A1:A3)";

            // Calculate formulas so B1 gets the correct value
            workbook.CalculateFormula();

            Console.WriteLine("Before copying rows:");
            Console.WriteLine($"B1 = {cells["B1"].Value}"); // Expected 60

            // Copy rows 0‑2 (first three rows) to start at row index 5 (Excel row 6)
            cells.CopyRows(cells, 0, 5, 3);

            // Recalculate after copying to update formulas in the new rows
            workbook.CalculateFormula();

            Console.WriteLine("After copying rows and recalculation:");
            Console.WriteLine($"B1 = {cells["B1"].Value}"); // Should remain 60
            Console.WriteLine($"B6 = {cells["B6"].Value}"); // Should also be 60 (sum of A6:A8)

            // Save the workbook
            string outputPath = "CopyRowsRecalculateDemo.xlsx";
            workbook.Save(outputPath);
            Console.WriteLine($"Workbook saved to {outputPath}");
        }
    }
}
