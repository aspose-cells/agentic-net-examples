using System;
using Aspose.Cells;

namespace AsposeCellsTemplateRowExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet which will serve as the template sheet
            Worksheet templateSheet = workbook.Worksheets[0];
            Cells templateCells = templateSheet.Cells;

            // Populate sample data in column A (A1:A5)
            for (int i = 0; i < 5; i++)
            {
                templateCells[i, 0].PutValue(i + 1); // Values 1,2,3,4,5
            }

            // Set a shared formula in column B starting from B1 for 5 rows:
            // Each cell in B will be double the value in the same row of column A
            // This creates a template row (row 0) with a formula that will be copied later
            templateCells["B1"].SetSharedFormula("=A1*2", 5, 1);

            // Add two additional worksheets where the template row will be copied
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            Worksheet sheet3 = workbook.Worksheets.Add("Sheet3");

            // Copy the first row (row index 0) from the template sheet to each new sheet
            // The CopyRows method copies data, formats, and formulas
            sheet2.Cells.CopyRows(templateCells, 0, 0, 1);
            sheet3.Cells.CopyRows(templateCells, 0, 0, 1);

            // After copying, also copy the sample data in column A to the new sheets
            // (so the formulas have proper referenced values)
            for (int i = 0; i < 5; i++)
            {
                sheet2.Cells[i, 0].PutValue(i + 1);
                sheet3.Cells[i, 0].PutValue(i + 1);
            }

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Validate and display the calculated results for the formula cell (B1) in each sheet
            Console.WriteLine("Template Sheet B1 value: " + templateCells["B1"].Value); // Expected 2
            Console.WriteLine("Sheet2 B1 value: " + sheet2.Cells["B1"].Value);          // Expected 2
            Console.WriteLine("Sheet3 B1 value: " + sheet3.Cells["B1"].Value);          // Expected 2

            // Save the workbook (lifecycle save)
            workbook.Save("TemplateRowCopyResult.xlsx");
        }
    }
}