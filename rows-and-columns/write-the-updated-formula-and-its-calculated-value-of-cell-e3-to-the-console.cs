using System;
using Aspose.Cells;

namespace AsposeCellsFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some cells that will be used in the formula
            cells["A1"].PutValue(10);
            cells["B1"].PutValue(20);
            cells["C1"].PutValue(30);

            // Set a formula in cell E3 that sums A1, B1 and C1
            // The formula string must start with '=' and use commas as delimiters
            cells["E3"].Formula = "=SUM(A1, B1, C1)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve the updated formula and its calculated value
            string updatedFormula = cells["E3"].Formula;
            object calculatedValue = cells["E3"].Value;

            // Output the results to the console
            Console.WriteLine("Updated formula in E3: " + updatedFormula);
            Console.WriteLine("Calculated value in E3: " + calculatedValue);
        }
    }
}