using System;
using Aspose.Cells;

namespace AsposeCellsNamedFormulaDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet (default name is "Sheet1")
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data in column A
            for (int i = 0; i < 10; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1); // A1:A10 = 1..10
            }

            // -----------------------------------------------------------------
            // 1. Create a named formula "TotalSum" that calculates the sum of A1:A10
            // -----------------------------------------------------------------
            int totalSumIndex = sheet.Workbook.Worksheets.Names.Add("TotalSum");
            Name totalSumName = sheet.Workbook.Worksheets.Names[totalSumIndex];
            // RefersTo can contain a formula that returns a value
            totalSumName.RefersTo = "=SUM(Sheet1!$A$1:$A$10)";

            // -----------------------------------------------------------------
            // 2. Create another named range "SumRange" that refers to the result
            //    of the named formula "TotalSum". This demonstrates using a named
            //    formula inside the RefersTo property of another name.
            // -----------------------------------------------------------------
            int sumRangeIndex = sheet.Workbook.Worksheets.Names.Add("SumRange");
            Name sumRangeName = sheet.Workbook.Worksheets.Names[sumRangeIndex];
            // The RefersTo property can reference another name by using its name
            sumRangeName.RefersTo = "=TotalSum";

            // -----------------------------------------------------------------
            // 3. Use the "SumRange" name in a worksheet formula
            // -----------------------------------------------------------------
            // Place the result in cell B1
            sheet.Cells["B1"].Formula = "=SumRange";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Display the calculated value of B1 (should be 55)
            Console.WriteLine("Calculated value of B1 (SumRange): " + sheet.Cells["B1"].Value);

            // -----------------------------------------------------------------
            // 4. Save the workbook to verify the named formulas are persisted
            // -----------------------------------------------------------------
            workbook.Save("NamedFormulaDemo.xlsx");
        }
    }
}