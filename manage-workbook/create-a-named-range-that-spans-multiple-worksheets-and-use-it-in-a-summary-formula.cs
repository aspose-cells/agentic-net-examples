using System;
using Aspose.Cells;

namespace AsposeCellsMultiSheetNamedRange
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (contains one default worksheet named "Sheet1")
            Workbook workbook = new Workbook();

            // Add a second worksheet named "Sheet2"
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Populate sample data in Sheet1 (A1:B2)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet1.Cells["B1"].PutValue(30);
            sheet1.Cells["B2"].PutValue(40);

            // Populate sample data in Sheet2 (A1:B2)
            sheet2.Cells["A1"].PutValue(5);
            sheet2.Cells["A2"].PutValue(15);
            sheet2.Cells["B1"].PutValue(25);
            sheet2.Cells["B2"].PutValue(35);

            // Create a named range that refers to ranges on both worksheets
            // The RefersTo string uses a comma to separate multiple areas
            int nameIndex = workbook.Worksheets.Names.Add("MultiSheetRange");
            Name multiRange = workbook.Worksheets.Names[nameIndex];
            multiRange.RefersTo = "=Sheet1!$A$1:$B$2,Sheet2!$A$1:$B$2";

            // Use the named range in a summary formula (sum of all cells in the multi‑sheet range)
            // Place the formula in Sheet1 cell C1
            sheet1.Cells["C1"].Formula = "=SUM(MultiSheetRange)";

            // Calculate formulas so that the result is stored in C1
            workbook.CalculateFormula();

            // Output the calculated result to the console (optional verification)
            Console.WriteLine("Sum of MultiSheetRange: " + sheet1.Cells["C1"].Value);

            // Save the workbook
            workbook.Save("MultiSheetNamedRangeDemo.xlsx");
        }
    }
}