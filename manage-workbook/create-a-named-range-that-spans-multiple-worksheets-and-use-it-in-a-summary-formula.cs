using System;
using Aspose.Cells;

namespace AsposeCellsMultiSheetNamedRange
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Add two worksheets (Sheet1 already exists)
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Populate some sample data in both sheets
            // Sheet1: A1:A3 = 10, 20, 30
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet1.Cells["A3"].PutValue(30);

            // Sheet2: A1:A3 = 5, 15, 25
            sheet2.Cells["A1"].PutValue(5);
            sheet2.Cells["A2"].PutValue(15);
            sheet2.Cells["A3"].PutValue(25);

            // Create a named range that spans both worksheets
            // The RefersTo string uses a comma to separate the areas
            int nameIndex = workbook.Worksheets.Names.Add("MultiRange");
            Name multiRange = workbook.Worksheets.Names[nameIndex];
            multiRange.RefersTo = "=Sheet1!$A$1:$A$3,Sheet2!$A$1:$A$3";

            // Use the named range in a summary formula (sum of all cells in the range)
            // Place the formula in Sheet1!B1
            sheet1.Cells["B1"].Formula = "=SUM(MultiRange)";

            // Calculate formulas so that the result is available
            workbook.CalculateFormula();

            // Output the calculated result to console
            Console.WriteLine("Sum of MultiRange (Sheet1!A1:A3 + Sheet2!A1:A3) = " + sheet1.Cells["B1"].Value);

            // Save the workbook
            workbook.Save("MultiSheetNamedRangeDemo.xlsx");
        }
    }
}