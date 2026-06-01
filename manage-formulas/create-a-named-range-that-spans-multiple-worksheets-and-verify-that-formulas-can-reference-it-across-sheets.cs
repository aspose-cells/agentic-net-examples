using System;
using Aspose.Cells;

namespace AsposeCellsMultiSheetNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (contains one default worksheet named "Sheet1")
            Workbook workbook = new Workbook();

            // Access the default worksheet and add data
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Sheet1";
            sheet1.Cells["A1"].PutValue(1);
            sheet1.Cells["A2"].PutValue(2);

            // Add a second worksheet and add data
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");
            sheet2.Cells["B1"].PutValue(3);
            sheet2.Cells["B2"].PutValue(4);

            // Add a third worksheet where the formula will be placed
            Worksheet sheet3 = workbook.Worksheets.Add("Summary");

            // Create a named range that refers to ranges on both Sheet1 and Sheet2
            // The RefersTo string uses a comma to separate multiple areas
            int nameIndex = workbook.Worksheets.Names.Add("MultiRange");
            Name multiRangeName = workbook.Worksheets.Names[nameIndex];
            multiRangeName.RefersTo = "=Sheet1!$A$1:$A$2,Sheet2!$B$1:$B$2";

            // Use the named range in a formula on the third sheet
            // SUM will add all numbers from both referenced areas
            sheet3.Cells["A1"].Formula = "=SUM(MultiRange)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Retrieve and display the result of the formula
            double result = sheet3.Cells["A1"].DoubleValue;
            Console.WriteLine("Result of SUM(MultiRange) on 'Summary' sheet: " + result);
            // Expected result: 1 + 2 + 3 + 4 = 10

            // Save the workbook to verify the named range and formula persist
            workbook.Save("MultiSheetNamedRangeDemo.xlsx");
        }
    }
}