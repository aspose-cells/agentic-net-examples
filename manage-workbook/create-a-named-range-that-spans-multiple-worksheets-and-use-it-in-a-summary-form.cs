using System;
using Aspose.Cells;

namespace AsposeCellsMultiSheetNamedRange
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle: create)
            Workbook workbook = new Workbook();

            // Add two worksheets that will contain data
            Worksheet sheet1 = workbook.Worksheets[0];               // default Sheet1
            sheet1.Name = "Sheet1";
            Worksheet sheet2 = workbook.Worksheets.Add("Sheet2");

            // Populate some numeric data in both sheets (A1:A3)
            sheet1.Cells["A1"].PutValue(10);
            sheet1.Cells["A2"].PutValue(20);
            sheet1.Cells["A3"].PutValue(30);

            sheet2.Cells["A1"].PutValue(5);
            sheet2.Cells["A2"].PutValue(15);
            sheet2.Cells["A3"].PutValue(25);

            // Create a named range that spans the same range on both worksheets
            // The RefersTo string uses a comma to separate the two areas
            int nameIndex = workbook.Worksheets.Names.Add("MultiRange");
            Name multiRange = workbook.Worksheets.Names[nameIndex];
            multiRange.RefersTo = "=Sheet1!$A$1:$A$3,Sheet2!$A$1:$A$3";

            // Use the named range in a summary formula on a new sheet
            Worksheet summarySheet = workbook.Worksheets.Add("Summary");
            // Place the formula in B1: sum of all values in the multi‑sheet named range
            summarySheet.Cells["B1"].Formula = "=SUM(MultiRange)";

            // Calculate formulas (required before reading the result)
            workbook.CalculateFormula();

            // Output the calculated sum to console
            Console.WriteLine("Sum of MultiRange (Sheet1!A1:A3 + Sheet2!A1:A3) = " +
                              summarySheet.Cells["B1"].Value);

            // Save the workbook (lifecycle: save)
            workbook.Save("MultiSheetNamedRange.xlsx");
        }
    }
}