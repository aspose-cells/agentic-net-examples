using System;
using Aspose.Cells;

namespace NamedRangeUpdateDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate some sample data in column A
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);
            sheet.Cells["A4"].PutValue(40); // extra cell for later expansion

            // Create a named range "MyRange" that initially refers to A1:A3
            int nameIndex = sheet.Workbook.Worksheets.Names.Add("MyRange");
            Name myRange = sheet.Workbook.Worksheets.Names[nameIndex];
            myRange.RefersTo = "=Sheet1!$A$1:$A$3";

            // Use the named range in a formula (SUM) placed in B1
            sheet.Cells["B1"].Formula = "=SUM(MyRange)";

            // Calculate formulas and display the result (should be 10+20+30 = 60)
            workbook.CalculateFormula();
            Console.WriteLine($"Initial SUM(MyRange) = {sheet.Cells["B1"].Value}"); // Expected: 60

            // Change the named range to include A4 as well (A1:A4)
            myRange.RefersTo = "=Sheet1!$A$1:$A$4";

            // Recalculate formulas; Aspose.Cells automatically updates formulas that reference the name
            workbook.CalculateFormula();

            // Display the updated result (should be 10+20+30+40 = 100)
            Console.WriteLine($"Updated SUM(MyRange) after expanding range = {sheet.Cells["B1"].Value}"); // Expected: 100

            // Verify that the name is indeed referred by other formulas
            Console.WriteLine($"Is 'MyRange' referred by any formula? {myRange.IsReferred}");

            // Save the workbook (optional, demonstrates lifecycle rule usage)
            workbook.Save("NamedRangeUpdateDemo.xlsx");
        }
    }
}