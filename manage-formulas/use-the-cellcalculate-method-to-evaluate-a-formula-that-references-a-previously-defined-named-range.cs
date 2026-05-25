using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeCalculation
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate some data that will be part of the named range
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Define a named range that refers to the cells A1:A3
            int nameIndex = workbook.Worksheets.Names.Add("MyRange");
            workbook.Worksheets.Names[nameIndex].RefersTo = "=Sheet1!$A$1:$A$3";

            // Set a formula in another cell that uses the named range
            Cell formulaCell = sheet.Cells["B1"];
            formulaCell.Formula = "=SUM(MyRange)";

            // Calculate the formula using the Cell.Calculate method
            formulaCell.Calculate(new CalculationOptions());

            // Output the calculated result
            Console.WriteLine("Result of SUM(MyRange): " + formulaCell.Value);

            // Save the workbook (lifecycle rule)
            workbook.Save("NamedRangeCalculation.xlsx");
        }
    }
}