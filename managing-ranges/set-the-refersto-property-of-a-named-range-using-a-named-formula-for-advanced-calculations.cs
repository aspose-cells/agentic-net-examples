using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeRefersToDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate some sample data in column A (A1:A5)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1); // Values 1,2,3,4,5
            }

            // -----------------------------------------------------------------
            // 1. Define a named formula "MyFormula" that calculates the sum of A1:A5
            // -----------------------------------------------------------------
            int formulaNameIndex = workbook.Worksheets.Names.Add("MyFormula");
            // RefersTo must start with '=' and can contain any valid formula
            workbook.Worksheets.Names[formulaNameIndex].RefersTo = "=SUM(Sheet1!$A$1:$A$5)";

            // -----------------------------------------------------------------
            // 2. Define a named range "MyRange" whose RefersTo points to the named formula
            //    This demonstrates using a named formula as the reference of another name.
            // -----------------------------------------------------------------
            int rangeNameIndex = workbook.Worksheets.Names.Add("MyRange");
            // The RefersTo property can reference another name directly
            workbook.Worksheets.Names[rangeNameIndex].RefersTo = "=MyFormula";

            // -----------------------------------------------------------------
            // 3. Use the named range in a worksheet formula
            // -----------------------------------------------------------------
            // Cell B1 will display the result of the named formula via the named range
            sheet.Cells["B1"].Formula = "=MyRange";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the calculated value to the console
            Console.WriteLine("Result of MyRange (which refers to MyFormula): " + sheet.Cells["B1"].Value);

            // Save the workbook to a file
            workbook.Save("NamedRangeRefersToDemo.xlsx");
        }
    }
}