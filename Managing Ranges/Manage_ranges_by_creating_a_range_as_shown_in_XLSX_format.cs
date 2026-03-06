using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet's cells collection
            Workbook workbook = new Workbook();
            Cells cells = workbook.Worksheets[0].Cells;

            // Create a range that covers cells A1 to A2
            AsposeRange myRange = cells.CreateRange("A1", "A2");
            // Assign a name to the range so it can be referenced in formulas
            myRange.Name = "MyRange";

            // Populate the range with sample numeric values
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);

            // Use the named range in a formula placed in cell B1
            cells["B1"].Formula = "=SUM(MyRange)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the result of the formula to the console
            Console.WriteLine("Sum of MyRange (A1:A2) is: " + cells["B1"].IntValue);

            // Save the workbook to an XLSX file
            workbook.Save("RangeDemo.xlsx");
        }
    }
}