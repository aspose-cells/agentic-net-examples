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

            // Define a range from A1 to A3 using the CreateRange(string, string) method
            AsposeRange myRange = cells.CreateRange("A1", "A3");
            // Optionally give the range a name for later reference in formulas
            myRange.Name = "MyDataRange";

            // Populate the range with sample values
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(15);
            cells["A3"].PutValue(25);

            // Use the defined range in a formula (sum of the range)
            cells["B1"].Formula = "=SUM(MyDataRange)";

            // Calculate the formula so that the result is stored in B1
            workbook.CalculateFormula();

            // Output the calculated sum to the console
            Console.WriteLine("Sum of MyDataRange (A1:A3) = " + cells["B1"].IntValue);

            // Save the workbook as an XLSX file
            workbook.Save("RangeDemo.xlsx");
        }
    }
}