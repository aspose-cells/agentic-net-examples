using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule)
            Workbook workbook = new Workbook();

            // Access the default worksheet's cells collection
            Cells cells = workbook.Worksheets[0].Cells;

            // Create a range from A1 to A2 using the CreateRange(string, string) method (feature rule)
            AsposeRange myRange = cells.CreateRange("A1", "A2");
            myRange.Name = "MyRange";

            // Populate the range with sample values
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);

            // Use the named range in a formula (sum of the range)
            cells["B1"].Formula = "=SUM(MyRange)";

            // Calculate the formula to obtain the result
            workbook.CalculateFormula();

            // Output the calculated sum to the console
            Console.WriteLine("Sum of MyRange (A1:A2) = " + cells["B1"].IntValue);

            // Save the workbook in XLSX format (lifecycle rule)
            workbook.Save("ManagedRangesDemo.xlsx");
        }
    }
}