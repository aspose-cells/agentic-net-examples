using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (XLSX format by default)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet worksheet = workbook.Worksheets[0];
            Cells cells = worksheet.Cells;

            // Populate some data in the range A1:B2
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);
            cells["B1"].PutValue(30);
            cells["B2"].PutValue(40);

            // Create a named range covering A1:B2
            AsposeRange dataRange = cells.CreateRange("A1", "B2");
            dataRange.Name = "MyDataRange";

            // Use the named range in a formula (sum of all cells in the range)
            cells["C1"].Formula = "=SUM(MyDataRange)";

            // Calculate the formula so that the result is stored in C1
            workbook.CalculateFormula();

            // Output the calculated value to the console (optional verification)
            Console.WriteLine("Sum of MyDataRange: " + cells["C1"].IntValue);

            // Save the workbook as an XLSX file
            workbook.Save("ConfiguredRanges.xlsx", SaveFormat.Xlsx);
        }
    }
}