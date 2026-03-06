using System;
using Aspose.Cells;
using AsposeRange = Aspose.Cells.Range;

namespace AsposeCellsRangeDemo
{
    class Program
    {
        static void Main()
        {
            ManageRanges.Run();
        }
    }

    public class ManageRanges
    {
        public static void Run()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the default worksheet's cells
            Cells cells = workbook.Worksheets[0].Cells;

            // Create a range from A1 to A2
            AsposeRange myRange = cells.CreateRange("A1", "A2");
            myRange.Name = "MyRange";

            // Populate the range with sample values
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);

            // Use the named range in a formula (SUM of the range)
            cells["B1"].Formula = "=SUM(MyRange)";

            // Calculate the formula to obtain the result
            workbook.CalculateFormula();

            // Output the calculated sum to the console
            Console.WriteLine("Sum of MyRange (A1:A2) is: " + cells["B1"].IntValue);

            // Save the workbook as an XLSX file
            workbook.Save("ManagedRanges.xlsx");
        }
    }
}