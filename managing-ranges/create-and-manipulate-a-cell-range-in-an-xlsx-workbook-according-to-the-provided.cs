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

            // Create a range from A1 to A2 using the CreateRange(string, string) method
            AsposeRange myRange = cells.CreateRange("A1", "A2");
            // Assign a name to the range (optional but useful for formulas)
            myRange.Name = "MyRange";

            // Populate the cells within the range
            cells["A1"].PutValue(10);
            cells["A2"].PutValue(20);

            // Use the named range in a formula placed in B1
            cells["B1"].Formula = "=SUM(MyRange)";

            // Calculate all formulas in the workbook
            workbook.CalculateFormula();

            // Output the result of the formula to the console
            Console.WriteLine("Sum of MyRange (A1:A2) is: " + cells["B1"].IntValue);

            // Demonstrate additional range manipulation:
            // 1. Clear the contents of the range
            myRange.ClearContents();

            // 2. Re‑populate the range with new values
            cells["A1"].PutValue(5);
            cells["A2"].PutValue(15);

            // 3. Insert a new column before column A and shift the range right
            // Define the area to insert (the whole column A)
            CellArea insertArea = CellArea.CreateCellArea(0, 0, cells.MaxDataRow, 0);
            cells.InsertRange(insertArea, ShiftType.Right);

            // After insertion, the original range has moved to B1:B2
            // Update the formula to reference the moved range
            cells["C1"].Formula = "=SUM(B1:B2)";
            workbook.CalculateFormula();
            Console.WriteLine("After inserting a column, sum is: " + cells["C1"].IntValue);

            // Save the workbook as an XLSX file
            workbook.Save("RangeManipulationDemo.xlsx");
        }
    }
}