using System;
using Aspose.Cells;

namespace AsposeCellsDynamicNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate initial data in the first row (A1:C1)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["B1"].PutValue(20);
            sheet.Cells["C1"].PutValue(30);

            // Create a named range that expands horizontally.
            // The OFFSET formula starts at A1, height = 1 row,
            // width = number of non‑empty cells in the first row (COUNTA).
            int nameIndex = workbook.Worksheets.Names.Add("MyDynamicRange");
            Name dynamicRange = workbook.Worksheets.Names[nameIndex];
            dynamicRange.RefersTo = "=OFFSET(Data!$A$1,0,0,1,COUNTA(Data!$1:$1))";

            // Use the named range in a formula (sum of the range)
            sheet.Cells["D1"].Formula = "=SUM(MyDynamicRange)";

            // Calculate the formula
            workbook.CalculateFormula();

            Console.WriteLine("Initial sum (A1:C1): " + sheet.Cells["D1"].Value); // Expected 60

            // Insert a new column to the right of the existing data (after column C)
            sheet.Cells.InsertColumn(3); // Inserts column D, shifting existing columns right

            // Add a new value in the newly inserted column (now column D)
            sheet.Cells["D1"].PutValue(40);

            // Re‑calculate to reflect the expanded range
            workbook.CalculateFormula();

            Console.WriteLine("Updated sum after adding column D: " + sheet.Cells["E1"].Value); // Sum now includes the new column

            // Save the workbook
            workbook.Save("DynamicNamedRangeDemo.xlsx");
        }
    }
}