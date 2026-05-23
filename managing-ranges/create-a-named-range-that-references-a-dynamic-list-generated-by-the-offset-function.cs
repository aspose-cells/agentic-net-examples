using System;
using Aspose.Cells;

namespace AsposeCellsDynamicNamedRange
{
    public class Program
    {
        public static void Main()
        {
            // Create a new workbook (lifecycle create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";
            Cells cells = sheet.Cells;

            // Populate a vertical list in column A (A1:A5)
            for (int i = 0; i < 5; i++)
            {
                cells[i, 0].PutValue($"Item {i + 1}");
            }

            // Define a named range "DynamicList" that uses OFFSET to refer to the list in column A.
            // The formula expands automatically as rows are added/removed in column A.
            // OFFSET(start, rows, cols, height, width)
            // Here: start = Sheet1!$A$1, rows = 0, cols = 0,
            // height = COUNTA(Sheet1!$A:$A) (number of non‑empty cells in column A),
            // width = 1 (single column)
            int nameIndex = workbook.Worksheets.Names.Add("DynamicList");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)";

            // Use the dynamic named range in a formula (e.g., count the items)
            cells["C1"].Formula = "=COUNTA(DynamicList)";

            // Calculate all formulas so that the result appears in C1
            workbook.CalculateFormula();

            // Output the calculated value to the console (optional verification)
            Console.WriteLine("Number of items in DynamicList: " + cells["C1"].IntValue);

            // Save the workbook (lifecycle save)
            workbook.Save("DynamicNamedRangeDemo.xlsx");
        }
    }
}