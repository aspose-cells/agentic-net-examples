using System;
using Aspose.Cells;

namespace DynamicNamedRangeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Data";

            // Populate some sample data in column A (non‑empty rows)
            sheet.Cells["A1"].PutValue("Header");
            sheet.Cells["A2"].PutValue(10);
            sheet.Cells["A3"].PutValue(20);
            sheet.Cells["A4"].PutValue(30);
            // Add more rows as needed...

            // Define a dynamic named range that expands based on non‑empty rows in column A
            // OFFSET(start, rows, cols, height, width) with COUNTA to count non‑empty cells
            int nameIndex = workbook.Worksheets.Names.Add("DynamicRange");
            Name dynamicName = workbook.Worksheets.Names[nameIndex];
            dynamicName.RefersTo = $"=OFFSET({sheet.Name}!$A$1,0,0,COUNTA({sheet.Name}!$A:$A),1)";

            // Optional: calculate formulas so the named range is evaluated immediately
            workbook.CalculateFormula();

            // Save the workbook (lifecycle rule: save)
            workbook.Save("DynamicRangeDemo.xlsx");
        }
    }
}