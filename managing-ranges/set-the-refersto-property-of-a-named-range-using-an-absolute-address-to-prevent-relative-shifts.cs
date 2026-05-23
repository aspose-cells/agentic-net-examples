using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            sheet.Name = "Sheet1";

            // Populate some sample data (optional, just for demonstration)
            sheet.Cells["A1"].PutValue(10);
            sheet.Cells["A2"].PutValue(20);
            sheet.Cells["A3"].PutValue(30);

            // Add a named range to the workbook's name collection
            int nameIndex = sheet.Workbook.Worksheets.Names.Add("MyAbsoluteRange");
            Name namedRange = sheet.Workbook.Worksheets.Names[nameIndex];

            // Set the RefersTo property using an absolute address.
            // The address is absolute ($A$1:$A$3) and the formula starts with '='.
            // isR1C1 = false (A1 style), isLocal = false (non‑locale formatted).
            namedRange.SetRefersTo("=Sheet1!$A$1:$A$3", false, false);

            // Verify the RefersTo value (optional)
            Console.WriteLine($"Named range refers to: {namedRange.RefersTo}");

            // Use the named range in a formula to ensure it works
            sheet.Cells["B1"].Formula = "=SUM(MyAbsoluteRange)";
            workbook.CalculateFormula();
            Console.WriteLine($"Result of SUM(MyAbsoluteRange): {sheet.Cells["B1"].Value}");

            // Save the workbook
            workbook.Save("NamedRangeAbsoluteAddress.xlsx");
        }
    }
}