using System;
using Aspose.Cells;

namespace AsposeCellsNamedRangeVLookup
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // -------------------- Sheet1: data and named range --------------------
            Worksheet sheet1 = workbook.Worksheets[0];
            sheet1.Name = "Data";

            // Populate lookup table (columns B and C)
            // Header
            sheet1.Cells["B1"].PutValue("Key");
            sheet1.Cells["C1"].PutValue("Value");

            // Sample data
            sheet1.Cells["B2"].PutValue("A");
            sheet1.Cells["C2"].PutValue(100);
            sheet1.Cells["B3"].PutValue("B");
            sheet1.Cells["C3"].PutValue(200);
            sheet1.Cells["B4"].PutValue("C");
            sheet1.Cells["C4"].PutValue(300);
            sheet1.Cells["B5"].PutValue("D");
            sheet1.Cells["C5"].PutValue(400);

            // Create a named range that refers to the whole columns B:C
            int nameIndex = workbook.Worksheets.Names.Add("LookupTable");
            Name lookupName = workbook.Worksheets.Names[nameIndex];
            // Reference whole columns B and C on Sheet1
            lookupName.RefersTo = "=Data!$B:$C";

            // -------------------- Sheet2: VLOOKUP usage --------------------
            Worksheet sheet2 = workbook.Worksheets.Add("Lookup");
            // Value to look up
            sheet2.Cells["A1"].PutValue("C");

            // VLOOKUP formula using the named range defined above
            // =VLOOKUP(A1, LookupTable, 2, FALSE)
            sheet2.Cells["B1"].Formula = "=VLOOKUP(A1, LookupTable, 2, FALSE)";

            // Calculate formulas so that the result is available
            workbook.CalculateFormula();

            // Output the result to console
            Console.WriteLine("Lookup result for key '{0}' is: {1}",
                sheet2.Cells["A1"].StringValue,
                sheet2.Cells["B1"].Value);

            // Save the workbook
            workbook.Save("NamedRangeVLookupDemo.xlsx");
        }
    }
}