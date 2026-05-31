using System;
using Aspose.Cells;

namespace NamedRangeSumExample
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook
                Workbook workbook = new Workbook();

                // Access the first worksheet
                Worksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Sheet1";

                // Populate some numeric data in column A (A1:A3)
                sheet.Cells["A1"].PutValue(10);
                sheet.Cells["A2"].PutValue(20);
                sheet.Cells["A3"].PutValue(30);

                // Add a workbook‑scoped named range called "MyRange"
                int nameIndex = workbook.Worksheets.Names.Add("MyRange");
                Name myRange = workbook.Worksheets.Names[nameIndex];
                // Define the range reference (absolute A1:A3 on Sheet1)
                myRange.RefersTo = "=Sheet1!$A$1:$A$3";

                // Use the named range in a formula to calculate its sum
                sheet.Cells["B1"].Formula = "=SUM(MyRange)";

                // Calculate all formulas in the workbook
                workbook.CalculateFormula();

                // Output the result
                Console.WriteLine("Sum of MyRange: " + sheet.Cells["B1"].Value);

                // Save the workbook (optional)
                workbook.Save("NamedRangeSum.xlsx");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}