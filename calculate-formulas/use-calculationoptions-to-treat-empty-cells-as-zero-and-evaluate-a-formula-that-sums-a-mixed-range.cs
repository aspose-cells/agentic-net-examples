using System;
using Aspose.Cells;

namespace AsposeCellsDemo
{
    class Program
    {
        static void Main()
        {
            try
            {
                // Create a new workbook and get the first worksheet
                Workbook workbook = new Workbook();
                Worksheet sheet = workbook.Worksheets[0];

                // Populate a mixed range with numbers and leave some cells empty
                sheet.Cells["A1"].PutValue(10);   // numeric
                sheet.Cells["A2"].PutValue(20);   // numeric
                // B1 is left empty
                sheet.Cells["B2"].PutValue(30);   // numeric
                sheet.Cells["C1"].PutValue(40);   // numeric
                // C2 and C3 are left empty
                sheet.Cells["A3"].PutValue(5);    // numeric

                // Evaluate the SUM formula over the mixed range (empty cells are treated as zero by default)
                object sumResult = sheet.CalculateFormula("=SUM(A1:C3)");

                // Output the result
                Console.WriteLine("SUM(A1:C3) treating empty cells as zero = " + sumResult);

                // Save the workbook (optional)
                string outputPath = "SumResult.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {outputPath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred: " + ex.Message);
            }
        }
    }
}