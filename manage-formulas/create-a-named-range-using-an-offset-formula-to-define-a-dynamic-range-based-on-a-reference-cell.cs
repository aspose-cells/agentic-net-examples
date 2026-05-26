using System;
using System.IO;
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
                sheet.Name = "Sheet1";

                // Populate column A with sample data (A1:A10)
                for (int i = 0; i < 10; i++)
                {
                    sheet.Cells[i, 0].PutValue(i + 1);
                }

                // Add a named range that uses OFFSET to create a dynamic range.
                // The range starts at A1, has a height equal to the count of non‑empty cells in column A,
                // and a width of 1 column.
                int nameIndex = workbook.Worksheets.Names.Add("DynamicRange");
                Name dynamicName = workbook.Worksheets.Names[nameIndex];
                dynamicName.RefersTo = "=OFFSET(Sheet1!$A$1,0,0,COUNTA(Sheet1!$A:$A),1)";

                // Retrieve the actual range the name currently points to
                Aspose.Cells.Range dynamicRange = dynamicName.GetRange();
                Console.WriteLine("Dynamic range address: " + dynamicRange.Address);

                // Use the dynamic named range in a formula (e.g., sum)
                sheet.Cells["B1"].Formula = "=SUM(DynamicRange)";
                workbook.CalculateFormula();
                Console.WriteLine("Sum of dynamic range: " + sheet.Cells["B1"].Value);

                // Save the workbook
                string outputPath = "DynamicNamedRange.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                // Handle any runtime errors gracefully
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}