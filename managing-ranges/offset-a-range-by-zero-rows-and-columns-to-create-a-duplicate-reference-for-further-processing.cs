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
                Worksheet worksheet = workbook.Worksheets[0];
                Cells cells = worksheet.Cells;

                // Create an original range (A1:B2) and populate it with sample data
                Aspose.Cells.Range originalRange = cells.CreateRange("A1", "B2");
                originalRange[0, 0].PutValue("R1C1");
                originalRange[0, 1].PutValue("R1C2");
                originalRange[1, 0].PutValue("R2C1");
                originalRange[1, 1].PutValue("R2C2");

                // Offset the range by zero rows and zero columns to obtain a duplicate reference
                Aspose.Cells.Range duplicateRange = originalRange.GetOffset(0, 0);

                // Modify a cell via the duplicate reference
                duplicateRange[0, 0].PutValue("Modified");

                // Save the workbook
                string outputPath = "OffsetZeroDemo.xlsx";
                workbook.Save(outputPath);
                Console.WriteLine($"Workbook saved to {Path.GetFullPath(outputPath)}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }
    }
}