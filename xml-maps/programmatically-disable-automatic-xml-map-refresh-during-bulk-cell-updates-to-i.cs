using System;
using Aspose.Cells;

namespace AsposeCellsXmlMapOptimization
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook (lifecycle rule: create)
            Workbook wb = new Workbook();

            // Example: bulk update a large range of cells
            Worksheet sheet = wb.Worksheets[0];
            Cells cells = sheet.Cells;

            // Simulate bulk data insertion (e.g., 10,000 rows × 10 columns)
            for (int row = 0; row < 10000; row++)
            {
                for (int col = 0; col < 10; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // (Optional) Manually refresh XML maps if needed
            // wb.RefreshXmlMap(); // Uncomment if a specific refresh method exists

            // Save the workbook (lifecycle rule: save)
            wb.Save("BulkUpdate_NoXmlRefresh.xlsx");

            Console.WriteLine("Workbook saved successfully.");
        }
    }
}