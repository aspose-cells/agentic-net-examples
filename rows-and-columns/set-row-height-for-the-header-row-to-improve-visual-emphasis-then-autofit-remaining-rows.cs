using System;
using Aspose.Cells;

namespace AspiseCellsExamples
{
    // Author: Aspose.Cells .NET example
    class HeaderRowHeightDemo
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate sample data
            for (int row = 0; row < 10; row++)
            {
                for (int col = 0; col < 5; col++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Set a custom height for the header row (row index 0)
            Aspose.Cells.Range headerRange = sheet.Cells.CreateRange("1:1");
            headerRange.RowHeight = 30;

            // Auto‑fit all other rows while preserving the custom header height
            sheet.AutoFitRows(true);

            // Save the workbook
            workbook.Save("HeaderRowHeightDemo.xlsx");
        }
    }
}