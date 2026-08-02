using System;
using Aspose.Cells;

namespace AsposeCellsColumnWidthDemo
{
    // Author: Aspose.Cells .NET example – auto‑fit then set exact pixel width
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Populate some sample data in columns A‑E (indices 0‑4)
            for (int col = 0; col < 5; col++)
            {
                for (int row = 0; row < 10; row++)
                {
                    cells[row, col].PutValue($"R{row + 1}C{col + 1}");
                }
            }

            // Auto‑fit the columns to their content (imprecise width)
            for (int col = 0; col < 5; col++)
            {
                sheet.AutoFitColumn(col);
            }

            // After auto‑fit, set exact pixel width for precise alignment
            // Desired width: 120 pixels for each column
            const int targetPixelWidth = 120;
            for (int col = 0; col < 5; col++)
            {
                cells.SetColumnWidthPixel(col, targetPixelWidth);
            }

            // Save the workbook
            workbook.Save("AutoFitThenSetPixelWidth.xlsx");
        }
    }
}