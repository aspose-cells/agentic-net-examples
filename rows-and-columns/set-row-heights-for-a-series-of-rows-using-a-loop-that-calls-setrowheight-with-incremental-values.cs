using System;
using Aspose.Cells;

namespace AsposeCellsRowHeightDemo
{
    // Author: Example code for setting row heights with a loop
    class Program
    {
        static void Main()
        {
            // Create a new workbook
            Workbook workbook = new Workbook();

            // Access the first worksheet
            Worksheet sheet = workbook.Worksheets[0];
            Cells cells = sheet.Cells;

            // Set row heights using a loop.
            // Height starts at 15 points and increases by 5 points for each subsequent row.
            for (int row = 0; row < 10; row++)
            {
                double height = 15 + (row * 5); // Incremental height
                cells.SetRowHeight(row, height); // Calls Cells.SetRowHeight(int, double)
            }

            // Save the workbook
            workbook.Save("RowHeightsDemo.xlsx");
        }
    }
}