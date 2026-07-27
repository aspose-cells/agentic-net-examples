using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace SparklineBatchExample
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for rows 1‑20, columns A‑D
            for (int row = 0; row < 20; row++)          // zero‑based index
            {
                for (int col = 0; col < 4; col++)       // columns A‑D
                {
                    // Example data: (row + 1) * (col + 1)
                    sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
                }
            }

            // Define the location range where sparklines will be placed (column E, rows 1‑20)
            CellArea location = CellArea.CreateCellArea("E1", "E20");

            // Add a sparkline group:
            // - Type: Line
            // - Data range: A1:D20 (covers all rows)
            // - isVertical: false (sparklines are plotted by row)
            // - Location range: E1:E20 (one sparkline per row)
            int groupIndex = sheet.SparklineGroups.Add(
                SparklineType.Line,
                "A1:D20",
                false,
                location);

            // The group now contains 20 sparklines, one for each row.
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Optional: customize appearance (e.g., show high/low points)
            group.ShowHighPoint = true;
            group.ShowLowPoint = true;

            // Save the workbook
            workbook.Save("SparklinesRows1to20.xlsx");
        }
    }
}