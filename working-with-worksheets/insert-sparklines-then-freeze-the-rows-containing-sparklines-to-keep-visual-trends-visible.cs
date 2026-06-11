using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeCellsSparklinesFreezeDemo
{
    class Program
    {
        static void Main()
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for 5 rows (A1:D5)
            for (int row = 0; row < 5; row++)
            {
                for (int col = 0; col < 4; col++)
                {
                    sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
                }
            }

            // Define the location range for sparklines (column 5, rows 0-4)
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 5,
                EndColumn = 5
            };

            // Add a sparkline group for the data range A1:D5
            int groupIdx = sheet.SparklineGroups.Add(SparklineType.Line, "A1:D5", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIdx];

            // Add a sparkline for each row in the group
            for (int r = 0; r < 5; r++)
            {
                // Data range for the current row
                string dataRange = $"A{r + 1}:D{r + 1}";
                // Add sparkline at column 5 (index 5) of the same row
                group.Sparklines.Add(dataRange, r, 5);
            }

            // Determine the lowest row index that contains a sparkline
            int maxSparklineRow = 0;
            foreach (Sparkline sp in group.Sparklines)
            {
                if (sp.Row > maxSparklineRow)
                    maxSparklineRow = sp.Row;
            }

            // Freeze rows up to and including the last sparkline row
            // FreezePanes(row, column, freezedRows, freezedColumns)
            // Use row = maxSparklineRow + 1 to set the split below the frozen area
            sheet.FreezePanes(maxSparklineRow + 1, 0, maxSparklineRow + 1, 0);

            // Save the workbook
            workbook.Save("SparklinesWithFrozenRows.xlsx");
        }
    }
}