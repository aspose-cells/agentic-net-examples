using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class BatchSparklineDemo
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet sheet = workbook.Worksheets[0];

        // Populate sample data for 20 rows and 20 columns (A1:T20)
        for (int row = 0; row < 20; row++)
        {
            for (int col = 0; col < 20; col++)
            {
                // Example data: (row index + 1) * (col index + 1)
                sheet.Cells[row, col].PutValue((row + 1) * (col + 1));
            }
        }

        // Add a sparkline group of type Line (default orientation is horizontal)
        int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup group = sheet.SparklineGroups[groupIndex];

        // Column where sparklines will be placed (V column, zero‑based index 21)
        int sparklineColumn = 21;

        // Create a sparkline for each row (1‑through‑20) using the corresponding column data range
        // Row i uses data from column i (A‑T) across rows 1‑20 (vertical orientation)
        for (int i = 0; i < 20; i++)
        {
            // Build the data range string for column i (e.g., "A1:A20", "B1:B20", ...)
            string columnLetter = ((char)('A' + i)).ToString();
            string dataRange = $"{columnLetter}1:{columnLetter}20";

            // Add the sparkline at row i, column V
            group.Sparklines.Add(dataRange, i, sparklineColumn);
        }

        // Save the workbook with the created sparklines
        workbook.Save("BatchSparklines.xlsx");
    }
}