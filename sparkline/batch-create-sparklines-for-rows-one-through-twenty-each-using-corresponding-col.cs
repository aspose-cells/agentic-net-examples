using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        // Create a new workbook and get the first worksheet
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.Worksheets[0];

        // Populate sample data: rows 1‑20, columns A‑T (20 columns)
        // Values are just illustrative (row index * column index)
        for (int row = 0; row < 20; row++)          // zero‑based rows 0‑19 correspond to Excel rows 1‑20
        {
            for (int col = 0; col < 20; col++)      // zero‑based columns 0‑19 correspond to A‑T
            {
                worksheet.Cells[row, col].PutValue((row + 1) * (col + 1));
            }
        }

        // Add a sparkline group of type Line (no initial sparklines)
        int groupIndex = worksheet.SparklineGroups.Add(SparklineType.Line);
        SparklineGroup sparklineGroup = worksheet.SparklineGroups[groupIndex];

        // For each row, create a sparkline that uses the data from columns A‑T of that row
        // and place the sparkline in column U (zero‑based index 20)
        for (int row = 0; row < 20; row++)
        {
            string dataRange = $"A{row + 1}:T{row + 1}"; // e.g., "A1:T1", "A2:T2", ...
            int sparklineColumn = 20;                    // column U
            sparklineGroup.Sparklines.Add(dataRange, row, sparklineColumn);
        }

        // Save the workbook with the created sparklines
        workbook.Save("SparklinesRows1to20.xlsx");
    }
}