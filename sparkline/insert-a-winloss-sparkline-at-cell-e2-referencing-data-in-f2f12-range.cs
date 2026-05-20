using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Program
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample win/loss data in the range F2:F12
            for (int i = 0; i < 11; i++)
            {
                // Alternating positive (win) and negative (loss) values
                sheet.Cells[1 + i, 5].PutValue(i % 2 == 0 ? 1 : -1);
            }

            // Define the location range for the sparkline (cells E2:E12)
            CellArea location = new CellArea
            {
                StartRow = 1,   // Row 2 (0‑based index)
                EndRow = 11,    // Row 12
                StartColumn = 4, // Column E (0‑based index)
                EndColumn = 4
            };

            // Add a Win/Loss sparkline (Stacked) that references F2:F12
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Stacked, "F2:F12", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // (Optional) Customize the sparkline group here, e.g., colors, line weight, etc.

            // Save the workbook
            workbook.Save("WinLossSparkline.xlsx", SaveFormat.Xlsx);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}