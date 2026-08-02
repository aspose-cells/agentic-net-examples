using System;
using Aspose.Cells;
using Aspose.Cells.Charts;

class Sparkline3DExample
{
    static void Main()
    {
        try
        {
            // Create a new workbook and get the first worksheet
            Workbook workbook = new Workbook();
            Worksheet sheet = workbook.Worksheets[0];

            // Populate sample data for the sparkline (A1:A5)
            for (int i = 0; i < 5; i++)
            {
                sheet.Cells[i, 0].PutValue(i + 1); // Values 1‑5 in column A
            }

            // Define the location range for the sparkline group (G1:G5)
            // Column G = index 6 (0‑based), rows 0‑4
            CellArea location = new CellArea
            {
                StartRow = 0,
                EndRow = 4,
                StartColumn = 6,
                EndColumn = 6
            };

            // Add a sparkline group of type Line with the data range A1:A5
            int groupIndex = sheet.SparklineGroups.Add(SparklineType.Line, "A1:A5", false, location);
            SparklineGroup group = sheet.SparklineGroups[groupIndex];

            // Note: Sparkline groups do not support 3‑D formatting.

            // Save the workbook
            workbook.Save("Sparkline3DExample.xlsx");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error: {ex.Message}");
        }
    }
}